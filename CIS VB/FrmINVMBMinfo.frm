VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmINVMBMinfo 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INV Member Info"
   ClientHeight    =   3012
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6384
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3012
   ScaleWidth      =   6384
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdExit 
      Caption         =   "E&xit"
      Height          =   315
      Left            =   5400
      TabIndex        =   21
      Top             =   2640
      Width           =   855
   End
   Begin VB.Frame FrameMBMInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2355
      Left            =   0
      TabIndex        =   8
      Top             =   240
      Width           =   6375
      Begin VB.TextBox txtTabCoverID 
         Height          =   285
         Left            =   960
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxtPayor 
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
         Left            =   5640
         TabIndex        =   20
         Top             =   2040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtTabMBMNo 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox TxtCardType 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox TxtVisitLimit 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox TxtBalLimit 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtPolStatus 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtMBMStatus 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtPatientName 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox TxtPolNo 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txtDentStatus 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtTabMBMName 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtDentLmt 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtSpecLmt 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtMEPLmt 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtGrpCompany 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1200
         Width           =   3495
      End
      Begin MSMask.MaskEdBox txtExpDate 
         Height          =   315
         Left            =   3600
         TabIndex        =   7
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1926
         _ExtentY        =   550
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtdob 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox TxtInsType 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1920
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtEffDate 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2350
         _ExtentY        =   508
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtPatientIC 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1695
         Width           =   1335
      End
      Begin VB.TextBox txtPlanCode 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Grp Company"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Dent Lmt"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   4800
         TabIndex        =   44
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "MEP Lmt"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   4800
         TabIndex        =   43
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Spec Lmt"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   4800
         TabIndex        =   42
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Dent Stat."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   4800
         TabIndex        =   41
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label61 
         Caption         =   "Payor"
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
         Left            =   4800
         TabIndex        =   38
         Top             =   2040
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Bal. Lmt"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   4800
         TabIndex        =   37
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblPatientIC 
         Caption         =   "Patient IC"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1725
         Width           =   810
      End
      Begin VB.Label Label3 
         Caption         =   "MBM Status"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Patient Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Member Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Policy No"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Width           =   810
      End
      Begin VB.Label Label7 
         Caption         =   "Eff. Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1455
         Width           =   735
      End
      Begin VB.Label Label36 
         Caption         =   "Exp. Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   30
         Top             =   1500
         Width           =   855
      End
      Begin VB.Label Label37 
         Caption         =   "Patient DOB"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   29
         Top             =   1725
         Width           =   975
      End
      Begin VB.Label Label38 
         Caption         =   "Pol. Status"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   28
         Top             =   525
         Width           =   855
      End
      Begin VB.Label Label60 
         Caption         =   "Member No"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label62 
         Caption         =   "Ins Type"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   26
         Top             =   1965
         Width           =   735
      End
      Begin VB.Label Label64 
         Caption         =   "Plan Code"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1965
         Width           =   735
      End
      Begin VB.Label Label34 
         Caption         =   "Card Type"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   4800
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label54 
         Caption         =   "Visit Lmt"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   4800
         TabIndex        =   23
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   " Inv Membership Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "FrmINVMBMinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MBMNo As String
Dim CovID As String
Dim strPolNo As String
Dim strPatientName As String
Dim strMBMDataStatus As String
Dim strPOlStatus As String
Dim strEffdate As String
Dim strExpDate As String
Dim strClnPln As String
Dim strClnIns As String
Dim strPatientDOB As String
Dim strPayor As String
Dim strPatientID As String
Dim strPatientIC As String
Dim strBalLimit As String
Dim strVisitLimit As String
Dim strDentLmt As String
Dim strSpecLmt As String
Dim strMEPLmt As String
Dim strCardType As String
Dim strMBMCLNLTM As String
Dim Strrelationtype As String
Dim strMBMCancelDate  As String
Dim strCMBMCancelDate As String
Dim strCMBMCCLNExpDate As String
Dim strHLTCode As String
Dim SuppLmtType As String
Dim strVisitIND, strAnnualInd As String
Dim StrDentStatus As String
Dim strCoPay As String
Dim strSpecStatus As String
Dim strMEPStatus As String
Dim StrDentLmtStatus As String
Dim StrSpecLmtStatus As String
Dim strBNFStatus As String
Dim strDentExtLmt As String
Dim str1stSpecLmt As String
Dim strAccumAmt As String
Dim GRPCompany As String
Dim strDept As String
Public Source As Integer
Dim strGRPMCOStatus As String
Dim strPLNGrossStatus As String
Dim strMCStatus As String
Dim strMBMMCDaysBal As String
Dim strGRPPostageFees As String
Dim strGRPBordxFees As String
Dim strPLNMMPerc As String

Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub Form_Load()
Const METHOD_NAME = "Form_Load"
On Error GoTo debugs

Call ActionLogUpdate(12, "FrmINVMBMinfo")


debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub


Private Sub Form_Activate()
Const METHOD_NAME = "Form_Activate"
On Error GoTo debugs
strMBMDataStatus = ""
strPatientName = ""
strPolNo = ""
strEffdate = ""
strExpDate = ""
strPatientIC = ""
strPatientDOB = ""
strCardType = ""
strVisitLimit = ""
strBalLimit = ""
strDentLmt = ""
strSpecLmt = ""
strMEPLmt = ""
strPOlStatus = ""
strClnPln = ""
strClnIns = ""
SuppLmtType = ""
    
    'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        If Source = 1 Then  'Clinical Invoice - Clinic
            txtTabMBMNo = frmInvoice.txtTabMBMNo
            txtTabCoverID = frmInvoice.txtTabCoverID
            MBMNo = txtTabMBMNo
            CovID = txtTabCoverID
        ElseIf Source = 2 Then 'Clinical Invoice - Member
            txtTabMBMNo = frmInvoiceMBM.txtTabMBMNo
            txtTabCoverID = frmInvoiceMBM.txtTabCoverID
            MBMNo = txtTabMBMNo
            CovID = txtTabCoverID
        ElseIf Source = 3 Then 'Clinical Invoice - Clinic & Member Adjustment
            txtTabMBMNo = frmInvoiceAdj.txtTabMBMNo
            txtTabCoverID = frmInvoiceAdj.txtTabCoverID
            MBMNo = txtTabMBMNo
            CovID = txtTabCoverID
        ElseIf Source = 4 Then 'Clinical Invoice - Clinic & Member Enquiry
            txtTabMBMNo = FrmInvoiceEnq.txtTabMBMNo
            txtTabCoverID = FrmInvoiceEnq.txtTabCoverID
            MBMNo = txtTabMBMNo
            CovID = txtTabCoverID
    
        ElseIf Source = 5 Then 'Clinical Invoice - Clinic & Member Enquiry
            txtTabMBMNo = frmInvoiceAdjAftBordx.txtTabMBMNo
            txtTabCoverID = frmInvoiceAdjAftBordx.txtTabCoverID
            MBMNo = txtTabMBMNo
            CovID = txtTabCoverID
     '   ElseIf Source = 5 Then 'Dental Invoice - Clinic
     '       txtTabMBMNo = FrmInvDentalCln.txtTabMBMNo
     '       txtTabCoverID = FrmInvDentalCln.txtTabCoverID
     '       MBMNo = txtTabMBMNo
     '       CovID = txtTabCoverID
    '    ElseIf Source = 6 Then 'Dental Invoice - Member
    '        txtTabMBMNo = FrmInvDentalMBM.txtTabMBMNo
    '        txtTabCoverID = FrmInvDentalMBM.txtTabCoverID
    '        MBMNo = txtTabMBMNo
    '        CovID = txtTabCoverID
        'ElseIf Source = 7 Then 'Dental Invoice - Clinic & Member
        '    txtTabMBMNo = FrmInvDentalAdj.txtTabMBMNo
        '    txtTabCoverID = FrmInvDentalAdj.txtTabCoverID
        '    MBMNo = txtTabMBMNo
        '    CovID = txtTabCoverID
            
        End If
        Call MemberSuppInfo(MBMNo, CovID, strBalLimit, strVisitLimit, strAccumAmt, _
                            StrDentStatus, strCoPay, strPayor, strHLTCode, strClnPln, _
                            strClnIns, strEffdate, strExpDate, strPolNo, strMBMDataStatus, _
                            strPOlStatus, strPatientName, strPatientDOB, strPatientIC, _
                            strMBMCLNLTM, strCardType, strVisitIND, strAnnualInd, strSpecStatus, _
                            strMEPStatus, strDentLmt, strSpecLmt, strMEPLmt, SuppLmtType, _
                            StrDentLmtStatus, StrSpecLmtStatus, strBNFStatus, strDentExtLmt, str1stSpecLmt, _
                            strGRPMCOStatus, strPLNGrossStatus, strDept, strMCStatus, strMBMMCDaysBal, strGRPPostageFees, strGRPBordxFees, _
                            strPLNMMPerc, Strrelationtype, strMBMCancelDate, strCMBMCancelDate, strCMBMCCLNExpDate)
        Call PlnDesc(strClnPln, GRPCompany)
    '    txtGrpCompany = GRPCompany
        txtTabMBMName = strPatientName
        txtMBMStatus = strMBMDataStatus
        txtPatientName = strPatientName
        TxtPolNo = strPolNo
        txtEffDate = strEffdate
        txtExpDate = strExpDate
        txtPatientIC = strPatientIC
        txtdob = strPatientDOB
        TxtCardType = strCardType
        TxtVisitLimit = strVisitLimit
        TxtBalLimit = strBalLimit
        txtPolStatus = strPOlStatus
        txtDentStatus = StrDentStatus
        txtPlanCode = strClnPln
        TxtInsType = strClnIns
        txtDentLmt = strDentLmt
        txtSpecLmt = strSpecLmt
        txtMEPLmt = strMEPLmt
        
    'MediConnCISDB.Close
    'Set MediConnCISDB = Nothing
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
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

Private Sub PlnDesc(PCode, GRPCompany As String)
Const METHOD_NAME = "PlnDesc"
On Error GoTo debugs
'Dim PLNRec As New ADODB.Recordset
'Dim PlnSql As String

'    PlnSql = "SELECT PLNCode, GRPCompany From PLN WHERE PLNCode='" & Trim(PCode) & "'"
'    PLNRec.Open PlnSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
'    Do While Not PLNRec.EOF
'        If Trim(PLNRec!GRPCompany) <> "" Then
'            GRPCompany = PLNRec!GRPCompany
 '       End If
'        PLNRec.MoveNext
 '   Loop
 '   PLNRec.Close
'    Set PLNRec = Nothing
Dim MBMO As New ADODB.Recordset

        MBMO.Open "SELECT MBMGroupCompany, MBMDepartment from MBMOthers where MBMNumber='" & Trim(MBMNo) & "'", medixmasterconn, adOpenForwardOnly, adLockReadOnly
        If MBMO.EOF = False Then
'            txtDept = IIf(Len(MBMO!MBMDepartment), MBMO!MBMDepartment, "")
            txtGrpCompany = IIf(Len(MBMO!MBMGroupCompany), MBMO!MBMGroupCompany, "")
        End If
        MBMO.Close
        Set MBMO = Nothing

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub





