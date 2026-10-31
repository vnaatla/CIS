VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearchCLN 
   Caption         =   "Clinic Lookup"
   ClientHeight    =   5124
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9036
   LinkTopic       =   "Form1"
   ScaleHeight     =   5124
   ScaleWidth      =   9036
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lstClinic 
      Height          =   3495
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   8775
      _ExtentX        =   15473
      _ExtentY        =   6160
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Clinic Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Clinic Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Location"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Doctor Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   320
         Left            =   1800
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtClinicName 
         Height          =   320
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label Label1 
         Caption         =   "Clinic Name"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmSearchCLN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSearch_Click()
    If txtClinicName.Text = "" Then
        MsgBox "Please Enter Clinic Name to search", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        Call SearchClinic
    End If
End Sub
Private Sub Form_Load()
  Call ActionLogUpdate(12, "frmSearchCLN")
End Sub

Private Sub SearchClinic()
Const METHOD_NAME = "SearchClinic"
On Error GoTo debugs
Dim rst2 As New ADODB.Recordset
Dim ClnMaster As ListItem
Dim sql As String
    
    sql = "Select CLNCode, CLNName, COALESCE(CLNCity, '') AS Location, COALESCE(CLNDocname,'') AS Docname from CLN WHERE CLNName LIKE '%" & Trim(txtClinicName.Text) & "%'"
                  
    lstClinic.ListItems.Clear
    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set ClnMaster = lstClinic.ListItems.Add(, , rst2!CLNCode)
                ClnMaster.SubItems(1) = rst2!CLNName
                ClnMaster.SubItems(2) = rst2!location
                ClnMaster.SubItems(3) = rst2!DocName
                rst2.MoveNext
        Loop
    rst2.Close
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lstClinic_DblClick()
    If lstClinic.ListItems.Count <> 0 Then
        frmMaintenanceEntity.txtCAClnCode.Text = lstClinic.SelectedItem.Text
        frmMaintenanceEntity.lblClinicName.Caption = Trim(lstClinic.SelectedItem.SubItems(1))
        Unload Me
    End If
End Sub
