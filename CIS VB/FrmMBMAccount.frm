VERSION 5.00
Begin VB.Form FrmMBMAccount 
   Caption         =   "Member Account"
   ClientHeight    =   2040
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   3756
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   3756
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtRecAmt 
      Height          =   285
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox txtCLMNo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   0
      Width           =   2175
   End
   Begin VB.TextBox txtOSAmt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   3735
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Default         =   -1  'True
         Height          =   375
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton CmdOK 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2640
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Receipt/Payment"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Amount O/S"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Claim No"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "FrmMBMAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Const METHOD_NAME = "Form_Load"
On Error GoTo debugs

Call ActionLogUpdate(12, "FrmMBMAccount")


debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

