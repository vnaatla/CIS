VERSION 5.00
Begin VB.Form frmClaimsExcelSinChew 
   Caption         =   "Generate Claims  (Sin Chew)"
   ClientHeight    =   2475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      Height          =   1575
      Left            =   80
      TabIndex        =   0
      Top             =   240
      Width           =   7095
      Begin VB.TextBox txtBordxNoTo 
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
         Left            =   3240
         TabIndex        =   3
         Top             =   1170
         Width           =   1455
      End
      Begin VB.TextBox txtBordxNoFr 
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
         Left            =   1440
         TabIndex        =   2
         Top             =   1170
         Width           =   1455
      End
      Begin VB.ComboBox CboExportStatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   1920
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label Label26 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   12
         Top             =   1185
         Width           =   315
      End
      Begin VB.Label Label26 
         Caption         =   "Bordx No"
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
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   1180
         Width           =   915
      End
      Begin VB.Label Label25 
         Caption         =   "Export File Name"
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
         Left            =   120
         TabIndex        =   10
         Top             =   900
         Width           =   1245
      End
      Begin VB.Label Label28 
         Caption         =   "File on PC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   9
         Top             =   615
         Width           =   765
      End
      Begin VB.Label txtPath 
         BackColor       =   &H00FFFFC0&
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
         Left            =   1440
         TabIndex        =   8
         Top             =   240
         Width           =   3225
      End
      Begin VB.Label Label24 
         Caption         =   "File on Server"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   315
         Width           =   1125
      End
      Begin VB.Label txtLocalFileLocation 
         BackColor       =   &H00FFFFC0&
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
         Left            =   1440
         TabIndex        =   6
         Top             =   540
         Width           =   3225
      End
      Begin VB.Label Label1 
         Caption         =   "Export Status"
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
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label txtFilName 
         BackColor       =   &H00FFFFC0&
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
         Left            =   1440
         TabIndex        =   4
         Top             =   840
         Width           =   3225
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   80
      TabIndex        =   13
      Top             =   1800
      Width           =   7095
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6240
         TabIndex        =   15
         Top             =   240
         Width           =   705
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "&OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5520
         TabIndex        =   14
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label17 
         Caption         =   "Total Bordx Amt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   1
         Left            =   2160
         TabIndex        =   19
         Top             =   270
         Width           =   1200
      End
      Begin VB.Label lblTotalAmt 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3360
         TabIndex        =   18
         Top             =   240
         Width           =   1920
      End
      Begin VB.Label Label17 
         Caption         =   "Total Visit(s)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   270
         Width           =   1080
      End
      Begin VB.Label txtTotalVisits 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1080
         TabIndex        =   16
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Generate Claims  (Sin Chew)"
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
      TabIndex        =   20
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmClaimsExcelSinChew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
