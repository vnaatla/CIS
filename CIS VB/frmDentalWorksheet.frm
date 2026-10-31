VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDentalWorksheet 
   Caption         =   "Dental Worksheet"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11850
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Dental WorkSheet"
      TabPicture(0)   =   "frmDentalWorksheet.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fgrid"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdExit"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "Exit"
         Height          =   255
         Left            =   10560
         TabIndex        =   1
         Top             =   7920
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid Fgrid 
         Height          =   7455
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   13150
         _Version        =   393216
         Rows            =   20
         Cols            =   9
         FixedCols       =   2
         BackColorFixed  =   -2147483638
         AllowUserResizing=   1
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Dental Worksheet"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   225
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmDentalWorksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
