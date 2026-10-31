VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmINVBatProcessing 
   Caption         =   "INvoice Batch Processing"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8235
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame FrameInvoice 
      Enabled         =   0   'False
      Height          =   4095
      Left            =   80
      TabIndex        =   75
      Top             =   960
      Width           =   11775
      Begin VB.Frame FrameCharges 
         Caption         =   "Charges - Cover"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2055
         Left            =   80
         TabIndex        =   84
         Top             =   1920
         Width           =   6090
         Begin VB.CommandButton CmdLab 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   220
            Left            =   1845
            TabIndex        =   88
            Top             =   1220
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.CommandButton cmdProcedures 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   220
            Left            =   1845
            TabIndex        =   87
            Top             =   980
            Width           =   240
         End
         Begin VB.CommandButton CmdMedication 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   220
            Left            =   1845
            TabIndex        =   86
            Top             =   740
            Width           =   240
         End
         Begin VB.TextBox TxtIndex 
            Height          =   195
            Left            =   960
            TabIndex        =   85
            Top             =   960
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox A1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   24
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox A2 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   25
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox A3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   26
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox A4 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   27
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox A5 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   28
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox A6 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3555
            MaxLength       =   6
            TabIndex        =   29
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox A7 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3555
            MaxLength       =   6
            TabIndex        =   30
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox A8 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3555
            MaxLength       =   6
            TabIndex        =   31
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox A9 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3555
            MaxLength       =   6
            TabIndex        =   32
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox A10 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3555
            MaxLength       =   6
            TabIndex        =   33
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox A11 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5355
            MaxLength       =   6
            TabIndex        =   34
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox A12 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5355
            MaxLength       =   6
            TabIndex        =   35
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox A13 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5355
            MaxLength       =   6
            TabIndex        =   36
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox A14 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5355
            MaxLength       =   6
            TabIndex        =   37
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox A15 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5355
            MaxLength       =   6
            TabIndex        =   38
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtOtherDesc1 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            MaxLength       =   100
            TabIndex        =   40
            Top             =   1680
            Width           =   2970
         End
         Begin VB.TextBox A16 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5355
            MaxLength       =   6
            TabIndex        =   39
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox A17 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5355
            MaxLength       =   6
            TabIndex        =   41
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label Label33 
            Caption         =   "T&&S / STO"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4320
            TabIndex        =   106
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label32 
            Caption         =   "I&&D"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4320
            TabIndex        =   105
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label31 
            Caption         =   "Vision"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4320
            TabIndex        =   104
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label30 
            Caption         =   "PRV Checkup"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4320
            TabIndex        =   103
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label29 
            Caption         =   "Maternity"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4320
            TabIndex        =   102
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label28 
            Caption         =   "Physio Theraphy"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2235
            TabIndex        =   101
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label27 
            Caption         =   "Traditional Med"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2235
            TabIndex        =   100
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label26 
            Caption         =   "Chinese Medicine"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2235
            TabIndex        =   99
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label25 
            Caption         =   "Dental"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2235
            TabIndex        =   98
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label23 
            Caption         =   "X-Ray / MRI"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2235
            TabIndex        =   97
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label22 
            Caption         =   "Lab"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   96
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label21 
            Caption         =   "Procedures"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   95
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label20 
            Caption         =   "Medication"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label19 
            Caption         =   "SP Cons Fee"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label18 
            Caption         =   "GP Cons Fee"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label13 
            Caption         =   "Others1 Desc"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label53 
            Caption         =   "Injection"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4320
            TabIndex        =   90
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4320
            TabIndex        =   89
            Top             =   1680
            Width           =   735
         End
      End
      Begin VB.Frame FrameInvDet 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1700
         Left            =   120
         TabIndex        =   111
         Top             =   100
         Width           =   11625
         Begin VB.OptionButton OptAccute 
            Caption         =   "Accute"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   90
            TabIndex        =   5
            Top             =   120
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OptChronic 
            Caption         =   "Chronic"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            TabIndex        =   6
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton OptMultiple 
            Caption         =   "Multiple"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2520
            TabIndex        =   7
            Top             =   120
            Width           =   975
         End
         Begin VB.Frame FrameReason 
            Caption         =   "Reason For Seeking Treatment"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1335
            Left            =   8680
            TabIndex        =   115
            Top             =   260
            Width           =   2870
            Begin VB.ComboBox CboReason1 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   200
               TabIndex        =   21
               Tag             =   "0"
               Top             =   240
               Width           =   2640
            End
            Begin VB.ComboBox CboReason2 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   200
               TabIndex        =   22
               Top             =   520
               Width           =   2640
            End
            Begin VB.ComboBox CboReason3 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   200
               TabIndex        =   23
               Top             =   800
               Width           =   2640
            End
            Begin VB.Label Label52 
               Caption         =   "3"
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
               Left            =   90
               TabIndex        =   118
               Top             =   840
               Width           =   255
            End
            Begin VB.Label Label51 
               Caption         =   "2"
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
               Left            =   90
               TabIndex        =   117
               Top             =   525
               Width           =   255
            End
            Begin VB.Label Label12 
               Caption         =   "1"
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
               Left            =   90
               TabIndex        =   116
               Top             =   240
               Width           =   135
            End
         End
         Begin VB.Frame FrameCLMCat 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1095
            Left            =   5880
            TabIndex        =   112
            Top             =   480
            Width           =   2775
            Begin VB.ComboBox CboNatureClm 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1125
               TabIndex        =   19
               Top             =   195
               Width           =   1575
            End
            Begin VB.TextBox TxtRemarks 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   1125
               MaxLength       =   100
               TabIndex        =   20
               Top             =   480
               Width           =   1575
            End
            Begin VB.Label Label41 
               Caption         =   "Remarks"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   114
               Top             =   480
               Width           =   615
            End
            Begin VB.Label Label39 
               Caption         =   "Nat. of CLM"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   113
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.TextBox txtMC 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   7500
            MaxLength       =   4
            TabIndex        =   18
            Top             =   255
            Width           =   375
         End
         Begin MSMask.MaskEdBox txtDateVisit 
            Height          =   315
            Left            =   4830
            TabIndex        =   14
            Top             =   420
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox Cboclinic 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   660
            TabIndex        =   8
            Top             =   400
            Width           =   3130
         End
         Begin VB.TextBox TxtInvPayType 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2640
            TabIndex        =   10
            Text            =   "Reimbursement"
            Top             =   700
            Width           =   1150
         End
         Begin VB.ComboBox CboPAYto 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   660
            TabIndex        =   9
            Text            =   "C - Clinic"
            Top             =   700
            Width           =   1095
         End
         Begin VB.TextBox txtInvoiceNo 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4830
            TabIndex        =   15
            Top             =   680
            Width           =   975
         End
         Begin MSMask.MaskEdBox txtInvoiceDate 
            Height          =   315
            Left            =   4830
            TabIndex        =   16
            Top             =   960
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtVisitTime 
            Height          =   315
            Left            =   4830
            TabIndex        =   17
            Top             =   1230
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtPayee 
            Height          =   330
            Left            =   660
            TabIndex        =   11
            Top             =   990
            Width           =   3130
         End
         Begin VB.ComboBox CboPayee 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   660
            TabIndex        =   12
            Top             =   990
            Visible         =   0   'False
            Width           =   3130
         End
         Begin VB.TextBox txtDoctor 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   660
            MaxLength       =   50
            TabIndex        =   13
            Top             =   1270
            Width           =   3130
         End
         Begin VB.Label Label50 
            Caption         =   "Doctor"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   105
            TabIndex        =   131
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Lbl45days 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "45 Days Late"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   210
            Left            =   4800
            TabIndex        =   130
            Top             =   180
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label11 
            Caption         =   "Cash Type"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   129
            Top             =   750
            Width           =   855
         End
         Begin VB.Label Label17 
            Caption         =   "Payee"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   105
            TabIndex        =   128
            Top             =   1035
            Width           =   495
         End
         Begin VB.Label Label10 
            Caption         =   "Pay To"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   105
            TabIndex        =   127
            Top             =   735
            Width           =   615
         End
         Begin VB.Label Label9 
            Caption         =   "Invoice Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3880
            TabIndex        =   126
            Top             =   990
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Invoice No"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3880
            TabIndex        =   125
            Top             =   735
            Width           =   975
         End
         Begin VB.Label Label14 
            Caption         =   "Clinic"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   124
            Top             =   435
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Submit Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3880
            TabIndex        =   123
            Top             =   465
            Width           =   1215
         End
         Begin VB.Label LblExpDate 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "Expired"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   3840
            TabIndex        =   122
            Top             =   180
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label66 
            Caption         =   "Visit Time"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3880
            TabIndex        =   121
            Top             =   1275
            Width           =   975
         End
         Begin VB.Label Label16 
            Caption         =   "Days"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8040
            TabIndex        =   120
            Top             =   280
            Width           =   345
         End
         Begin VB.Label Label15 
            Caption         =   "Medical Cert.(MC)"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6000
            TabIndex        =   119
            Top             =   285
            Width           =   1335
         End
      End
      Begin VB.Frame FrameRepudiate 
         Caption         =   "Repudiation Reasons"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   8480
         TabIndex        =   107
         Top             =   1920
         Width           =   3225
         Begin VB.ComboBox CboRepudiate1 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   260
            TabIndex        =   49
            Top             =   240
            Width           =   2880
         End
         Begin VB.ComboBox CboRepudiate2 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   240
            TabIndex        =   50
            Top             =   540
            Width           =   2880
         End
         Begin VB.ComboBox CboRepudiate3 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   260
            TabIndex        =   51
            Top             =   760
            Width           =   2880
         End
         Begin VB.Label Label48 
            Caption         =   "3"
            Height          =   255
            Left            =   120
            TabIndex        =   110
            Top             =   780
            Width           =   255
         End
         Begin VB.Label Label47 
            Caption         =   "2"
            Height          =   180
            Left            =   120
            TabIndex        =   109
            Top             =   480
            Width           =   255
         End
         Begin VB.Label Label43 
            Caption         =   "1"
            Height          =   255
            Left            =   120
            TabIndex        =   108
            Top             =   255
            Width           =   255
         End
      End
      Begin VB.Frame FrameNotCover 
         Caption         =   "Not Cover"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   6210
         TabIndex        =   76
         Top             =   1920
         Width           =   2210
         Begin VB.TextBox B1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            MaxLength       =   8
            TabIndex        =   42
            Top             =   240
            Width           =   680
         End
         Begin VB.TextBox B2 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            MaxLength       =   8
            TabIndex        =   43
            Top             =   480
            Width           =   680
         End
         Begin VB.TextBox B3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            MaxLength       =   8
            TabIndex        =   44
            Top             =   720
            Width           =   680
         End
         Begin VB.TextBox B4 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            MaxLength       =   8
            TabIndex        =   45
            Top             =   960
            Width           =   680
         End
         Begin VB.TextBox B5 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            MaxLength       =   8
            TabIndex        =   46
            Top             =   1200
            Width           =   680
         End
         Begin VB.TextBox B6 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            MaxLength       =   8
            TabIndex        =   47
            Top             =   1440
            Width           =   680
         End
         Begin VB.TextBox B7 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            MaxLength       =   8
            TabIndex        =   48
            Top             =   1680
            Width           =   680
         End
         Begin VB.Label Label24 
            Caption         =   "SP Cons Fee"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label45 
            Caption         =   "Medication"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   82
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label46 
            Caption         =   "Physio Theraphy"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   81
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label55 
            Caption         =   "Maternity"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label56 
            Caption         =   "PRV Checkup"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   79
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label57 
            Caption         =   "Vision"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label59 
            Caption         =   "Dental"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.Frame FrameMBMInfo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1155
         Left            =   240
         TabIndex        =   132
         Top             =   1800
         Visible         =   0   'False
         Width           =   11610
         Begin VB.TextBox TxtClnContract 
            Height          =   285
            Left            =   10920
            TabIndex        =   169
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox TxtInsType 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9240
            TabIndex        =   145
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtPlanCode 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9240
            TabIndex        =   144
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txtdob 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7320
            TabIndex        =   143
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtPatientIC 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5280
            TabIndex        =   142
            Top             =   740
            Width           =   975
         End
         Begin VB.TextBox txtPatientName 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            TabIndex        =   141
            Top             =   720
            Width           =   3135
         End
         Begin VB.TextBox txtTabMBMName 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            TabIndex        =   140
            Top             =   480
            Width           =   3135
         End
         Begin VB.TextBox txtTabMBMNo 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            TabIndex        =   139
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtMBMStatus 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3600
            TabIndex        =   138
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox TxtPolNo 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5280
            TabIndex        =   135
            Top             =   240
            Width           =   3015
         End
         Begin VB.TextBox txtPolStatus 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9240
            TabIndex        =   134
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox TxtPayor 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4320
            TabIndex        =   133
            Top             =   720
            Visible         =   0   'False
            Width           =   150
         End
         Begin MSMask.MaskEdBox txtExpDate 
            Height          =   315
            Left            =   7320
            TabIndex        =   136
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtEffDate 
            Height          =   285
            Left            =   5280
            TabIndex        =   137
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtBalLimit 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   10920
            TabIndex        =   146
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox TxtVisitLimit 
            Height          =   285
            Left            =   10920
            TabIndex        =   167
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label40 
            Caption         =   "Cln Contract"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   255
            Left            =   9840
            TabIndex        =   168
            Top             =   300
            Width           =   975
         End
         Begin VB.Label Label34 
            Caption         =   "Visit Limit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   255
            Left            =   9840
            TabIndex        =   166
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label5 
            BackColor       =   &H80000013&
            Caption         =   "Balance Limit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   9840
            TabIndex        =   160
            Top             =   540
            Width           =   1095
         End
         Begin VB.Label lblPatientIC 
            Caption         =   "Patient IC"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4520
            TabIndex        =   159
            Top             =   760
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "MBM Status"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2715
            TabIndex        =   158
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Patient Name"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   157
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Member Name"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   156
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Policy No"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4520
            TabIndex        =   155
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Eff. Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4520
            TabIndex        =   154
            Top             =   495
            Width           =   1095
         End
         Begin VB.Label Label36 
            Caption         =   "Exp. Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6360
            TabIndex        =   153
            Top             =   495
            Width           =   855
         End
         Begin VB.Label Label37 
            Caption         =   "Patient DOB"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6360
            TabIndex        =   152
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label38 
            Caption         =   "Pol. Status"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8400
            TabIndex        =   151
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label60 
            Caption         =   "Member No"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   150
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label61 
            Caption         =   "Payor"
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
            Left            =   7440
            TabIndex        =   149
            Top             =   760
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label Label62 
            Caption         =   "InsType"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8400
            TabIndex        =   148
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label64 
            Caption         =   "Plan Code"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8400
            TabIndex        =   147
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.Label lblTotAmt 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Total Invoice Amount Excess"
         ForeColor       =   &H00004080&
         Height          =   210
         Left            =   8760
         TabIndex        =   162
         Top             =   3600
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.Label lblMedicine 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Medication Amount Excess"
         ForeColor       =   &H00004080&
         Height          =   210
         Left            =   8760
         TabIndex        =   161
         Top             =   3360
         Visible         =   0   'False
         Width           =   2100
      End
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   80
      TabIndex        =   73
      Top             =   360
      Width           =   4095
      Begin VB.TextBox txtMBMNo 
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
         Left            =   1320
         TabIndex        =   0
         Top             =   180
         Width           =   2055
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "&Go"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         TabIndex        =   1
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label54 
         Caption         =   "Member No."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   195
         Width           =   1095
      End
   End
   Begin VB.Frame Frame7 
      Height          =   555
      Left            =   80
      TabIndex        =   66
      Top             =   7620
      Width           =   11775
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print Listing"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4320
         TabIndex        =   53
         Top             =   160
         Width           =   1095
      End
      Begin VB.CommandButton CmdMBMInfo 
         Caption         =   "&MBM Info"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5400
         TabIndex        =   54
         Top             =   160
         Width           =   975
      End
      Begin VB.CommandButton CMdChkItem 
         Caption         =   "C&hk Items"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4800
         TabIndex        =   52
         Top             =   160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "E&dit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7920
         TabIndex        =   68
         Top             =   160
         Width           =   735
      End
      Begin VB.CommandButton CmdClaerInv 
         Caption         =   "Clear &Inv"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7080
         TabIndex        =   56
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdClearAll 
         Caption         =   "&Clear All"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8640
         TabIndex        =   58
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10920
         TabIndex        =   61
         Top             =   160
         Width           =   735
      End
      Begin VB.CommandButton CmdAddRec 
         Caption         =   "&Add"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6360
         TabIndex        =   55
         Top             =   160
         Width           =   735
      End
      Begin VB.CommandButton CmdRemove 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9480
         TabIndex        =   59
         Top             =   160
         Width           =   735
      End
      Begin VB.CommandButton CmdVerify 
         Caption         =   "&Verify"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10200
         TabIndex        =   67
         Top             =   160
         Width           =   735
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10200
         TabIndex        =   60
         Top             =   160
         Width           =   735
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7920
         TabIndex        =   57
         Top             =   160
         Width           =   735
      End
      Begin VB.Label lblWs 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   72
         Top             =   200
         Width           =   495
      End
      Begin VB.Label Label63 
         Caption         =   "WS?"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   71
         Top             =   200
         Width           =   495
      End
      Begin VB.Label lblBordx 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   70
         Top             =   200
         Width           =   375
      End
      Begin VB.Label Label58 
         Caption         =   "Bordx?"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   69
         Top             =   200
         Width           =   615
      End
   End
   Begin VB.Frame FrameCase 
      Height          =   555
      Left            =   4200
      TabIndex        =   62
      Top             =   360
      Width           =   7575
      Begin VB.ComboBox cboClaimability 
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
         Left            =   3960
         TabIndex        =   3
         Text            =   "A - Accepted"
         Top             =   180
         Width           =   1335
      End
      Begin VB.ComboBox cboStatus 
         Enabled         =   0   'False
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
         Left            =   1560
         TabIndex        =   2
         Text            =   "O - Open"
         Top             =   180
         Width           =   1095
      End
      Begin VB.ComboBox cboPending 
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
         Left            =   6480
         TabIndex        =   4
         Text            =   "N - No"
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label44 
         Caption         =   "Claimability"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   65
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label42 
         Caption         =   "Case Status"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   64
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label49 
         Caption         =   "Pending"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   63
         Top             =   240
         Width           =   660
      End
   End
   Begin MSComctlLib.ListView lstInvoice 
      Height          =   2505
      Left            =   80
      TabIndex        =   163
      Top             =   5160
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   4419
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   71
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Index"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Case No."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Member No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cov ID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Patient Name"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "WS Status"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Bordxed"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Inv No"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Inv Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Submission Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "NatureClm"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Clinic"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Claim Cat"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "Amount"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Pay To"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Payee"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Doctor"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "PreparedBy"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   18
         Text            =   "GP Amt"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   19
         Text            =   "SP Amt"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   20
         Text            =   "Medicine"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   21
         Text            =   "Procedures"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   22
         Text            =   "Lab"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   23
         Text            =   "X-Ray"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   24
         Text            =   "Dental"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   25
         Text            =   "Chinese Med"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   26
         Text            =   "Traditional Med"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   27
         Text            =   "Physio Theraphy"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   28
         Text            =   "Maternity"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   29
         Text            =   "PRV Checkup"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   30
         Text            =   "Vision"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   31
         Text            =   "ID"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   32
         Text            =   "SSTO"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   33
         Text            =   "Others1 Desc"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   34
         Text            =   "Others2 Desc"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   35
         Text            =   "Others1"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   36
         Text            =   "Others2"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   37
         Text            =   "NC GP Fees"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   38
         Text            =   "NC Medicine"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   39
         Text            =   "NC Dental"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   40
         Text            =   "NC Physio"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   41
         Text            =   "NC Maternity"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   42
         Text            =   "NC PRV"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   43
         Text            =   "NC Vision"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   44
         Text            =   "Claimability"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   45
         Text            =   "Pending"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   46
         Text            =   "Reason1"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   47
         Text            =   "Reason2"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(49) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   48
         Text            =   "Reason3"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(50) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   49
         Text            =   "MC"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(51) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   50
         Text            =   "Remarks"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(52) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   51
         Text            =   "RepCode1"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(53) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   52
         Text            =   "RepCode2"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(54) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   53
         Text            =   "RepCode3"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(55) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   54
         Text            =   "Medication1"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(56) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   55
         Text            =   "Medication2"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(57) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   56
         Text            =   "Medication3"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(58) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   57
         Text            =   "Proc1"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(59) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   58
         Text            =   "Proc2"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(60) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   59
         Text            =   "Proc3"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(61) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   60
         Text            =   "Proc4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(62) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   61
         Text            =   "Proc5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(63) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   62
         Text            =   "Proc6"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(64) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   63
         Text            =   "Proc7"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(65) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   64
         Text            =   "InvPayType"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(66) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   65
         Text            =   "Payor"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(67) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   66
         Text            =   "ChkItemAmt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(68) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   67
         Text            =   "Visit Time"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(69) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   68
         Text            =   "Bal Lmt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(70) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   69
         Text            =   "Visit Lmt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(71) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   70
         Text            =   "ClnContract"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lstErr 
      Height          =   255
      Left            =   0
      TabIndex        =   164
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "MBM No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cov"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "MBM Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Policy No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Patient Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Patient IC"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Eff Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Exp. Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Bal. Limit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Err Msg"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label35 
      BackColor       =   &H00404040&
      Caption         =   " Invoice Batch Processing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   0
      TabIndex        =   165
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmINVBatProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim bExit As Boolean
Dim intIndex As String
Dim EditFlag As Boolean
Dim ScanCode As String
Dim AA As Integer
Dim strBillChkItem, strMEdChkItem As String
Dim strBillChkAmt, strMEdChkAmt, IntInvAmt As Integer
Dim CASENumber, NewCaseNo, MembershipNo, CovPersonID As String
Dim strPolNo, strMBMName, strPatientName, strMBMDataStatus, strPOlStatus As String
Dim strEffdate, strExpDate, strPlanCode, strInsType, strPatientDOB As String
Dim strPayor, strPatientID, strPatientIC, strBalLimit, strVisitLimit, strCLNContract As String
Dim MBM As New ADODB.Recordset
Dim lstErrRec As ListItem
Dim VMBMNo, VName, VPolNo, VIc, VPlan, VExpDate, VEffDate, VCoverID, VDiagCode, VDiagIND As String
Dim VBalLmt As Variant

Private Sub A1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A3_GotFocus()
    lblMedicine.Visible = False
End Sub

Private Sub A3_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A4_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A5_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A6_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A7_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A8_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A9_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A10_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A11_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A12_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A13_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A14_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A15_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A16_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A17_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub B1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub B2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub B3_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub B4_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub B5_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub B6_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub B7_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub cboClaimability_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboClaimability.Text, cboClaimability.SelStart) + Chr(KeyAscii) + Mid(cboClaimability.Text, cboClaimability.SelStart + cboClaimability.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboClaimability.ListCount - 1
       
        If NewText = LCase(Left(cboClaimability.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = cboClaimability.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            cboClaimability.Text = ValidValue
            cboClaimability.SelStart = 0
            cboClaimability.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub cboClaimability_LostFocus()
    If Mid(cboClaimability, 1, 1) = "R" Then
        FrameRepudiate.Enabled = True
    End If
End Sub

Private Sub Cboclinic_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(Cboclinic.Text, Cboclinic.SelStart) + Chr(KeyAscii) + Mid(Cboclinic.Text, Cboclinic.SelStart + Cboclinic.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To Cboclinic.ListCount - 1
       
        If NewText = LCase(Left(Cboclinic.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = Cboclinic.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            Cboclinic.Text = ValidValue
            Cboclinic.SelStart = 0
            Cboclinic.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CboNatureClm_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboNatureClm.Text, CboNatureClm.SelStart) + Chr(KeyAscii) + Mid(CboNatureClm.Text, CboNatureClm.SelStart + CboNatureClm.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboNatureClm.ListCount - 1
       
        If NewText = LCase(Left(CboNatureClm.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboNatureClm.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboNatureClm.Text = ValidValue
            CboNatureClm.SelStart = 0
            CboNatureClm.SelLength = Len(ValidValue)
        End If
    End If
End Sub


Private Sub CboPAYto_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPAYto.Text, CboPAYto.SelStart) + Chr(KeyAscii) + Mid(CboPAYto.Text, CboPAYto.SelStart + CboPAYto.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboPAYto.ListCount - 1
       
        If NewText = LCase(Left(CboPAYto.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboPAYto.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboPAYto.Text = ValidValue
            CboPAYto.SelStart = 0
            CboPAYto.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub cboPending_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPending.Text, cboPending.SelStart) + Chr(KeyAscii) + Mid(cboPending.Text, cboPending.SelStart + cboPending.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboPending.ListCount - 1
       
        If NewText = LCase(Left(cboPending.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = cboPending.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            cboPending.Text = ValidValue
            cboPending.SelStart = 0
            cboPending.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CboReason1_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboReason1.Text, CboReason1.SelStart) + Chr(KeyAscii) + Mid(CboReason1.Text, CboReason1.SelStart + CboReason1.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboReason1.ListCount - 1
       
        If NewText = LCase(Left(CboReason1.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboReason1.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboReason1.Text = ValidValue
            CboReason1.SelStart = 0
            CboReason1.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CboReason2_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboReason2.Text, CboReason2.SelStart) + Chr(KeyAscii) + Mid(CboReason2.Text, CboReason2.SelStart + CboReason2.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboReason2.ListCount - 1
       
        If NewText = LCase(Left(CboReason2.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboReason2.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboReason2.Text = ValidValue
            CboReason2.SelStart = 0
            CboReason2.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CboReason3_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboReason3.Text, CboReason3.SelStart) + Chr(KeyAscii) + Mid(CboReason3.Text, CboReason3.SelStart + CboReason3.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboReason3.ListCount - 1
       
        If NewText = LCase(Left(CboReason3.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboReason3.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboReason3.Text = ValidValue
            CboReason3.SelStart = 0
            CboReason3.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CboRepudiate1_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboRepudiate1.Text, CboRepudiate1.SelStart) + Chr(KeyAscii) + Mid(CboRepudiate1.Text, CboRepudiate1.SelStart + CboRepudiate1.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboRepudiate1.ListCount - 1
       
        If NewText = LCase(Left(CboRepudiate1.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboRepudiate1.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboRepudiate1.Text = ValidValue
            CboRepudiate1.SelStart = 0
            CboRepudiate1.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CboRepudiate2_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboRepudiate2.Text, CboRepudiate2.SelStart) + Chr(KeyAscii) + Mid(CboRepudiate2.Text, CboRepudiate2.SelStart + CboRepudiate2.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboRepudiate2.ListCount - 1
       
        If NewText = LCase(Left(CboRepudiate2.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboRepudiate2.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboRepudiate2.Text = ValidValue
            CboRepudiate2.SelStart = 0
            CboRepudiate2.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CboRepudiate3_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboRepudiate3.Text, CboRepudiate3.SelStart) + Chr(KeyAscii) + Mid(CboRepudiate2.Text, CboRepudiate3.SelStart + CboRepudiate3.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboRepudiate3.ListCount - 1
       
        If NewText = LCase(Left(CboRepudiate3.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboRepudiate3.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboRepudiate3.Text = ValidValue
            CboRepudiate3.SelStart = 0
            CboRepudiate3.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CmdAddRec_Click()
    Call ChkData
End Sub

Private Sub CmdClaerInv_Click()
    Call ClearEntries
    Call ClearMBMInfo
    Call ClearInvoiceINfo
    Unload frmINVMedicine
'    EditFlag = False
    CmdMBMInfo.Enabled = False
    Call EntriesEnable(False)
End Sub


Private Sub CmdGo_Click()
Dim ZZ As String
    
    txtMBMNo = ChkStr(txtMBMNo)
    
    If Trim(txtMBMNo) <> "" Then
        ZZ = InStr(1, txtMBMNo, "-")
        MembershipNo = Trim(Left(txtMBMNo, IIf(ZZ - 1 <= 0, Len(Trim(txtMBMNo)), ZZ - 1)))
        If ZZ = 0 Then
            CovPersonID = "00"
        Else
            CovPersonID = Trim(Right(txtMBMNo, IIf(ZZ - 1 <= 0, 0, Len(Trim(txtMBMNo)) - ZZ)))
        End If
        Call MBMInfo
        If Trim(MembershipNo) <> "" Then
            EditFlag = False
            Call ClearEntries
            CmdEdit.Visible = True
            Call EntriesEnable(True)
            txtTabMBMNo = MembershipNo
            TxtPayor = strPayor
            txtPlanCode = strPlanCode
            TxtInsType = strInsType
            txtMBMName = strMBMName
            txtTabMBMName = strMBMName
            txtEffDate = strEffdate
            txtExpDate = strExpDate
            TxtPolNo = strPolNo
            txtPolStatus = strPOlStatus
            txtMBMStatus = strMBMDataStatus
            txtdob = strPatientDOB
            txtPatientIC = strPatientIC
            txtPatientName = strPatientName
            TxtBalLimit = strBalLimit
            TxtVisitLimit = strVisitLimit
            TxtClnContract = strCLNContract
        End If
    End If
End Sub

Private Sub CmdMBMInfo_Click()
    FrmINVMBMinfo.Source = 1
    FrmINVMBMinfo.Show
End Sub

Private Sub CmdMedication_Click()
    frmINVMedicine.Source = 3
    frmINVMedicine.Show
End Sub

Private Sub CmdPrint_Click()
If lstInvoice.ListItems.count > 0 Then
    Call ExportToExcelError(lstInvoice)
End If
End Sub

Private Sub cmdProcedures_Click()
    frmINVMedicine.Source = 4
    frmINVMedicine.Show
End Sub

Private Sub CmdLab_Click()
'    frmINVMedicine.Source = 6
'    frmINVMedicine.Show
End Sub

Private Sub CmdRemove_Click()
Dim RecRemove
    
    RecRemove = MsgBox("Are you sure?", vbYesNo + vbExclamation)
    If RecRemove = vbYes Then
        Call RemoveRecFrListV
        Call ClearEntries
        Call ClearMBMInfo
        Call ClearInvoiceINfo
    End If
    
End Sub

Private Sub CmdUpdate_Click()
    EditFlag = False
    Call RemoveRecFrListV
    Call ChkData
    CmdAddRec.Enabled = True
    CmdEdit.Visible = True
    CmdEdit.Enabled = True
    CmdUpdate.Visible = False
    CmdUpdate.Enabled = False
    Call ClearEntries
    Call ClearMBMInfo
    txtMBMNo = ""
End Sub


Private Sub CmdVerify_Click()
Dim CLNRec As New ADODB.Recordset
Dim PLNRec As New ADODB.Recordset
Dim GENRec As New ADODB.Recordset
Dim CLNSql As String
Dim MBMSql, COVSql As String
Dim CC As Integer
Dim ErrMsg As String
Dim ChkSql As String
Dim ChKItemRec As New ADODB.Recordset
Dim VMedChkItem, VBillChkAmt As String

    For CC = 1 To lstInvoice.ListItems.count
    
        MBMSql = "SELECT MBMName,MBMNumber, MBMPayorEffDate, MBMPolicyNo, MBMICBcPp, MBMPayorExpDate,MBMCLNBalLmt, " & _
                " MBMDOB FROM MBM WHERE MBMNumber='" & Trim(lstInvoice.ListItems(CC).SubItems(2)) & "'"
        MBM.Open MBMSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        With MBM
            VMBMNo = !MBMNumber
            VName = !MBMName
            VPolNo = !MBMPolicyNo
            VIc = !MBMIcBcPp
        End With
        MBM.Close
        Set MBM = Nothing
        
        CLNSql = "SELECT MBMPolEffDate, MBMPolExpDate, PLNType FROM MBMCLNOthers WHERE MBMNumber='" & Trim(VMBMNo) & "'"
        CLNRec.Open CLNSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        VCoverID = lstInvoice.ListItems(CC).SubItems(3)
        VPlan = CLNRec!PLNType
        VEffDate = CLNRec!MBMPolEffDate
        VExpDate = CLNRec!MBMPolExpDate
        CLNRec.Close
        Set CLNRec = Nothing
        VBalLmt = lstInvoice.ListItems(CC).SubItems(69)
'        PLNRec.Open "SELECT DiagIND FROM PLN where PLNCode ='" & VPlan & "'", MediConnCISDB, adOpenKeyset, adLockOptimistic
'        VDiagIND = PLNRec!DiagIND
        
        If Len(Trim(lstInvoice.ListItems(CC).SubItems(46))) Then
            VDiagCode = Trim(lstInvoice.ListItems(CC).SubItems(46))
            Call VerifyDiagnosis
        End If
        If Len(Trim(lstInvoice.ListItems(CC).SubItems(47))) Then
            VDiagCode = Trim(lstInvoice.ListItems(CC).SubItems(47))
            Call VerifyDiagnosis
        End If
        If Len(Trim(lstInvoice.ListItems(CC).SubItems(48))) Then
            VDiagCode = Trim(lstInvoice.ListItems(CC).SubItems(48))
            Call VerifyDiagnosis
        End If

        If CDate(lstInvoice.ListItems(CC).SubItems(8)) < CDate(VEffDate) Then
            Call ErrRpt
            lstErrRec.SubItems(9) = "Invoice date MUST BE GREATER than Clinical Effective date ! "
        End If
        If CDate(lstInvoice.ListItems(CC).SubItems(8)) > CDate(VExpDate) + 14 Then
            Call ErrRpt
            lstErrRec.SubItems(9) = "Invoice date IS GREATER THAN Clinical Expiry date + 14 days ! "
        End If
        If IsDate(lstInvoice.ListItems(CC).SubItems(8)) Then
            If (CDate(Date) - CDate(lstInvoice.ListItems(CC).SubItems(8))) > 45 Then
                Call ErrRpt
                lstErrRec.SubItems(9) = "45 Days LATE ! "
            End If
        End If
        If CDate(lstInvoice.ListItems(CC).SubItems(9)) < CDate(VEffDate) Then
            Call ErrRpt
            lstErrRec.SubItems(9) = "Submission date MUST BE GREATER than Effective date ! "
        End If
        If CDate(lstInvoice.ListItems(CC).SubItems(9)) > CDate(VExpDate) + 14 Then
            Call ErrRpt
            lstErrRec.SubItems(9) = "Submission date IS GREATER THAN Expiry date + 14 days ! "
        End If
        
        If IsNumeric(lstInvoice.ListItems(CC).SubItems(68)) Then
            If CDbl(lstInvoice.ListItems(CC).SubItems(13)) > CDbl(lstInvoice.ListItems(CC).SubItems(68)) Then
                Call ErrRpt
                lstErrRec.SubItems(9) = "Excess Balance Limit ! "
            End If
        End If
        
        If IsNumeric(lstInvoice.ListItems(CC).SubItems(69)) Then
            If CDbl(lstInvoice.ListItems(CC).SubItems(13)) > CDbl(lstInvoice.ListItems(CC).SubItems(69)) Then
                Call ErrRpt
                lstErrRec.SubItems(9) = "Excess Visit Limit ! "
            End If
        End If
        
        ChkSql = "SELECT Catcode, Amount FROM CheckItems WHERE CatName = '" & lstInvoice.ListItems(CC).SubItems(12) & _
                "' OR CatName ='" & lstInvoice.ListItems(CC).SubItems(66) & "'"
        ChKItemRec.Open ChkSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        Do Until ChKItemRec.EOF
            If ChKItemRec!CatCode = "BILL" Then
                VBillChkAmt = ChKItemRec!Amount
            Else
                VMedChkAmt = ChKItemRec!Amount
            End If
            ChKItemRec.MoveNext
        Loop
        ChKItemRec.Close
        Set ChKItemRec = Nothing
    
        If CDbl(lstInvoice.ListItems(CC).SubItems(20)) > CDbl(VMedChkAmt) Then
            Call ErrRpt
            lstErrRec.SubItems(9) = "Excess Amount " & lstInvoice.ListItems(CC).SubItems(66)
        End If
        If CDbl(lstInvoice.ListItems(CC).SubItems(13)) > CDbl(VBillChkAmt) Then
            Call ErrRpt
            lstErrRec.SubItems(9) = "Excess Amount " & lstInvoice.ListItems(CC).SubItems(12)
        End If
        ChkSql = "SELECT Catcode, Amount FROM CheckItems WHERE CatCode = 'CSMD'" & _
                " and CatName ='" & lstInvoice.ListItems(CC).SubItems(70) & "'"
        ChKItemRec.Open ChkSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        If ChKItemRec.RecordCount Then
            If IsNumeric(ChKItemRec!Amount) Then
                If CDbl(lstInvoice.ListItems(CC).SubItems(13)) > CDbl(ChKItemRec!Amount) Then
                    Call ErrRpt
                    lstErrRec.SubItems(9) = "Excess Amount " & lstInvoice.ListItems(CC).SubItems(70)
                End If
            End If
        End If
        ChKItemRec.Close
        Set ChKItemRec = Nothing
    Next CC
    CmdVerify.Visible = False
    CmdSave.Visible = True
    CmdSave.Enabled = True
    If lstErr.ListItems.count > 0 Then
        Call ExportToExcelError(lstErr)
    End If
End Sub

Private Sub lstInvoice_Click()
    CmdRemove.Enabled = True
    If EditFlag = True Then
        
        CmdUpdate.Enabled = True
        CmdAddRec.Enabled = False
        txtMBMNo = ""
        txtMBMNo = lstInvoice.SelectedItem.SubItems(2) & "-" & lstInvoice.SelectedItem.SubItems(3)
        MembershipNo = lstInvoice.SelectedItem.SubItems(2)
        Call MBMInfo
        txtTabMBMNo = lstInvoice.SelectedItem.SubItems(2)
        TxtPayor = strPayor
        
        txtPlanCode = strPlanCode
        TxtInsType = strInsType
        txtMBMName = strMBMName
        txtTabMBMName = strMBMName
        txtEffDate = strEffdate
        txtExpDate = strExpDate
        TxtPolNo = strPolNo
        txtPolStatus = strPOlStatus
        txtMBMStatus = strMBMDataStatus
        txtdob = strPatientDOB
        txtPatientIC = strPatientIC
        txtPatientName = strPatientName
        TxtBalLimit = strBalLimit
        TxtVisitLimit = strVisitLimit
        TxtClnContract = strCLNContract
        txtInvoiceNo = lstInvoice.SelectedItem.SubItems(7)
        txtInvoiceDate = lstInvoice.SelectedItem.SubItems(8)
        txtDateVisit = lstInvoice.SelectedItem.SubItems(9)
        CboNatureClm.Text = lstInvoice.SelectedItem.SubItems(10)
        Cboclinic.Text = lstInvoice.SelectedItem.SubItems(11)
        If lstInvoice.SelectedItem.SubItems(12) = "BILL ACCUTE" Then
            OptAccute.Value = True
            strMEdChkItem = "MEDICINE ACCUTE"
        ElseIf lstInvoice.SelectedItem.SubItems(12) = "BILL CHRONIC" Then
            OptChronic.Value = True
            strMEdChkItem = "MEDICINE CHRONIC"
        ElseIf lstInvoice.SelectedItem.SubItems(12) = "BILL MULTIPLE" Then
            OptMultiple.Value = True
            strMEdChkItem = "MEDICINE MULTIPLE"
        End If
    '    lstRec.SubItems(13) = IntInvAmt
        CboPAYto.Text = lstInvoice.SelectedItem.SubItems(14)
        TxtPayee = lstInvoice.SelectedItem.SubItems(15)
        txtDoctor = lstInvoice.SelectedItem.SubItems(16)
    '    lstRec.SubItems(17) = Logon
        A1 = lstInvoice.SelectedItem.SubItems(18)
        A2 = lstInvoice.SelectedItem.SubItems(19)
        A3 = lstInvoice.SelectedItem.SubItems(20)
        A4 = lstInvoice.SelectedItem.SubItems(21)
        A5 = lstInvoice.SelectedItem.SubItems(22)
        A6 = lstInvoice.SelectedItem.SubItems(23)
        A7 = lstInvoice.SelectedItem.SubItems(24)
        A8 = lstInvoice.SelectedItem.SubItems(25)
        A9 = lstInvoice.SelectedItem.SubItems(26)
        A10 = lstInvoice.SelectedItem.SubItems(27)
        A11 = lstInvoice.SelectedItem.SubItems(28)
        A12 = lstInvoice.SelectedItem.SubItems(29)
        A13 = lstInvoice.SelectedItem.SubItems(30)
        A14 = lstInvoice.SelectedItem.SubItems(31)
        A15 = lstInvoice.SelectedItem.SubItems(32)
        txtOtherDesc1 = lstInvoice.SelectedItem.SubItems(33)
'        txtOtherDesc2 = lstInvoice.SelectedItem.SubItems(34)
        A16 = lstInvoice.SelectedItem.SubItems(35)
        A17 = lstInvoice.SelectedItem.SubItems(36)
        B1 = lstInvoice.SelectedItem.SubItems(37)
        B2 = lstInvoice.SelectedItem.SubItems(38)
        B3 = lstInvoice.SelectedItem.SubItems(39)
        B4 = lstInvoice.SelectedItem.SubItems(40)
        B5 = lstInvoice.SelectedItem.SubItems(41)
        B6 = lstInvoice.SelectedItem.SubItems(42)
        B7 = lstInvoice.SelectedItem.SubItems(43)
        cboClaimability.Text = lstInvoice.SelectedItem.SubItems(44)
        cboPending.Text = lstInvoice.SelectedItem.SubItems(45)
        CboReason1.Text = lstInvoice.SelectedItem.SubItems(46)
        CboReason2.Text = lstInvoice.SelectedItem.SubItems(47)
        CboReason3.Text = lstInvoice.SelectedItem.SubItems(48)
        txtMC = lstInvoice.SelectedItem.SubItems(49)
        TxtRemarks = lstInvoice.SelectedItem.SubItems(50)
        CboRepudiate1.Text = lstInvoice.SelectedItem.SubItems(51)
        CboRepudiate2.Text = lstInvoice.SelectedItem.SubItems(52)
        CboRepudiate3.Text = lstInvoice.SelectedItem.SubItems(53)
        frmINVMedicine.txtMEdication1.Text = lstInvoice.SelectedItem.SubItems(54)
        frmINVMedicine.txtMEdication2.Text = lstInvoice.SelectedItem.SubItems(55)
        frmINVMedicine.txtMEdication3.Text = lstInvoice.SelectedItem.SubItems(56)
        frmINVMedicine.CboProc1.Text = lstInvoice.SelectedItem.SubItems(57)
        frmINVMedicine.CboProc2.Text = lstInvoice.SelectedItem.SubItems(58)
        frmINVMedicine.CboProc3.Text = lstInvoice.SelectedItem.SubItems(59)
        frmINVMedicine.CboProc4.Text = lstInvoice.SelectedItem.SubItems(60)
        frmINVMedicine.CboProc5.Text = lstInvoice.SelectedItem.SubItems(61)
        frmINVMedicine.CboProc6.Text = lstInvoice.SelectedItem.SubItems(62)
        frmINVMedicine.CboProc7.Text = lstInvoice.SelectedItem.SubItems(63)
        TxtInvPayType = lstInvoice.SelectedItem.SubItems(64)
        TxtPayor = lstInvoice.SelectedItem.SubItems(65)
    End If

End Sub

Private Sub OptAccute_LostFocus()
    If OptAccute.Value = True Then
        strBillChkItem = "BILL ACCUTE"
        strMEdChkItem = "MEDICINE ACCUTE"
    End If
End Sub

Private Sub OptChronic_LostFocus()
    If OptChronic.Value = True Then
        strBillChkItem = "BILL CHRONIC"
        strMEdChkItem = "MEDICINE CHRONIC"
    End If
End Sub

Private Sub OptMultiple_LostFocus()
    If OptMultiple.Value = True Then
        strBillChkItem = "BILL MULTIPLE"
        strMEdChkItem = "MEDICINE MULTIPLE"
    End If
End Sub

Private Sub txtDoctor_LostFocus()
    txtDoctor = ChkStr(txtDoctor)
End Sub

Private Sub txtMBMNo_GotFocus()
    CmdGo.Default = True
End Sub

Private Sub txtMC_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Cboclinic_LostFocus()
    If Mid(CboPAYto, 1, 1) = "C" Then
        TxtPayee = Cboclinic
'    Else
'        TxtPayee = strPatientID
    End If
End Sub

Private Sub CboPayee_LostFocus()
    CboPayee = ChkStr(CboPayee)
End Sub


Private Sub CboPAYto_LostFocus()
    If Mid(CboPAYto, 1, 1) = "C" Then
        TxtInvPayType = "Cashless"
    Else
        TxtInvPayType = "Reimbursement"
        TxtPayee = "MEMBER"
    End If
End Sub

Private Sub CmdClearAll_Click()
    Call ClearEntries
    Call ClearMBMInfo
    Call ClearInvoiceINfo
    lstInvoice.ListItems.Clear
    CmdSave.Enabled = False
    CmdSave.Visible = False
    EditFlag = False
    CmdMBMInfo.Enabled = False
    CmdRemove.Enabled = True
End Sub

Private Sub cmdEdit_Click()
    CmdEdit.Visible = False
    CmdEdit.Enabled = False
    CmdUpdate.Visible = True
    CmdUpdate.Enabled = True
    CmdSave.Enabled = False
    Call EntriesEnable(True)
    EditFlag = True
    txtInvoiceNo.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
Dim RecSaved

    If lstInvoice.ListItems.count <> 0 Then
        RecSaved = MsgBox("Are you sure ? ", vbYesNo + vbExclamation)
        If RecSaved = vbYes Then
            For AA = 1 To lstInvoice.ListItems.count
                Call SaveRecord
            Next AA
            
            MsgBox "Record Saved ! ", vbOKOnly
    
            CmdSave.Enabled = False
            CmdSave.Visible = False
            CmdVerify.Visible = True
            CmdVerify.Enabled = True
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
Private Sub Form_Load()
    bExit = False
    EditFlag = False
    Call EntriesEnable(False)
    Call ComboListing
    CmdClearAll.Enabled = True
    CmdSave.Enabled = False
    CmdSave.Visible = False
    strBillChkItem = "BILL ACCUTE"
    strMEdChkItem = "MEDICINE ACCUTE"
    intIndex = 0
End Sub

Private Sub ComboListing()
Dim ClinicRec As New ADODB.Recordset
Dim RepRec As New ADODB.Recordset
Dim TreatRec As New ADODB.Recordset
Dim NatureCLM As New ADODB.Recordset
        
    TxtInvPayType = "Cashless"
    Cboclinic.Clear
'    CboPayee.Clear
    ClinicRec.Open "SELECT CLNCode, CLNName FROM CLN WHERE CLNPanel = 'Y'", medixmasterconn, adOpenKeyset, adLockOptimistic
    Do Until ClinicRec.EOF
        Cboclinic.AddItem ClinicRec!CLNCode & " - " & ClinicRec!CLNName
'        CboPayee.AddItem ClinicRec!CLNCode & " - " & ClinicRec!CLNName
        ClinicRec.MoveNext
    Loop
    ClinicRec.Close
    Set ClinicRec = Nothing
    
    CboRepudiate1.Clear
    CboRepudiate2.Clear
    CboRepudiate3.Clear
    RepRec.Open "SELECT REPCode, RepDescription FROM REPCode", medixmasterconn, adOpenKeyset, adLockOptimistic
    Do Until RepRec.EOF
        CboRepudiate1.AddItem RepRec!REPCode & " - " & RepRec!RepDescription
        CboRepudiate2.AddItem RepRec!REPCode & " - " & RepRec!RepDescription
        CboRepudiate3.AddItem RepRec!REPCode & " - " & RepRec!RepDescription
        RepRec.MoveNext
    Loop
    RepRec.Close
    Set RepRec = Nothing
    
    CboReason1.Clear
    CboReason2.Clear
    CboReason3.Clear
    TreatRec.Open "SELECT ReasonCode, ReasonDesc FROM TreatmentReason", MediConnCISDB, adOpenKeyset, adLockOptimistic
    Do Until TreatRec.EOF
        CboReason1.AddItem TreatRec!ReasonCode & " - " & TreatRec!ReasonDesc
        CboReason2.AddItem TreatRec!ReasonCode & " - " & TreatRec!ReasonDesc
        CboReason3.AddItem TreatRec!ReasonCode & " - " & TreatRec!ReasonDesc
        TreatRec.MoveNext
    Loop
    TreatRec.Close
    Set RepRec = Nothing
    
    CboNatureClm.Clear
    NatureCLM.Open "SELECT ClaimNatureCode, ClaimNatureName FROM ClaimNature", medixmasterconn, adOpenKeyset, adLockOptimistic
    Do Until NatureCLM.EOF
        CboNatureClm.AddItem NatureCLM!ClaimNatureCode & " - " & NatureCLM!ClaimNatureName
        NatureCLM.MoveNext
    Loop
    NatureCLM.Close
    Set NatureCLM = Nothing
    
    CboPAYto.Clear
    CboPAYto.AddItem "C - Clinic"
    CboPAYto.AddItem "M - Member"
    CboPAYto.ListIndex = 0

    cboClaimability.Clear
    cboClaimability.AddItem "A - Accepted"
    cboClaimability.AddItem "R - Rejected"
    cboClaimability.AddItem "C - Cancelled"
    cboClaimability.ListIndex = 0
    
    cboPending.Clear
    cboPending.AddItem "Y - Yes"
    cboPending.AddItem "N - No"
    cboPending.ListIndex = 1
    
    cboStatus.Clear
    cboStatus.AddItem "O - Open"
    cboStatus.AddItem "C - Close"
    cboStatus.AddItem "D - Delete"
    cboStatus.ListIndex = 0
End Sub
Private Sub ClearEntries()
   On Error Resume Next
   Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            If ctl.Name <> "txtMBMNo" And ctl.Name <> "txtDoctor" And ctl.Name <> "TxtPayee" Then
               ctl.Text = ""
            End If
        ElseIf TypeOf ctl Is ComboBox Then
        
'        ElseIf TypeOf ctl Is ListView Then
'            ctl.ListItems.Clear
        ElseIf TypeOf ctl Is MaskEdBox Then
            If ctl.Name <> "txtDateVisit" Then
                ctl.mask = ""
                ctl.Text = ""
                ctl.mask = "##/##/####"
            End If
        End If
    Next ctl
    CboNatureClm.Text = ""
    CboReason1.Text = ""
    CboReason2.Text = ""
    CboReason3.Text = ""
    CboRepudiate1.Text = ""
    CboRepudiate2.Text = ""
    CboRepudiate3.Text = ""
    cboClaimability.Text = "A - Accepted"
    cboPending.Text = "N - No"
    OptAccute.Value = True
    lblTotAmt.Visible = False
    lblMedicine.Visible = False
    LblExpDate.Visible = False
    lblBordx = ""
    Lbl45days.Visible = False
    lblWs = ""
    IntInvAmt = 0
    txtVisitTime.mask = ""
    txtVisitTime.Text = ""
    txtVisitTime.mask = "##:##"
    If Mid(CboPAYto, 1, 1) = "C" Then
        TxtInvPayType = "Cashless"
    Else
        TxtInvPayType = "Reimbursement"
        TxtPayee.Text = "MEMBER"
    End If
End Sub

Private Sub txtInvoiceNo_LostFocus()
    txtInvoiceNo = ChkStr(txtInvoiceNo)
End Sub

Private Sub txtMBMNo_LostFocus()
    CmdGo.Default = False
End Sub

Private Sub txtDateVisit_LostFocus()
    If txtDateVisit <> "__/__/____" Then
        If IsDate(txtDateVisit) = False Then
            MsgBox "Invalid date format!", vbCritical
            txtDateVisit.SetFocus
        ElseIf CDate(txtDateVisit) > CDate(Date) Then
            MsgBox "Submission Date cannot greater than system date ! ", vbCritical
            txtDateVisit.SetFocus
        End If
    End If
End Sub

Private Sub txtInvoiceDate_LostFocus()
        If txtInvoiceDate <> "__/__/____" Then
        If IsDate(txtInvoiceDate) = False Then
            MsgBox "Invalid date format!", vbCritical
            txtInvoiceDate.SetFocus
        ElseIf CDate(txtInvoiceDate) > CDate(Date) Then
            MsgBox "Invoice date cannot greater than system date ! ", vbCritical
            txtInvoiceDate.SetFocus
        End If
    End If
End Sub

Private Sub MBMInfo()
Dim MBMRec As New ADODB.Recordset
Dim CLNOth As New ADODB.Recordset
Dim MBMSql, CLNSql, COVSql, COVLmtSql As String
Dim COVRec As New ADODB.Recordset
Dim Cnt As Integer
Dim CISPln As New ADODB.Recordset
Dim CISAnnLMt As New ADODB.Recordset
Dim CISPlnSql, CISAnnLmtSql, SuppLmtType As String
Dim COVLmt As New ADODB.Recordset

    If Len(Trim(MembershipNo)) Then
        MBMSql = "SELECT MBMNumber, MBMName, MBMICBcPp, PAYCode, MBMCLNBalLmt, MBMStatus, MBMCOVID, MBMCLNVisitLmt, " & _
                " MBMDOB, MBMDataStatus, HLTCode FROM MBM WHERE MBMNumber='" & Trim(MembershipNo) & " ' AND MBMDataStatus <> 'D' " & _
                " AND MBMStatus <> 'D' AND (HSCLNInd = 'C' OR HSCLNInd = 'M')"
        MBMRec.Open MBMSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If MBMRec.RecordCount <= 0 Then
            MsgBox "Membership No not found ! ", vbOKOnly + vbCritical, DialogBoxTitle
            TxtInvPayType = "Cashless"
            CboPAYto.Text = "C - Clinic"
            txtMBMNo = MembershipNo & "-" & CovPersonID
            MembershipNo = ""
            CovPersonID = ""
            Call ClearEntries
            Call ClearMBMInfo
            CmdAddRec.Enabled = False
            CmdMBMInfo.Enabled = False
            txtMBMNo.SetFocus
            
        Else
            CLNSql = "SELECT MBMNumber, MBMPolNumber, MBMPolEffDate, MBMPolExpDate, " & _
                    " INSType, PLNType FROM MBMCLNOthers WHERE MBMNumber='" & Trim(MembershipNo) & "'"
            CLNOth.Open CLNSql, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            CISPlnSql = "SELECT AnnLmtIND, VisLmtIND, ClnContract from PLN WHERE PLNCode ='" & Trim(CLNOth!PLNType) & "' and HLTCode='" & Trim(MBMRec!HLTCode) & "'"
            CISPln.Open CISPlnSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            
            If Trim(CISPln!AnnLmtIND) = "Y" Then
                CISAnnLmtSql = "SELECT SuppLimitStatus FROM AnnualLimit WHERE PLNCode='" & Trim(CLNOth!PLNType) & "' and HLTCode='" & Trim(MBMRec!HLTCode) & "' AND INSCode='" & Trim(CLNOth!INSType) & "'"
                CISAnnLMt.Open CISAnnLmtSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            
                SuppLmtType = CISAnnLMt!SuppLimitStatus
                CISAnnLMt.Close
                Set CISAnnLMt = Nothing
            Else
                SuppLmtType = ""
            End If
            
            If Trim(CISPln!AnnLmtIND) = "N" Then
                strBalLimit = "FULL"
            Else
                strBalLimit = MBMRec!MBMCLNBalLmt
            End If
            If Trim(CISPln!VisLmtIND) = "N" Then
                strVisitLimit = "FULL"
            Else
                strVisitLimit = MBMRec!MBMCLNVisitLmt
            End If
            strCLNContract = CISPln!ClnContract
            
            CISPln.Close
            Set CISPln = Nothing
                        
            strMBMNo = MBMRec!MBMNumber
            strPayor = MBMRec!PAYCode
            strPlanCode = CLNOth!PLNType
            strInsType = CLNOth!INSType
            strMBMName = MBMRec!MBMName
            strTabMBMName = MBMRec!MBMName
            strEffdate = CLNOth!MBMPolEffDate
            strExpDate = CLNOth!MBMPolExpDate
            strPolNo = CLNOth!MBMPolNumber
            strMBMDataStatus = MBMRec!MBMDataStatus
            strPOlStatus = MBMRec!MBMStatus
            strPatientName = MBMRec!MBMName
            strPatientDOB = MBMRec!MBMDOB
            strPatientIC = MBMRec!MBMIcBcPp
            CmdAddRec.Enabled = True
            CmdMBMInfo.Enabled = True
            If CovPersonID <> "00" Then
                COVSql = "SELECT MBMCName, MBMCCoverID, MBMCICBCPPNo, MBMCDOB, MBMCCLNBalLmt, MBMCCLNVisitLmt FROM MBMCoveredPersons WHERE MBMCNumber='" & Trim(MembershipNo) & "' AND MBMCCoverID = '" & Trim(CovPersonID) & "'"
                COVRec.Open COVSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                If COVRec.RecordCount <= 0 Then
                    MsgBox "No such Cover Person ! ", vbOKOnly + vbCritical, DialogBoxTitle
                    TxtInvPayType = "Cashless"
                    CboPAYto.Text = "C - Clinic"
                    txtMBMNo = MembershipNo & "-" & CovPersonID
                    MembershipNo = ""
                    CovPersonID = ""
                    Call ClearEntries
                    Call ClearMBMInfo
                    CmdAddRec.Enabled = False
                    txtMBMNo.SetFocus
                Else
                    If SuppLimitStatus = "B" Then
                        If CovPersonID <> "01" Then
                            COVLmtSql = "SELECT MBMCCLNBalLmt, MBMCCLNVisitLmt FROM MBMCoveredPersons WHERE MBMCNumber='" & Trim(MembershipNo) & "' AND MBMCCoverID = '01'"
                            COVLmt.Open COVSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                            strBalLimit = COVLmt!MBMCCLNBalLmt
                            strVisitLimit = COVLmt!MBMCCLNVisitLmt
                            COVLmt.Close
                            Set COVLmt = Nothing
                        End If
                    ElseIf SuppLimitStatus = "C" Then
                        COVLmtSql = "SELECT MBMCCLNBalLmt, MBMCCLNVisitLmt FROM MBMCoveredPersons WHERE MBMCNumber='" & Trim(MembershipNo) & "' AND MBMCCoverID = '" & Trim(CovPersonID) & "'"
                        COVLmt.Open COVSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                        strBalLimit = COVLmt!MBMCCLNBalLmt
                        strVisitLimit = COVLmt!MBMCCLNVisitLmt
                        COVLmt.Close
                        Set COVLmt = Nothing
                    End If
                    strPatientName = COVRec!MBMCName
                    strPatientDOB = COVRec!MBMCDOB
                    strPatientIC = COVRec!MBMCICBCPPNo
                    strPatientID = COVRec!MBMCCoverID
                End If
                COVRec.Close
                Set COVRec = Nothing
                CLNOth.Close
                Set CLNOth = Nothing
            End If
        End If
        MBMRec.Close
        Set MBMRec = Nothing
        
    End If
End Sub

Private Function GenerateCISCaseNo(PayorCode)
    Dim CISCAQ As New ADODB.Recordset
    Dim strsql As String
    Dim CISSeqNo As String
    Dim CAQCharacter As String
    
    strsql = "Select PAYCode, CISSeqNo from CISCAQ Where PAYCode = '" & PayorCode & "'"
    CISCAQ.Open strsql, MediConnCISDB, adOpenStatic, adLockOptimistic
    
    If CISCAQ.RecordCount <= 0 Then
        With CISCAQ
            .AddNew
            !PAYCode = PayorCode
            !CISSeqNo = "000001"
            .Update
            CISSeqNo = "000001"
        End With
    Else
        CISCAQ!CISSeqNo = Format(CISCAQ!CISSeqNo + 1, "000000")
        CISSeqNo = CISCAQ!CISSeqNo
        CISCAQ.Update
    End If

    GenerateCISCaseNo = PayorCode & "K" & CISSeqNo
    CISCAQ.Close
    Set CISCAQ = Nothing
End Function



Private Sub SaveRecord()
Dim startTrans As Boolean
    
            MediConnCISWS.beginTrans
            startTrans = True
            Call SaveCase
            Call SaveInvoice
            Call saveActual
            Call saveActualOthers
            Call saveNotCover
        
            MediConnCISWS.CommitTrans
            startTrans = False

            CmdClearAll.Enabled = True

            Call EntriesEnable(False)
            CmdEdit.Visible = True
            CmdEdit.Enabled = True
            EditFlag = False
    Screen.MousePointer = vbDefault
Exit Sub
ErrorSave:
    Screen.MousePointer = vbDefault
    If startTrans = True Then
        MediConnCISWS.RollbackTrans
    End If
    MsgBox Err.Description
End Sub

Private Function getTotal() As Double
    Dim ctl  As Control
    IntInvAmt = 0
    For Each ctl In Me.Controls
        If Mid(ctl.Name, 1, 1) = "A" And Len(ctl.Name) <= 7 Then
            If IsNumeric(ctl.Text) = True Then
                IntInvAmt = IntInvAmt + CDbl(ctl.Text)
            End If
        End If
    Next
End Function
Private Sub SaveInvoice()
Dim ClaimInvoice  As New ADODB.Recordset
    
    ClaimInvoice.Open "Select INVINdex , * from ClaimInvoice", MediConnCISWS, adOpenKeyset, adLockOptimistic
    ClaimInvoice.AddNew
    With ClaimInvoice
        !INVWSStatus = lstInvoice.ListItems(AA).SubItems(6)
        !INVBordxed = lstInvoice.ListItems(AA).SubItems(5)
        !INVEnteredDate = Date
        !INVEnteredBy = Logon
        !INVNo = lstInvoice.ListItems(AA).SubItems(7)
        !CASNumber = NewCaseNo
        !InvDate = lstInvoice.ListItems(AA).SubItems(8)
        !MCDays = lstInvoice.ListItems(AA).SubItems(49)
        !INVLastUpdateBy = Logon
        !INVLastUpdateDate = Date
        !INVAmount = lstInvoice.ListItems(AA).SubItems(13)
        !INVPayee = lstInvoice.ListItems(AA).SubItems(15)
        !INVPAYType = lstInvoice.ListItems(AA).SubItems(64)
        !INVPayTo = lstInvoice.ListItems(AA).SubItems(14)

    End With
    ClaimInvoice.Update
    
'    If EditFlag = False Then
        Dim ClaimInvoiceLastIndex  As New ADODB.Recordset
        ClaimInvoiceLastIndex.Open "Select INVINdex from ClaimInvoice Order by INVIndex DESC", MediConnCISWS, adOpenKeyset, adLockOptimistic
        If Not ClaimInvoiceLastIndex.EOF Then
            intIndex = ClaimInvoiceLastIndex!InvINdex
        Else
            intIndex = 1
        End If
        ClaimInvoiceLastIndex.Close
        Set ClaimInvoiceLastIndex = Nothing
    
    ClaimInvoice.Close
    Set ClaimInvoice = Nothing
End Sub
Private Sub saveActual()
    Dim ClaimActual As New ADODB.Recordset
    Dim strsql As String
 
    ClaimActual.Open "Select INVINdex , * from ClaimInvoiceActual", MediConnCISWS, adOpenKeyset, adLockOptimistic
    ClaimActual.AddNew
    ClaimActual!InvINdex = intIndex
    
    With ClaimActual
        !CLMGPConsAmt = lstInvoice.ListItems(AA).SubItems(18)
        !CLMSPConsAmt = lstInvoice.ListItems(AA).SubItems(19)
        !CLMMedicine = lstInvoice.ListItems(AA).SubItems(20)
        !CLMProcedures = lstInvoice.ListItems(AA).SubItems(21)
        !CLMLab = lstInvoice.ListItems(AA).SubItems(22)
        !CLMXray = lstInvoice.ListItems(AA).SubItems(23)
        !CLMDental = lstInvoice.ListItems(AA).SubItems(24)
        !CLMChineseMed = lstInvoice.ListItems(AA).SubItems(25)
        !CLMTraditionalMEd = lstInvoice.ListItems(AA).SubItems(26)
        !CLMphysioTheraphy = lstInvoice.ListItems(AA).SubItems(27)
        !CLMAccident = lstInvoice.ListItems(AA).SubItems(28)
        !CLMPRVCheckup = lstInvoice.ListItems(AA).SubItems(29)
        !CLMVision = lstInvoice.ListItems(AA).SubItems(30)
        !CLMID = lstInvoice.ListItems(AA).SubItems(31)
        !CLMTSSTO = lstInvoice.ListItems(AA).SubItems(32)
        !CLMOthers1Desc = lstInvoice.ListItems(AA).SubItems(33)
'        !CLMOthers2Desc = lstInvoice.ListItems(AA).SubItems(34)
        !CLMInjection = lstInvoice.ListItems(AA).SubItems(35)
        !CLMOTHERS1 = lstInvoice.ListItems(AA).SubItems(36)
'        !CLMOTHERS2 = lstInvoice.ListItems(AA).SubItems(36)
    End With
    ClaimActual.Update
    
    ClaimActual.Close
    Set ClaimActual = Nothing
End Sub

Private Sub saveNotCover()
Dim ClaimNotCover  As New ADODB.Recordset
Dim strsql As String
    If EditFlag = False Then
        ClaimNotCover.Open "Select INVINdex , * from ClaimInvoiceNotCover ", MediConnCISWS, adOpenKeyset, adLockOptimistic
        ClaimNotCover.AddNew
        ClaimNotCover!InvINdex = intIndex
    Else
        strsql = "Select INVINdex , * from ClaimInvoiceNotCover where INVIndex =" & intIndex
        ClaimNotCover.Open strsql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    End If
    With ClaimNotCover
        !SPFee = lstInvoice.ListItems(AA).SubItems(37)
        !Medicine = lstInvoice.ListItems(AA).SubItems(38)
        !Dental = lstInvoice.ListItems(AA).SubItems(39)
        !Physio = lstInvoice.ListItems(AA).SubItems(40)
        !Maternity = lstInvoice.ListItems(AA).SubItems(41)
        !PRVCheckup = lstInvoice.ListItems(AA).SubItems(42)
        !Vision = lstInvoice.ListItems(AA).SubItems(43)
    End With
    ClaimNotCover.Update
    
    ClaimNotCover.Close
    Set ClaimNotCover = Nothing
End Sub

Private Sub SaveCase()
Dim CISCAS  As New ADODB.Recordset
Dim CISSql As String
    
    CISSql = "SELECT * FROM CISCAS "
    CISCAS.Open CISSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    CISCAS.AddNew
    NewCaseNo = GenerateCISCaseNo(lstInvoice.ListItems(AA).SubItems(65))

    With CISCAS
        !CASNumber = NewCaseNo
        lstInvoice.ListItems(AA).SubItems(1) = NewCaseNo
        !MBMNumber = lstInvoice.ListItems(AA).SubItems(2)
        !COVERID = lstInvoice.ListItems(AA).SubItems(3)
        !CASStatus = "O"
        !Claimability = lstInvoice.ListItems(AA).SubItems(44)
        !Pending = lstInvoice.ListItems(AA).SubItems(45)
        !REgisterDate = Date
        !REgisterBy = Logon
        !ChkItemCat = lstInvoice.ListItems(AA).SubItems(12)
        !Remarks = lstInvoice.ListItems(AA).SubItems(50)
        !ClinicCode = lstInvoice.ListItems(AA).SubItems(11)
        !DateVisit = lstInvoice.ListItems(AA).SubItems(9)
        If isTime(lstInvoice.ListItems(AA).SubItems(67)) Then
            !TimeVisit = lstInvoice.ListItems(AA).SubItems(67)
        End If
        !CASDoctor = lstInvoice.ListItems(AA).SubItems(16)
        !NatureCLM = lstInvoice.ListItems(AA).SubItems(10)
        !CASTreatReason1 = lstInvoice.ListItems(AA).SubItems(46)
        !CASTreatReason2 = lstInvoice.ListItems(AA).SubItems(47)
        !CASTreatReason3 = lstInvoice.ListItems(AA).SubItems(48)
        If lstInvoice.ListItems(AA).SubItems(44) = "R" Then
            Call SaveRepudiation
        End If
      
    End With
    CISCAS.Update
    
    CISCAS.Close
    Set CISCAS = Nothing
End Sub
Private Sub SaveRepudiation()
Dim CASREP As New ADODB.Recordset
Dim CAS As New ADODB.Recordset
Dim strsql As String
    
    CASREP.Open "select * from CISCASRepudiation", MediConnCISDB, adOpenKeyset, adLockOptimistic
        If CASREP.RecordCount = 0 Then
            CASREP.AddNew
            
        End If
        
        With CASREP
            
            !CASNumber = CASENumber
            !CASREPCode1 = lstInvoice.ListItems(AA).SubItems(44)
            !CASREPCode2 = lstInvoice.ListItems(AA).SubItems(51)
            !CASREPCode3 = lstInvoice.ListItems(AA).SubItems(52)
            
            !CASREPIssuedBy = Logon
            !CASREPDate = Date
        End With
        CASREP.Update
       
    CASREP.Close
    Set CASREP = Nothing
End Sub

Private Sub EntriesEnable(status As Boolean)
    FrameInvoice.Enabled = status
    If Mid(cboClaimability, 1, 1) = "R" Then
        FrameRepudiate.Enabled = True
    End If
End Sub

Private Sub txtOtherDesc1_LostFocus()
    txtOtherDesc1 = ChkStr(txtOtherDesc1)
End Sub

Private Sub TxtPayee_Change()
    TxtPayee = ChkStr(TxtPayee)
End Sub

Private Sub TxtRemarks_LostFocus()
    TxtRemarks = ChkStr(TxtRemarks)
End Sub

Private Sub saveActualOthers()
    Dim ClmActOth As New ADODB.Recordset
    Dim strsql As String
 
    If EditFlag = False Then
        ClmActOth.Open "Select INVINdex , * from ClaimInvoiceActualOthers", MediConnCISWS, adOpenKeyset, adLockOptimistic
        ClmActOth.AddNew
        ClmActOth!InvINdex = intIndex
    Else
        strsql = "Select INVINdex , * from ClaimInvoiceActualOthers where INVIndex =" & intIndex
        ClmActOth.Open strsql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    End If
    
    With ClmActOth
        !CLMMedication1 = lstInvoice.ListItems(AA).SubItems(54)
        !CLMMedication2 = lstInvoice.ListItems(AA).SubItems(55)
        !CLMMedication3 = lstInvoice.ListItems(AA).SubItems(56)
        !CLMProc1 = lstInvoice.ListItems(AA).SubItems(57)
        !CLMProc2 = lstInvoice.ListItems(AA).SubItems(58)
        !CLMProc3 = lstInvoice.ListItems(AA).SubItems(59)
        !CLMProc4 = lstInvoice.ListItems(AA).SubItems(60)
        !CLMProc5 = lstInvoice.ListItems(AA).SubItems(61)
        !CLMProc6 = lstInvoice.ListItems(AA).SubItems(62)
        !CLMProc7 = lstInvoice.ListItems(AA).SubItems(63)
    End With
    ClmActOth.Update
    
    ClmActOth.Close
    Set ClmActOth = Nothing
End Sub

Private Sub ChkData()
Dim CISCasRec As New ADODB.Recordset
Dim CISCasSql, ChkDupKey As String
Dim XX, QQ, BB As Integer
Dim AddNewInv As Boolean
Dim DupInvoice
Dim ArrChk() As String

    Call getTotal
    If Trim(txtMBMNo) = "" Then
        MsgBox "Please keyin Member No ! ", vbCritical
        txtMBMNo.SetFocus
    ElseIf txtInvoiceDate = "__/__/____" Or txtInvoiceDate = "" Then
        MsgBox "Please key in Invoice date! ", vbCritical
        txtInvoiceDate.SetFocus
    ElseIf txtDateVisit = "__/__/____" Or txtDateVisit = "" Then
        MsgBox "Please Enter Submission Date! ", vbCritical
        txtDateVisit.SetFocus
    ElseIf Trim(Cboclinic) = "" Then
        MsgBox "Please keyin the Clinic code ! ", vbCritical
        Cboclinic.SetFocus
    ElseIf Trim(CboPAYto) = "" Then
        MsgBox "Please select who to be paid (Member / Clinic )! ", vbCritical
        CboPAYto.SetFocus
    ElseIf Trim(TxtPayee) = "" Then
        MsgBox "Please keyin the person/Clinic to be paid! ", vbCritical
        TxtPayee.SetFocus
     ElseIf Mid(CboPAYto, 1, 1) = "C" And Trim(Left(Cboclinic, IIf(InStr(1, Cboclinic, "-") - 1 < 0, Len(Cboclinic), InStr(1, Cboclinic, "-") - 1))) <> Trim(Left(TxtPayee, IIf(InStr(1, TxtPayee, "-") - 1 < 0, Len(TxtPayee), InStr(1, TxtPayee, "-") - 1))) Then
        MsgBox "Clinic code AND Payee code is not same !", vbCritical
        Cboclinic.SetFocus
    ElseIf IntInvAmt = 0 Then
        MsgBox "Total Invoice Amount is cannot be Zero ! ", vbCritical
        A1.SetFocus
    Else
        Erase ArrChk
        ReDim ArrChk(0)
        QQ = 0
        CISCasSql = "SELECT * FROM View_Clm_Invoice where MBMNumber = '" & Trim(MembershipNo) & _
              "' AND COVERID ='" & CovPersonID & "' AND INVDate = '" & Format(txtInvoiceDate, "mm/dd/yyyy") & "'"
        If Trim(txtInvoiceNo) <> "" Then
            CISCasSql = CISCasSql & " AND InvNo = '" & Trim(txtInvoiceNo) & "'"
        End If
        CISCasRec.Open CISCasSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
        If CISCasRec.RecordCount Then
            Do Until CISCasRec.EOF
                QQ = QQ + 1
                ReDim Preserve ArrChk(QQ)
                ArrChk(QQ) = Trim(CISCasRec!MBMNumber & "~" & CISCasRec!COVERID & "#" & CISCasRec!DateVisit & "*" & CISCasRec!INVNo)
                CISCasRec.MoveNext
            Loop
        End If
        CISCasRec.Close
        Set CISCasRec = Nothing
        For BB = 1 To lstInvoice.ListItems.count
            ReDim Preserve ArrChk(QQ + BB)
            ArrChk(QQ + BB) = Trim(lstInvoice.ListItems(BB).SubItems(2) & "~" & lstInvoice.ListItems(BB).SubItems(3) & "#" & lstInvoice.ListItems(BB).SubItems(9) & "*" & lstInvoice.ListItems(BB).SubItems(7))
        Next BB
        ChkDupKey = MembershipNo & "~" & CovPersonID & "#" & txtDateVisit & "*" & txtInvoiceNo
        AddNewInv = True
        For XX = 0 To UBound(ArrChk)
            If Trim(ChkDupKey) = ArrChk(XX) Then
                AddNewInv = False
                If Trim(txtInvoiceNo) = "" Then
                    DupInvoice = MsgBox("Invoice has been entered !, Press 'Yes' to Proceed, 'No' to Cancel", vbYesNo + vbExclamation)
                    If DupInvoice = vbYes Then
                        Call AssignData2ListView
                    ElseIf DupInvoice = vbNo Then
                        Call EntriesEnable(False)
                    End If
                Else
                    MsgBox "Record Duplicated, Same Invoice has been entered ", vbCritical, "Error"
                End If
            End If
        Next XX
        If AddNewInv = True Then
            Call AssignData2ListView
        End If
        CmdVerify.Visible = True
        CmdVerify.Enabled = True
        CmdEdit.Visible = True
        CmdEdit.Enabled = True
        CmdSave.Visible = False
        CmdSave.Enabled = False
        CmdMBMInfo.Enabled = False
    End If
End Sub

Private Sub AssignData2ListView()
Dim lstRec
    CmdSave.Enabled = True
    Set lstRec = lstInvoice.ListItems.Add(, , "")
    
    Call getTotal
    lstRec.SubItems(2) = MembershipNo
    lstRec.SubItems(3) = CovPersonID
    lstRec.SubItems(4) = txtPatientName
    lstRec.SubItems(5) = 0
    lstRec.SubItems(6) = 0
    lstRec.SubItems(7) = IIf(Len(Trim(txtInvoiceNo)), txtInvoiceNo, "")
    lstRec.SubItems(8) = IIf(IsDate(txtInvoiceDate), txtInvoiceDate, "__/__/____")
    lstRec.SubItems(9) = IIf(IsDate(txtDateVisit), txtDateVisit, "__/__/____")
    lstRec.SubItems(10) = Trim(Left(CboNatureClm, IIf(InStr(1, CboNatureClm, "-") - 1 < 0, Len(CboNatureClm), InStr(1, CboNatureClm, "-") - 1)))
    lstRec.SubItems(11) = Trim(Left(Cboclinic, IIf(InStr(1, Cboclinic, "-") - 1 < 0, Len(Cboclinic), InStr(1, Cboclinic, "-") - 1)))
    lstRec.SubItems(12) = strBillChkItem
    lstRec.SubItems(13) = IntInvAmt
    lstRec.SubItems(14) = Mid(CboPAYto, 1, 1)
    lstRec.SubItems(15) = Trim(Left(TxtPayee, IIf(InStr(1, TxtPayee, "-") - 1 < 0, Len(TxtPayee), InStr(1, TxtPayee, "-") - 1)))
    lstRec.SubItems(16) = txtDoctor
    lstRec.SubItems(17) = Logon
    lstRec.SubItems(18) = IIf(IsNumeric(A1), A1, 0)
    lstRec.SubItems(19) = IIf(IsNumeric(A2), A2, 0)
    lstRec.SubItems(20) = IIf(IsNumeric(A3), A3, 0)
    lstRec.SubItems(21) = IIf(IsNumeric(A4), A4, 0)
    lstRec.SubItems(22) = IIf(IsNumeric(A5), A5, 0)
    lstRec.SubItems(23) = IIf(IsNumeric(A6), A6, 0)
    lstRec.SubItems(24) = IIf(IsNumeric(A7), A7, 0)
    lstRec.SubItems(25) = IIf(IsNumeric(A8), A8, 0)
    lstRec.SubItems(26) = IIf(IsNumeric(A9), A9, 0)
    lstRec.SubItems(27) = IIf(IsNumeric(A10), A10, 0)
    lstRec.SubItems(28) = IIf(IsNumeric(A11), A11, 0)
    lstRec.SubItems(29) = IIf(IsNumeric(A12), A12, 0)
    lstRec.SubItems(30) = IIf(IsNumeric(A13), A13, 0)
    lstRec.SubItems(31) = IIf(IsNumeric(A14), A14, 0)
    lstRec.SubItems(32) = IIf(IsNumeric(A15), A15, 0)
    lstRec.SubItems(33) = IIf(Len(Trim(txtOtherDesc1)), Trim(txtOtherDesc1), "")
'    lstRec.SubItems(34) = IIf(Len(Trim(txtOtherDesc2)), Trim(txtOtherDesc2), "")
    lstRec.SubItems(35) = IIf(IsNumeric(A16), A16, 0)
    lstRec.SubItems(36) = IIf(IsNumeric(A17), A17, 0)
    lstRec.SubItems(37) = IIf(IsNumeric(B1), B1, 0)
    lstRec.SubItems(38) = IIf(IsNumeric(B2), B2, 0)
    lstRec.SubItems(39) = IIf(IsNumeric(B3), B3, 0)
    lstRec.SubItems(40) = IIf(IsNumeric(B4), B4, 0)
    lstRec.SubItems(41) = IIf(IsNumeric(B5), B5, 0)
    lstRec.SubItems(42) = IIf(IsNumeric(B6), B6, 0)
    lstRec.SubItems(43) = IIf(IsNumeric(B7), B7, 0)
    lstRec.SubItems(44) = Mid(cboClaimability, 1, 1)
    lstRec.SubItems(45) = Mid(cboPending, 1, 1)
    lstRec.SubItems(46) = Trim(Left(CboReason1, IIf(InStr(1, CboReason1, "-") - 1 < 0, Len(CboReason1), InStr(1, CboReason1, "-") - 1)))
    lstRec.SubItems(47) = Trim(Left(CboReason2, IIf(InStr(1, CboReason2, "-") - 1 < 0, Len(CboReason2), InStr(1, CboReason2, "-") - 1)))
    lstRec.SubItems(48) = Trim(Left(CboReason3, IIf(InStr(1, CboReason3, "-") - 1 < 0, Len(CboReason3), InStr(1, CboReason3, "-") - 1)))
    lstRec.SubItems(49) = IIf(IsNumeric(txtMC), txtMC, 0)
    lstRec.SubItems(50) = IIf(Len(Trim(TxtRemarks)), Trim(TxtRemarks), "")
    lstRec.SubItems(51) = Trim(Left(CboRepudiate1, IIf(InStr(1, CboRepudiate1, "-") - 1 < 0, Len(CboRepudiate1), InStr(1, CboRepudiate1, "-") - 1)))
    lstRec.SubItems(52) = Trim(Left(CboRepudiate2, IIf(InStr(1, CboRepudiate2, "-") - 1 < 0, Len(CboRepudiate2), InStr(1, CboRepudiate2, "-") - 1)))
    lstRec.SubItems(53) = Trim(Left(CboRepudiate3, IIf(InStr(1, CboRepudiate3, "-") - 1 < 0, Len(CboRepudiate3), InStr(1, CboRepudiate3, "-") - 1)))
    lstRec.SubItems(54) = IIf(Len(Trim(frmINVMedicine.txtMEdication1)), frmINVMedicine.txtMEdication1, "")
    lstRec.SubItems(55) = IIf(Len(Trim(frmINVMedicine.txtMEdication2)), frmINVMedicine.txtMEdication2, "")
    lstRec.SubItems(56) = IIf(Len(Trim(frmINVMedicine.txtMEdication3)), frmINVMedicine.txtMEdication3, "")
    lstRec.SubItems(57) = Trim(Left(frmINVMedicine.CboProc1, IIf(InStr(1, frmINVMedicine.CboProc1, "-") - 1 < 0, Len(frmINVMedicine.CboProc1), InStr(1, frmINVMedicine.CboProc1, "-") - 1)))
    lstRec.SubItems(58) = Trim(Left(frmINVMedicine.CboProc2, IIf(InStr(1, frmINVMedicine.CboProc2, "-") - 1 < 0, Len(frmINVMedicine.CboProc2), InStr(1, frmINVMedicine.CboProc2, "-") - 1)))
    lstRec.SubItems(59) = Trim(Left(frmINVMedicine.CboProc3, IIf(InStr(1, frmINVMedicine.CboProc3, "-") - 1 < 0, Len(frmINVMedicine.CboProc3), InStr(1, frmINVMedicine.CboProc3, "-") - 1)))
    lstRec.SubItems(60) = Trim(Left(frmINVMedicine.CboProc4, IIf(InStr(1, frmINVMedicine.CboProc4, "-") - 1 < 0, Len(frmINVMedicine.CboProc4), InStr(1, frmINVMedicine.CboProc4, "-") - 1)))
    lstRec.SubItems(61) = Trim(Left(frmINVMedicine.CboProc5, IIf(InStr(1, frmINVMedicine.CboProc5, "-") - 1 < 0, Len(frmINVMedicine.CboProc5), InStr(1, frmINVMedicine.CboProc5, "-") - 1)))
    lstRec.SubItems(62) = Trim(Left(frmINVMedicine.CboProc6, IIf(InStr(1, frmINVMedicine.CboProc6, "-") - 1 < 0, Len(frmINVMedicine.CboProc6), InStr(1, frmINVMedicine.CboProc6, "-") - 1)))
    lstRec.SubItems(63) = Trim(Left(frmINVMedicine.CboProc7, IIf(InStr(1, frmINVMedicine.CboProc7, "-") - 1 < 0, Len(frmINVMedicine.CboProc7), InStr(1, frmINVMedicine.CboProc7, "-") - 1)))
    lstRec.SubItems(64) = IIf(Len(Trim(TxtInvPayType)), Mid(TxtInvPayType, 1, 1), "")
    lstRec.SubItems(65) = TxtPayor
    lstRec.SubItems(66) = strMEdChkItem
    lstRec.SubItems(67) = IIf(isTime(txtVisitTime), txtVisitTime, "__:__")
    lstRec.SubItems(68) = strBalLimit
    lstRec.SubItems(69) = strVisitLimit
    lstRec.SubItems(70) = strCLNContract
    Call ClearEntries
    Call ClearMBMInfo
    txtMBMNo = ""
    Unload frmINVMedicine
    Call EntriesEnable(False)
    txtMBMNo.SetFocus
End Sub

Private Sub ClearMBMInfo()
    strMBMNo = ""
    strPayor = ""
    strPlanCode = ""
    strInsType = ""
    strMBMName = ""
    strTabMBMName = ""
    strEffdate = ""
    strExpDate = ""
    strPolNo = ""
    strMBMDataStatus = ""
    strPOlStatus = ""
    strPatientName = ""
    strPatientDOB = ""
    strPatientIC = ""
    strPatientID = ""
    strBalLimit = ""
    strVisitLimit = ""
    strCLNContract = ""
    strBillChkItem = "BILL ACCUTE"
    strMEdChkItem = "MEDICINE ACCUTE"
    MembershipNo = ""
    CASENumber = ""
End Sub

Private Sub RemoveRecFrListV()
On Error Resume Next
Dim ForLoop As Integer
Dim ListCount As Integer
Dim strsql As String
Dim DEL As New ADODB.Recordset

    ListCount = 0
    If Trim(lstInvoice.SelectedItem.SubItems(2)) = "" Then
        MsgBox "Please Choose a Record", vbOKOnly + vbCritical, DialogBoxTitle
        lstInvoice.SetFocus
    Else
            ListCount = lstInvoice.ListItems.count
            For ForLoop = 1 To ListCount
                If lstInvoice.ListItems(ForLoop).Selected = True Then
                      lstInvoice.ListItems.Remove (ForLoop)
                End If
            Next ForLoop

    End If
    
End Sub

Private Sub ClearInvoiceINfo()
    txtMBMNo = ""
    TxtInvPayType = "Cashless"
    CboPAYto.Text = "C - Clinic"
    Cboclinic.Text = ""
    TxtPayee = ""
    txtDateVisit = "__/__/____"
    txtDoctor = ""
End Sub

Private Sub txtVisitTime_LostFocus()
    If txtVisitTime <> "__:__" Then
        If isTime(txtVisitTime) = False Then
            MsgBox "Invalid Time Format! ", vbCritical
            txtVisitTime.SetFocus
        End If
    End If
End Sub

Private Sub ErrRpt()
Dim MBMCov As New ADODB.Recordset
Dim COVSql As String

    Set lstErrRec = lstErr.ListItems.Add(, , VMBMNo)
        lstErrRec.SubItems(1) = VCoverID
        lstErrRec.SubItems(2) = VName
        lstErrRec.SubItems(3) = VPolNo
        lstErrRec.SubItems(4) = VName
'        lstErrRec.SubItems(5) = MBM!MBMDOB
        lstErrRec.SubItems(5) = Ic
        lstErrRec.SubItems(6) = VEffDate
        lstErrRec.SubItems(7) = VExpDate
        lstErrRec.SubItems(8) = VBalLmt
        If VCoverID <> "00" Then
            COVSql = "SELECT MBMCName, MBMCICBCPPNo, MBMCDOB " & _
            " FROM MBMCoveredPersons WHERE MBMCNumber='" & VMBMNo & _
            "' AND MBMCCoverID = '" & VCoverID & "'"
            MBMCov.Open COVSql, medixmasterconn, adOpenKeyset, adLockOptimistic
            With MBMCov
                lstErrRec.SubItems(4) = !MBMCName
'                lstErrRec.SubItems(5) = !MBMCDOB
                lstErrRec.SubItems(5) = !MBMCICBCPPNo
            End With
        End If
End Sub

Private Sub VerifyDiagnosis()
Dim Diag As New ADODB.Recordset
Dim DiagSql As String

'    If VDiagIND = "0" Then
        DiagSql = "SELECT * FROM DiagNotPayable WHERE PLNCode='" & VPlan & "' AND DiagCode='" & VDiagCode & "'"
        Diag.Open DiagSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
'    ElseIf VDiagIND = "1" Then
'        DiagSql = "SELECT * FROM GenDiagNotPay WHERE PLNCode='" & VPlan & "' AND DiagCode='" & VDiagCode & "'"
'        Diag.Open DiagSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
'    End If
    If Diag.RecordCount > 0 Then
        Call ErrRpt
        lstErrRec.SubItems(9) = "Diagnosis '" & VDiagCode & "' is not payable ! "
    End If
    Diag.Close
    Set Diag = Nothing
End Sub
