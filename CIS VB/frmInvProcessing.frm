VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmINVProcessing 
   Caption         =   "Invoice Processing"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8205
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame FrameInvoice 
      Enabled         =   0   'False
      Height          =   5175
      Left            =   80
      TabIndex        =   96
      Top             =   2360
      Width           =   11775
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
         Height          =   1695
         Left            =   120
         TabIndex        =   134
         Top             =   135
         Width           =   11535
         Begin VB.OptionButton OptAccute 
            Caption         =   "Accute"
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
            TabIndex        =   157
            Top             =   120
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OptChronic 
            Caption         =   "Chronic"
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
            Left            =   1200
            TabIndex        =   156
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton OptMultiple 
            Caption         =   "Multiple"
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
            Left            =   2280
            TabIndex        =   155
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox txtInvoiceNo 
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
            TabIndex        =   19
            Top             =   470
            Width           =   1455
         End
         Begin VB.Frame FrameReason 
            Caption         =   "Reason For Seeking Treatment"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1220
            Left            =   7320
            TabIndex        =   135
            Top             =   240
            Width           =   4155
            Begin VB.ComboBox CboReason1 
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
               Left            =   240
               TabIndex        =   28
               Tag             =   "0"
               Top             =   220
               Width           =   3840
            End
            Begin VB.ComboBox CboReason2 
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
               Left            =   240
               TabIndex        =   29
               Top             =   500
               Width           =   3840
            End
            Begin VB.ComboBox CboReason3 
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
               Left            =   240
               TabIndex        =   30
               Top             =   780
               Width           =   3840
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
               Left            =   120
               TabIndex        =   138
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
               Left            =   120
               TabIndex        =   137
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
               Left            =   120
               TabIndex        =   136
               Top             =   240
               Width           =   135
            End
         End
         Begin MSMask.MaskEdBox txtInvoiceDate 
            Height          =   285
            Left            =   1320
            TabIndex        =   20
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin MSMask.MaskEdBox txtDateVisit 
            Height          =   285
            Left            =   1320
            TabIndex        =   21
            Top             =   970
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin VB.ComboBox CboPAYto 
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
            Left            =   3600
            TabIndex        =   23
            Text            =   "C - Clinic"
            Top             =   420
            Width           =   1335
         End
         Begin VB.TextBox TxtInvPayType 
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
            Left            =   6040
            TabIndex        =   24
            Text            =   "Cashless"
            Top             =   420
            Width           =   1095
         End
         Begin VB.ComboBox Cboclinic 
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
            Left            =   3600
            TabIndex        =   25
            Top             =   705
            Width           =   3540
         End
         Begin VB.ComboBox CboPayee 
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
            Left            =   3600
            TabIndex        =   26
            Top             =   990
            Width           =   3540
         End
         Begin VB.ComboBox CboNatureClm 
            Height          =   315
            Left            =   1320
            TabIndex        =   22
            Top             =   1230
            Width           =   1455
         End
         Begin VB.TextBox txtDoctor 
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
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   27
            Top             =   1270
            Width           =   3540
         End
         Begin VB.Label Label50 
            Caption         =   "Doctor"
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
            Left            =   2925
            TabIndex        =   158
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Lbl45days 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "45 Days Late"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   210
            Left            =   3600
            TabIndex        =   154
            Top             =   165
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label39 
            Caption         =   "Nature of CLM"
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
            TabIndex        =   148
            Top             =   1280
            Width           =   1095
         End
         Begin VB.Label Label11 
            Caption         =   "Cash Type"
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
            Left            =   5115
            TabIndex        =   145
            Top             =   450
            Width           =   975
         End
         Begin VB.Label Label17 
            Caption         =   "Payee"
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
            Left            =   2925
            TabIndex        =   144
            Top             =   1035
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Pay To"
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
            Left            =   2925
            TabIndex        =   143
            Top             =   435
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "Invoice Date"
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
            TabIndex        =   142
            Top             =   750
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Invoice No"
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
            TabIndex        =   141
            Top             =   500
            Width           =   975
         End
         Begin VB.Label Label14 
            Caption         =   "Clinic"
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
            Left            =   2925
            TabIndex        =   140
            Top             =   750
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Submission Date"
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
            TabIndex        =   139
            Top             =   1000
            Width           =   1215
         End
      End
      Begin VB.Frame FrameRepudiate 
         Caption         =   "Repudiation Reasons"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   130
         Top             =   4000
         Width           =   11505
         Begin VB.ComboBox CboRepudiate1 
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
            Left            =   260
            TabIndex        =   62
            Top             =   240
            Width           =   11160
         End
         Begin VB.ComboBox CboRepudiate2 
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
            Left            =   260
            TabIndex        =   63
            Top             =   480
            Width           =   11160
         End
         Begin VB.ComboBox CboRepudiate3 
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
            Left            =   260
            TabIndex        =   64
            Top             =   720
            Width           =   11160
         End
         Begin VB.Label Label48 
            Caption         =   "3"
            Height          =   255
            Left            =   120
            TabIndex        =   133
            Top             =   780
            Width           =   255
         End
         Begin VB.Label Label47 
            Caption         =   "2"
            Height          =   180
            Left            =   120
            TabIndex        =   132
            Top             =   480
            Width           =   255
         End
         Begin VB.Label Label43 
            Caption         =   "1"
            Height          =   255
            Left            =   120
            TabIndex        =   131
            Top             =   255
            Width           =   255
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
         Height          =   2175
         Left            =   8640
         TabIndex        =   126
         Top             =   1840
         Width           =   3015
         Begin VB.TextBox txtMC 
            Alignment       =   1  'Right Justify
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
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   60
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox TxtRemarks 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   840
            MaxLength       =   100
            TabIndex        =   61
            Top             =   585
            Width           =   2055
         End
         Begin VB.Label lblTotAmt 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "Total Invoice Amount Excess"
            ForeColor       =   &H00004080&
            Height          =   210
            Left            =   120
            TabIndex        =   150
            Top             =   1875
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.Label lblMedicine 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "Medication Amount Excess"
            ForeColor       =   &H00004080&
            Height          =   210
            Left            =   120
            TabIndex        =   149
            Top             =   1605
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.Label Label16 
            Caption         =   "Days"
            Height          =   255
            Left            =   2520
            TabIndex        =   129
            Top             =   270
            Width           =   375
         End
         Begin VB.Label Label15 
            Caption         =   "Medical Certificate (MC)"
            Height          =   255
            Left            =   120
            TabIndex        =   128
            Top             =   270
            Width           =   1695
         End
         Begin VB.Label Label41 
            Caption         =   "Remarks"
            Height          =   255
            Left            =   120
            TabIndex        =   127
            Top             =   600
            Width           =   855
         End
      End
      Begin VB.Frame FrameCharges 
         Caption         =   "Charges - Cover"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2175
         Left            =   120
         TabIndex        =   105
         Top             =   1840
         Width           =   6210
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
            Left            =   1960
            TabIndex        =   38
            Top             =   1220
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
            Left            =   1960
            TabIndex        =   36
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
            Left            =   1960
            TabIndex        =   34
            Top             =   740
            Width           =   240
         End
         Begin VB.TextBox TxtIndex 
            Height          =   195
            Left            =   960
            TabIndex        =   106
            Top             =   960
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox A1 
            Alignment       =   1  'Right Justify
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
            MaxLength       =   8
            TabIndex        =   31
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox A2 
            Alignment       =   1  'Right Justify
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
            MaxLength       =   8
            TabIndex        =   32
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox A3 
            Alignment       =   1  'Right Justify
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
            MaxLength       =   8
            TabIndex        =   33
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox A4 
            Alignment       =   1  'Right Justify
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
            MaxLength       =   8
            TabIndex        =   35
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox A5 
            Alignment       =   1  'Right Justify
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
            MaxLength       =   8
            TabIndex        =   37
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox A6 
            Alignment       =   1  'Right Justify
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
            Left            =   3680
            MaxLength       =   8
            TabIndex        =   39
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox A7 
            Alignment       =   1  'Right Justify
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
            Left            =   3680
            MaxLength       =   8
            TabIndex        =   40
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox A8 
            Alignment       =   1  'Right Justify
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
            Left            =   3680
            MaxLength       =   8
            TabIndex        =   41
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox A9 
            Alignment       =   1  'Right Justify
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
            Left            =   3680
            MaxLength       =   8
            TabIndex        =   42
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox A10 
            Alignment       =   1  'Right Justify
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
            Left            =   3680
            MaxLength       =   8
            TabIndex        =   43
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox A11 
            Alignment       =   1  'Right Justify
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
            Left            =   5480
            MaxLength       =   8
            TabIndex        =   44
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox A12 
            Alignment       =   1  'Right Justify
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
            Left            =   5480
            MaxLength       =   8
            TabIndex        =   45
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox A13 
            Alignment       =   1  'Right Justify
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
            Left            =   5480
            MaxLength       =   8
            TabIndex        =   46
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox A14 
            Alignment       =   1  'Right Justify
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
            Left            =   5480
            MaxLength       =   8
            TabIndex        =   47
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox A15 
            Alignment       =   1  'Right Justify
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
            Left            =   5480
            MaxLength       =   8
            TabIndex        =   48
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtOtherDesc1 
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
            MaxLength       =   100
            TabIndex        =   49
            Top             =   1440
            Width           =   2970
         End
         Begin VB.TextBox txtOtherDesc2 
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
            MaxLength       =   100
            TabIndex        =   51
            Top             =   1680
            Width           =   2970
         End
         Begin VB.TextBox A16 
            Alignment       =   1  'Right Justify
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
            Left            =   5480
            MaxLength       =   8
            TabIndex        =   50
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox A17 
            Alignment       =   1  'Right Justify
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
            Left            =   5480
            MaxLength       =   8
            TabIndex        =   52
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label Label34 
            Caption         =   "Others1 Desc"
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
            TabIndex        =   125
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label33 
            Caption         =   "T&&S / STO"
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
            Left            =   4440
            TabIndex        =   124
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label32 
            Caption         =   "I&&D"
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
            Left            =   4440
            TabIndex        =   123
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label31 
            Caption         =   "Vision"
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
            Left            =   4440
            TabIndex        =   122
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label30 
            Caption         =   "PRV Checkup"
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
            Left            =   4440
            TabIndex        =   121
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label29 
            Caption         =   "Maternity"
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
            Left            =   4440
            TabIndex        =   120
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label28 
            Caption         =   "Physio Theraphy"
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
            Left            =   2360
            TabIndex        =   119
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label27 
            Caption         =   "Traditional Med"
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
            Left            =   2360
            TabIndex        =   118
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label26 
            Caption         =   "Chinese Medicine"
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
            Left            =   2360
            TabIndex        =   117
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label25 
            Caption         =   "Dental"
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
            Left            =   2360
            TabIndex        =   116
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label23 
            Caption         =   "X-Ray / MRI"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2360
            TabIndex        =   115
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label22 
            Caption         =   "Lab"
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
            TabIndex        =   114
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label21 
            Caption         =   "Procedures"
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
            TabIndex        =   113
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label20 
            Caption         =   "Medication"
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
            TabIndex        =   112
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label19 
            Caption         =   "SP Cons Fee"
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
            TabIndex        =   111
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label18 
            Caption         =   "GP Cons Fee"
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
            TabIndex        =   110
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label13 
            Caption         =   "Others2 Desc"
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
            TabIndex        =   109
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label53 
            Caption         =   "Amount"
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
            Left            =   4440
            TabIndex        =   108
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "Amount"
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
            Left            =   4440
            TabIndex        =   107
            Top             =   1680
            Width           =   735
         End
      End
      Begin VB.Frame FrameNotCover 
         Caption         =   "Not Cover"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   6360
         TabIndex        =   97
         Top             =   1840
         Width           =   2250
         Begin VB.TextBox B1 
            Alignment       =   1  'Right Justify
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
            MaxLength       =   8
            TabIndex        =   53
            Top             =   240
            Width           =   680
         End
         Begin VB.TextBox B2 
            Alignment       =   1  'Right Justify
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
            MaxLength       =   8
            TabIndex        =   54
            Top             =   480
            Width           =   680
         End
         Begin VB.TextBox B3 
            Alignment       =   1  'Right Justify
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
            MaxLength       =   8
            TabIndex        =   55
            Top             =   720
            Width           =   680
         End
         Begin VB.TextBox B4 
            Alignment       =   1  'Right Justify
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
            MaxLength       =   8
            TabIndex        =   56
            Top             =   960
            Width           =   680
         End
         Begin VB.TextBox B5 
            Alignment       =   1  'Right Justify
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
            MaxLength       =   8
            TabIndex        =   57
            Top             =   1200
            Width           =   680
         End
         Begin VB.TextBox B6 
            Alignment       =   1  'Right Justify
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
            MaxLength       =   8
            TabIndex        =   58
            Top             =   1440
            Width           =   680
         End
         Begin VB.TextBox B7 
            Alignment       =   1  'Right Justify
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
            MaxLength       =   8
            TabIndex        =   59
            Top             =   1680
            Width           =   680
         End
         Begin VB.Label Label24 
            Caption         =   "SP Cons Fee"
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
            TabIndex        =   104
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label45 
            Caption         =   "Medication"
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
            TabIndex        =   103
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label46 
            Caption         =   "Physio Theraphy"
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
            TabIndex        =   102
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label55 
            Caption         =   "Maternity"
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
            TabIndex        =   101
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label56 
            Caption         =   "PRV Checkup"
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
            TabIndex        =   100
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label57 
            Caption         =   "Vision"
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
            TabIndex        =   99
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label59 
            Caption         =   "Dental"
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
            TabIndex        =   98
            Top             =   720
            Width           =   975
         End
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
      Left            =   80
      TabIndex        =   82
      Top             =   1200
      Width           =   11775
      Begin VB.TextBox txtMBMStatus 
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
         Left            =   3600
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtTabMBMNo 
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
         Left            =   1200
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtPolNo 
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
         Left            =   5520
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtPolStatus 
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
         Left            =   8400
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtPlanCode 
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
         Left            =   10800
         TabIndex        =   17
         Top             =   255
         Width           =   855
      End
      Begin VB.TextBox TxtInsType 
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
         Left            =   10800
         TabIndex        =   18
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox TxtPayor 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5280
         TabIndex        =   83
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtTabMBMName 
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
         Left            =   1200
         TabIndex        =   10
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txtPatientName 
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
         Left            =   1200
         TabIndex        =   11
         Top             =   720
         Width           =   3135
      End
      Begin MSMask.MaskEdBox txtEffDate 
         Height          =   285
         Left            =   5520
         TabIndex        =   14
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin MSMask.MaskEdBox txtExpDate 
         Height          =   315
         Left            =   8400
         TabIndex        =   15
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin VB.TextBox txtdob 
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
         Left            =   8400
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtPatientIC 
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
         Left            =   5520
         TabIndex        =   151
         Top             =   740
         Width           =   1815
      End
      Begin VB.TextBox TxtBalLimit 
         Height          =   285
         Left            =   10800
         TabIndex        =   161
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Balance Limit"
         Height          =   255
         Left            =   9720
         TabIndex        =   160
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblPatientIC 
         Caption         =   "Patient IC"
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
         Left            =   4520
         TabIndex        =   152
         Top             =   760
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "MBM Status"
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
         Left            =   2715
         TabIndex        =   95
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Patient Name"
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
         TabIndex        =   94
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Member Name"
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
         TabIndex        =   93
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Policy No"
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
         Left            =   4520
         TabIndex        =   92
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Eff. Date"
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
         Left            =   4520
         TabIndex        =   91
         Top             =   495
         Width           =   1095
      End
      Begin VB.Label Label36 
         Caption         =   "Exp. Date"
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
         TabIndex        =   90
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label37 
         Caption         =   "Patient DOB"
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
         TabIndex        =   89
         Top             =   765
         Width           =   975
      End
      Begin VB.Label Label38 
         Caption         =   "Pol. Status"
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
         TabIndex        =   88
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label60 
         Caption         =   "Member No"
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
         TabIndex        =   87
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
         TabIndex        =   86
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label62 
         Caption         =   "Insured Type"
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
         Left            =   9720
         TabIndex        =   85
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label64 
         Caption         =   "Plan Code"
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
         Left            =   9720
         TabIndex        =   84
         Top             =   255
         Width           =   975
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
      Height          =   800
      Left            =   80
      TabIndex        =   79
      Top             =   400
      Width           =   5655
      Begin VB.OptionButton Option2 
         Caption         =   "Case No"
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
         TabIndex        =   147
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Member No"
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
         TabIndex        =   146
         Top             =   180
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.ComboBox CboPatientID 
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
         ItemData        =   "frmInvProcessing.frx":0000
         Left            =   3240
         List            =   "frmInvProcessing.frx":0002
         TabIndex        =   2
         Top             =   180
         Width           =   1665
      End
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
         TabIndex        =   1
         Top             =   180
         Width           =   1455
      End
      Begin VB.CommandButton cmdShow 
         Appearance      =   0  'Flat
         Caption         =   "&Go"
         Default         =   -1  'True
         Height          =   300
         Left            =   4960
         TabIndex        =   4
         Top             =   180
         Width           =   570
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4960
         TabIndex        =   80
         Top             =   180
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.TextBox txtCaseNo 
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
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   440
         Width           =   1455
      End
      Begin VB.Label LblExpDate 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Expired"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3240
         TabIndex        =   159
         Top             =   520
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label40 
         Caption         =   "Cov"
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
         Left            =   2900
         TabIndex        =   153
         Top             =   220
         Width           =   375
      End
   End
   Begin VB.Frame Frame7 
      Height          =   615
      Left            =   80
      TabIndex        =   73
      Top             =   7560
      Width           =   11655
      Begin VB.CommandButton CMdChkItem 
         Caption         =   "C&hk Items"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8040
         TabIndex        =   65
         Top             =   200
         Width           =   975
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Enabled         =   0   'False
         Height          =   330
         Left            =   9000
         TabIndex        =   66
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "E&dit"
         Enabled         =   0   'False
         Height          =   330
         Left            =   9000
         TabIndex        =   74
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdClaerInv 
         Caption         =   "Clear &Inv"
         Height          =   330
         Left            =   7200
         TabIndex        =   67
         Top             =   200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear All"
         Enabled         =   0   'False
         Height          =   330
         Left            =   9720
         TabIndex        =   68
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   10560
         TabIndex        =   69
         Top             =   200
         Width           =   855
      End
      Begin VB.Label lblWs 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   78
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label63 
         Caption         =   "WS?"
         Height          =   255
         Left            =   1680
         TabIndex        =   77
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblBordx 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   76
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label58 
         Caption         =   "Bordx?"
         Height          =   255
         Left            =   360
         TabIndex        =   75
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame FrameCase 
      Enabled         =   0   'False
      Height          =   615
      Left            =   5760
      TabIndex        =   0
      Top             =   480
      Width           =   6015
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
         Left            =   3000
         TabIndex        =   6
         Text            =   "A - Accepted"
         Top             =   200
         Width           =   1215
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
         Left            =   1080
         TabIndex        =   5
         Text            =   "O - Open"
         Top             =   200
         Width           =   975
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
         Left            =   5040
         TabIndex        =   7
         Text            =   "N - No"
         Top             =   200
         Width           =   855
      End
      Begin VB.Label Label44 
         Caption         =   "Claimability"
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
         Left            =   2160
         TabIndex        =   72
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label42 
         Caption         =   "Case Status"
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
         TabIndex        =   71
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label49 
         Caption         =   "Pending"
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
         Left            =   4320
         TabIndex        =   70
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Label Label35 
      BackColor       =   &H00404040&
      Caption         =   "  Invoice Processing"
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
      Height          =   400
      Left            =   0
      TabIndex        =   81
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmINVProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim bExit As Boolean
Dim intIndex As String
Dim EditFlag As Boolean
'Dim objword As Word.Application
'Dim objdoc As Word.Document
Dim ScanCode As String
Dim DOBArr(0 To 30)
Dim NAmeArr(0 To 30)
Dim ICArr(0 To 30)
Dim BalArr(0 To 30)
Dim strBillChkItem, strMEdChkItem As String
Dim strBillChkAmt, strMEdChkAmt, IntInvAmt As Integer
Dim CASENumber, NewCaseNo, MembershipNo As String
Dim strPolNo, strMBMName, strPatientName, strMBMDataStatus, strPOlStatus As String
Dim strEffdate, strExpDate, strPlanCode, strInsType, strPatientDOB As String
Dim strPayor, strPatientID, strPatientIC, strBalLimit As String
Private Hospitals As String

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

Private Sub CboPatientID_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPatientID.Text, CboPatientID.SelStart) + Chr(KeyAscii) + Mid(CboPatientID.Text, CboPatientID.SelStart + CboPatientID.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboPatientID.ListCount - 1
       
        If NewText = LCase(Left(CboPatientID.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboPatientID.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboPatientID.Text = ValidValue
            CboPatientID.SelStart = 0
            CboPatientID.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CboPayee_Click()
    CboPayee = ChkStr(CboPayee)
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

Private Sub CMdChkItem_Click()
    Call CheckItemsAmt
End Sub

Private Sub CmdClaerInv_Click()
    Call ClearInvoice
End Sub


Private Sub CmdMedication_Click()
    frmINVMedicine.Source = 1
    frmINVMedicine.Show
End Sub

Private Sub cmdProcedures_Click()
    frmINVMedicine.Source = 2
    frmINVMedicine.Show
End Sub

Private Sub CmdLab_Click()
    frmINVMedicine.Source = 3
    frmINVMedicine.Show
End Sub

Public Sub cmdShow_Click()
    
    If Option1.Value = True Then
        EditFlag = False
        Call ClearInvoice
        CmdUpdate.Visible = True
        CmdUpdate.Caption = "&Save"
        CmdUpdate.Enabled = True
        CmdEdit.Visible = False
        Call EntriesEnable(True)
    ElseIf Option2.Value = True Then
        CmdEdit.Enabled = True
        CmdEdit.Visible = True
        CmdUpdate.Visible = False
        Call InvoiceINFO
        Call MBMInfo
    End If
    Call CheckItemsAmt
    If Trim(MembershipNo) <> "" Then
        CboPatientID = strPatientID & " - " & strPatientName
        txtMBMNo = MembershipNo
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
    End If
End Sub

Private Sub OptAccute_LostFocus()
    If OptAccute.Value = True Then
        strBillChkItem = "001"
        strMEdChkItem = "004"
    End If
End Sub

Private Sub OptChronic_LostFocus()
    If OptChronic.Value = True Then
        strBillChkItem = "002"
        strMEdChkItem = "005"
    End If
End Sub

Private Sub OptMultiple_LostFocus()
    If OptMultiple.Value = True Then
        strBillChkItem = "003"
        strMEdChkItem = "006"
    End If
End Sub

Private Sub Option1_Click()
    If Option1.Value = True Then
        Call ClearEntries
        Call ComboListing
        txtMBMNo.Enabled = True
        CboPatientID.Enabled = True
        txtCaseNo.Enabled = False
        txtMBMNo.SetFocus
        CmdUpdate.Enabled = True
        CmdUpdate.Caption = "&Save"
    End If
End Sub

Private Sub Option2_Click()
    If Option2.Value = True Then
        Call ClearEntries
        Call ComboListing
        txtMBMNo.Enabled = False
        CboPatientID.Enabled = False
        txtCaseNo.Enabled = True
        txtCaseNo.SetFocus
    End If
End Sub

Private Sub txtCaseNo_LostFocus()
    If Trim(txtCaseNo) <> "" Then
        CASENumber = txtCaseNo
    End If
End Sub

Private Sub txtDoctor_LostFocus()
    txtDoctor = ChkStr(txtDoctor)
End Sub

Private Sub txtMC_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Cboclinic_LostFocus()
    If Mid(CboPAYto, 1, 1) = "C" Then
        CboPayee.Text = Cboclinic
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
        CboPayee.Text = "MEMBER"
    End If
End Sub

Private Sub CboPatientID_LostFocus()
    If CboPatientID.ListIndex >= 0 Then
        strPatientID = Trim(Left(CboPatientID, IIf(InStr(1, CboPatientID, "-") - 1 < 0, Len(Trim(CboPatientID)), InStr(1, CboPatientID, "-") - 1)))
        strPatientDOB = DOBArr(CboPatientID.ListIndex)
        strPatientName = NAmeArr(CboPatientID.ListIndex)
        strPatientIC = ICArr(CboPatientID.ListIndex)
'        strBalLimit = BalArr(CboPatientID.ListIndex)
    End If
End Sub

Private Sub CmdClear_Click()
    Call ClearEntries
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
    TxtInvPayType = "Cashless"
    Option1.Value = True
    OptAccute.Value = True
    strBillChkItem = "001"
    strMEdChkItem = "004"
    CboPatientID = ""
    MembershipNo = ""
    CASENumber = ""
    lblTotAmt.Visible = False
    lblMedicine.Visible = False
    LblExpDate.Visible = False
    Unload frmINVMedicine
    Call EntriesEnable(False)
End Sub

Private Sub cmdEdit_Click()
    CmdUpdate.Visible = True
    CmdUpdate.Enabled = True
    CmdUpdate.Caption = "&Update"
    CmdEdit.Visible = False
    CmdEdit.Enabled = False
    Call EntriesEnable(True)
    EditFlag = True
    txtInvoiceNo.SetFocus
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    frmCasList.Source = 6
    frmCasList.Show vbModal
End Sub

Private Sub CmdUpdate_Click()
Dim CISCasRec As New ADODB.Recordset
Dim CISCasSql As String
Dim InvChk As Boolean
Dim DupInvoice

    If Trim(txtMBMNo) = "" Then
        MsgBox "Please keyin Member No or Case No. ! ", vbCritical
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
    ElseIf Trim(CboPayee) = "" Then
        MsgBox "Please keyin the person/Clinic to be paid! ", vbCritical
        CboPayee.SetFocus
    ElseIf Mid(CboPAYto, 1, 1) = "C" And Trim(Left(Cboclinic, IIf(InStr(1, Cboclinic, "-") - 1 < 0, Len(Cboclinic), InStr(1, Cboclinic, "-") - 1))) <> Trim(Left(CboPayee, IIf(InStr(1, CboPayee, "-") - 1 < 0, Len(CboPayee), InStr(1, CboPayee, "-") - 1))) Then
        MsgBox "Clinic code is different with the Payee code ", vbCritical
        Cboclinic.SetFocus
    Else
        If Option1.Value = True Then
            CISCasSql = "SELECT * FROM View_Clm_Invoice where MBMNumber = '" & Trim(txtMBMNo) & _
                  "' AND COVERID ='" & Mid(CboPatientID, 1, 2) & "' AND INVDate = '" & Format(txtInvoiceDate, "mm/dd/yyyy") & "'"
            If Trim(txtInvoiceNo) <> "" Then
                InvChk = True
                CISCasSql = CISCasSql & " AND InvNo = '" & Trim(txtInvoiceNo) & "'"
            End If
        ElseIf Option2.Value = True Then
            CISCasSql = "SELECT * FROM View_Clm_Invoice where CASNumber = '" & Trim(CASENumber) & "'"
        End If
        CISCasRec.Open CISCasSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
        If CISCasRec.RecordCount And EditFlag = False Then
            If InvChk = False Then
                DupInvoice = MsgBox("Invoice has been entered !, Press 'Yes' to Proceed, 'No' to Cancel", vbYesNo + vbExclamation)
                If DupInvoice = vbYes Then
                    Call SaveRecord
                ElseIf DupInvoice = vbNo Then
                    Call EntriesEnable(False)
                End If
            Else
                MsgBox "Record Duplicated, Same Invoice has been entered !", vbCritical, "Error"
            End If
            
            Exit Sub
        Else
            Call SaveRecord
        End If
        CmdUpdate.Visible = False
        If Option1.Value = True Then
             CmdUpdate.Visible = True
             CmdUpdate.Enabled = False
             CmdEdit.Visible = False
        End If
        CISCasRec.Close
        Set CISCasRec = Nothing
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
    CmdClear.Enabled = True
    CmdUpdate.Visible = False
    strBillChkItem = "001"
    strMEdChkItem = "004"
    intIndex = 0
End Sub

Private Sub ComboListing()
Dim ClinicRec As New ADODB.Recordset
Dim RepRec As New ADODB.Recordset
Dim TreatRec As New ADODB.Recordset
Dim NatureCLM As New ADODB.Recordset
        
'    option1.Value = 1
    TxtInvPayType = "Cashless"
    Cboclinic.Clear
    CboPayee.Clear
    ClinicRec.Open "SELECT CLNCode, CLNName FROM CLN", medixmasterconn, adOpenKeyset, adLockOptimistic
    Do Until ClinicRec.EOF
        Cboclinic.AddItem ClinicRec!CLNCode & " - " & ClinicRec!CLNName
        CboPayee.AddItem ClinicRec!CLNCode & " - " & ClinicRec!CLNName
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
            ctl.Text = ""
        ElseIf TypeOf ctl Is ComboBox Then
        
        ElseIf TypeOf ctl Is ListView Then
            ctl.ListItems.Clear
        ElseIf TypeOf ctl Is MaskEdBox Then
            ctl.mask = ""
            ctl.Text = ""
            ctl.mask = "##/##/####"
        End If
    Next ctl
    TxtInvPayType = "Cashless"
    txtINvDate = ""
    Cboclinic.Text = ""
    CboPayee.Text = ""
    CboPatientID.Text = ""
    CboNatureClm.Text = ""
    CboReason1.Text = ""
    CboReason2.Text = ""
    CboReason3.Text = ""
    CboRepudiate1.Text = ""
    CboRepudiate2.Text = ""
    CboRepudiate3.Text = ""
    CmdUpdate.Visible = False
    CmdEdit.Visible = True
    CmdEdit.Enabled = False
    lblBordx = ""
    lblVersion = ""
    lblLastVersion = ""
    Lbl45days.Visible = False
    lblWs = ""
    IntInvAmt = 0
End Sub

Private Sub txtInvoiceNo_LostFocus()
    txtInvoiceNo = ChkStr(txtInvoiceNo)
End Sub

Private Sub txtMBMNo_LostFocus()
    If Trim(txtMBMNo) <> "" Then
        Call ClearInvoice
        txtMBMNo = ChkStr(txtMBMNo)
        MembershipNo = txtMBMNo
        Call MBMInfo
    End If
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
        ElseIf CDate(txtInvoiceDate) < CDate(strEffdate) Then
            MsgBox "Invoice date must be greater than Effective date ! ", vbCritical
            txtInvoiceDate.SetFocus
        ElseIf CDate(txtInvoiceDate) > CDate(strExpDate) + 14 Then
            MsgBox "Invoice date cannot greater than Expiry date + 14 days ! ", vbCritical
            txtInvoiceDate.SetFocus
        Else
            If CDate(Date) - CDate(txtInvoiceDate) > 45 Then
                Lbl45days.Visible = True
            Else
                Lbl45days.Visible = False
            End If
        '    txtDateVisit = txtInvoiceDate
        End If
    End If
End Sub

Private Sub MBMInfo()
Dim MBMRec As New ADODB.Recordset
Dim MBMSql As String
Dim COVRec As New ADODB.Recordset
Dim COVSql As String
Dim Cnt As Integer

    If Len(Trim(MembershipNo)) Then
        MBMSql = "SELECT MBMNumber, MBMName, MBMPolicyNo, MBMICBcPp, MBMPayorEffDate, MBMPayorExpDate, PAYCode, MBMCLNBalLmt, " & _
                " MBMPolicyNo,MBMStatus, MBMCOVID,INSCode, PLNCode, MBMDOB,MBMDataStatus FROM MBM WHERE MBMNumber='" & Trim(MembershipNo) & " ' AND MBMClinical='Y' AND MBMDataStatus <> 'D' AND MBMStatus <> 'D'"
        MBMRec.Open MBMSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        CboPatientID.Clear
        If MBMRec.RecordCount <= 0 Then
            MsgBox "Membership No not found ! ", vbOKOnly + vbCritical, DialogBoxTitle
            Call ClearEntries
            txtMBMNo = MembershipNo
            MembershipNo = ""
            txtMBMNo.SetFocus
        Else
            CboPatientID.AddItem MBMRec!MBMCOVID & " - " & MBMRec!MBMName
            DOBArr(0) = MBMRec!MBMDOB
            NAmeArr(0) = MBMRec!MBMName
            ICArr(0) = MBMRec!MBMICBcPp
            BalArr(0) = MBMRec!MBMCLNBalLmt
            strMBMNo = MBMRec!MBMNumber
            strPayor = MBMRec!PAYCode
            strPlanCode = MBMRec!PLNCode
            strInsType = MBMRec!INSCode
            strMBMName = MBMRec!MBMName
            strTabMBMName = MBMRec!MBMName
            strEffdate = MBMRec!MBMPayorEffDate
            strExpDate = MBMRec!MBMPayorExpDate
            strPolNo = MBMRec!MBMPolicyNo
            strMBMDataStatus = MBMRec!MBMDataStatus
            strPOlStatus = MBMRec!MBMStatus
            strPatientName = MBMRec!MBMName
            strPatientDOB = MBMRec!MBMDOB
            strPatientIC = MBMRec!MBMICBcPp
            strBalLimit = MBMRec!MBMCLNBalLmt
            If Trim(MBMRec!INSCode) = "I" Or Trim(MBMRec!INSCode) = "G" Then
                CboPatientID.Text = MBMRec!MBMCOVID & " - " & MBMRec!MBMName
                CboPatientID.Enabled = False
            Else
                CboPatientID.Enabled = True
            End If
            If Option1.Value = True Then
                strPatientID = MBMRec!MBMCOVID
                COVSql = "SELECT MBMCName, MBMCCoverID, MBMCICBCPPNo, MBMCDOB, MBMCAnnualLimit, MBMCAvailableLimit FROM MBMCoveredPersons WHERE MBMCNumber='" & Trim(MembershipNo) & " '"
            ElseIf Option2.Value = True Then
                COVSql = "SELECT MBMCName, MBMCCoverID, MBMCICBCPPNo, MBMCDOB, MBMCAnnualLimit, MBMCAvailableLimit FROM MBMCoveredPersons WHERE MBMCNumber='" & Trim(MembershipNo) & "' AND MBMCCoverID = '" & Trim(strPatientID) & "'"
            End If
            COVRec.Open COVSql, medixmasterconn, adOpenKeyset, adLockOptimistic
            If COVRec.RecordCount Then
                If Option2.Value = True Then
                    strPatientName = COVRec!MBMCName
                    strPatientDOB = COVRec!MBMCDOB
                    strPatientIC = COVRec!MBMCICBCPPNo
                    strPatientID = COVRec!MBMCCoverID
'                    If IsNumeric(MBMCAnnualLimit) And MBMCAnnualLimit > 0 Then
'                        strBalLimit = COVRec!MBMCAvailableLimit
'                    End If
                End If
                Do Until COVRec.EOF
                    Cnt = Cnt + 1
                    CboPatientID.AddItem COVRec!MBMCCoverID & " - " & COVRec!MBMCName
                    DOBArr(Cnt) = COVRec!MBMCDOB
                    NAmeArr(Cnt) = COVRec!MBMCName
                    ICArr(Cnt) = COVRec!MBMCICBCPPNo
'                    If IsNumeric(COVRec!MBMCAnnualLimit) And COVRec!MBMCAnnualLimit > 0 Then
'                        BalArr(Cnt) = COVRec!MBMCAvailableLimit
'                    Else
'                        BalArr(Cnt) = strBalLimit = MBMRec!MBMAvailableLimit
'                    End If
                    COVRec.MoveNext
                Loop
            End If
            COVRec.Close
            Set COVRec = Nothing
            If Date - CDate(strExpDate) > 14 Then
                LblExpDate.Visible = True
            Else
                LblExpDate.Visible = False
            End If
            CboPatientID = strPatientID & " - " & strPatientName
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

    GenerateCISCaseNo = PayorCode & "C" & CISSeqNo
    CISCAQ.Close
    Set CISCAQ = Nothing
End Function

Private Sub InvoiceINFO()
Dim CISCAS1 As New ADODB.Recordset
Dim CASSql, CLMInvSql As String
Dim rst2 As New ADODB.Recordset
Dim tempMBM As String
Dim RecList As ListItem
Dim ClmInvACt As New ADODB.Recordset
Dim CLmInvACTSql As String
Dim strsql As String
Dim SSS As Integer
    CASENumber = txtCaseNo
    If Option2.Value = True Then
        strsql = " WHERE CASNumber='" & Trim(CASENumber) & "'"
    ElseIf Option1.Value = True Then
'        strsql = " WHERE MBMNumber='" & MembershipNo & "' AND COVERID='" & Mid(CboPatientID, 1, 2) & "' AND DateVisit='" & Format(txtDateVisit, "MM/DD/YYYY") & "'"
        strsql = " WHERE MBMNumber='" & MembershipNo & "' AND COVERID='" & Mid(CboPatientID, 1, 2) & "'"
    End If
    CASSql = "SELECT MBMNumber, * FROM View_Clm_Invoice " & strsql
    CISCAS1.Open CASSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    If CISCAS1.RecordCount <= 0 Then
        MsgBox "Not Record Found ! ", vbOKOnly + vbCritical, DialogBoxTitle
        Call ClearEntries
        txtCaseNo = CASENumber
        MembershipNo = ""
        CASENumber = ""
        txtCaseNo.SetFocus
    Else
        With CISCAS1
            MembershipNo = !MBMNumber
            CASENumber = !CASNumber
            strPatientID = !CoverID
            TxtIndex = !InvINdex
            txtInvoiceNo = !INVNo
            txtInvoiceDate = !InvDate
            txtDateVisit = !DateVisit
            Cboclinic = IIf(Len(!ClinicCode), !ClinicCode, "")
            CboNatureClm = IIf(Len(!NatureCLM), !NatureCLM, "")
            OptAccute.Value = False
            OptChronic.Value = False
            OptChronic.Value = False
            If !ChkItemCat = "001" Then
                OptAccute.Value = True
            ElseIf !ChkItemCat = "002" Then
                OptChronic.Value = True
            ElseIf !ChkItemCat = "003" Then
                OptChronic.Value = True
            End If
            TxtRemarks = IIf(Len(!Remarks), !Remarks, "")
            CboPAYto = !INVPayTo
            If Mid(CboPAYto, 1, 2) = "C" Then
                TxtInvPayType = "Cashless"
            Else
                TxtInvPayType = "Reimbursement"
            End If
            CboPayee = !INVPayee
            A1 = IIf(Len(!CLMGPConsAmt), !CLMGPConsAmt, 0)
            A2 = IIf(Len(!CLMSPConsAmt), !CLMSPConsAmt, 0)
            A3 = IIf(Len(!CLMMedicine), !CLMMedicine, 0)
            A4 = IIf(Len(!CLMProcedures), !CLMProcedures, 0)
            A5 = IIf(Len(!CLMLab), !CLMLab, 0)
            A6 = IIf(Len(!CLMXray), !CLMXray, 0)
            A7 = IIf(Len(!CLMDental), !CLMDental, 0)
            A8 = IIf(Len(!CLMChineseMed), !CLMChineseMed, 0)
            A9 = IIf(Len(!CLMTraditionalMEd), !CLMTraditionalMEd, 0)
            A10 = IIf(Len(!CLMphysioTheraphy), !CLMphysioTheraphy, 0)
            A11 = IIf(Len(!CLMAccident), !CLMAccident, 0)
            A12 = IIf(Len(!CLMPRVCheckup), !CLMPRVCheckup, 0)
            A13 = IIf(Len(!CLMVision), !CLMVision, 0)
            A14 = IIf(Len(!CLMID), !CLMID, 0)
            A15 = IIf(Len(!CLMTSSTO), !CLMTSSTO, 0)
            A16 = IIf(Len(!CLMOTHERS1), !CLMOTHERS1, 0)
            A17 = IIf(Len(!CLMOTHERS2), !CLMOTHERS2, 0)
            
            B1 = IIf(Len(!SPFee), !SPFee, 0)
            B2 = IIf(Len(!Medicine), !Medicine, 0)
            B3 = IIf(Len(!Dental), !Dental, 0)
            B4 = IIf(Len(!Physio), !Physio, 0)
            B5 = IIf(Len(!Maternity), !Maternity, 0)
            B6 = IIf(Len(!PRVCheckup), !PRVCheckup, 0)
            B7 = IIf(Len(!Vision), !Vision, 0)
            
            If !Claimability = "A" Then
                cboClaimability.Text = "A - Accepted"
            ElseIf !Claimability = "R" Then
                cboClaimability.Text = "R - Rejected"
            ElseIf !Claimability = "C" Then
                cboClaimability.Text = "C - Cancelled"
            End If
            If !Pending = "Y" Then
                cboPending.Text = "Y - Yes"
            ElseIf !Pending = "N" Then
                cboPending.Text = "N - No"
            End If
            CboReason1.Text = IIf(Len(!CASTreatReason1), !CASTreatReason1, "")
            CboReason2.Text = IIf(Len(!CASTreatReason2), !CASTreatReason2, "")
            CboReason3.Text = IIf(Len(!CASTreatReason3), !CASTreatReason3, "")
            txtMC = !MCDays
            CboRepudiate1.Text = IIf(Len(!CASREPCode1), !CASREPCode1, "")
            CboRepudiate2.Text = IIf(Len(!CASREPCode2), !CASREPCode2, "")
            CboRepudiate3.Text = IIf(Len(!CASREPCode3), !CASREPCode3, "")
            lblBordx = !INVBordxed
            lblWs = !INVWSStatus
            
'  frmInvMedicine
            frmINVMedicine.CboDrug1.Text = IIf(Len(!CLMDrug1), !CLMDrug1, "")
            frmINVMedicine.CboDrug2.Text = IIf(Len(!CLMDrug2), !CLMDrug2, "")
            frmINVMedicine.CboDrug3.Text = IIf(Len(!CLMDrug3), !CLMDrug3, "")
            frmINVMedicine.CboDrug4.Text = IIf(Len(!CLMDrug4), !CLMDrug4, "")
            frmINVMedicine.CboDrug5.Text = IIf(Len(!CLMDrug5), !CLMDrug5, "")
            frmINVMedicine.CboDrug6.Text = IIf(Len(!CLMDrug6), !CLMDrug6, "")
            frmINVMedicine.CboDrug7.Text = IIf(Len(!CLMDrug7), !CLMDrug7, "")
            frmINVMedicine.txtDrugAmt1 = IIf(IsNumeric(!CLMDrugAmt1), !CLMDrugAmt1, 0)
            frmINVMedicine.txtDrugAmt2 = IIf(IsNumeric(!CLMDrugAmt2), !CLMDrugAmt2, 0)
            frmINVMedicine.txtDrugAmt3 = IIf(IsNumeric(!CLMDrugAmt3), !CLMDrugAmt3, 0)
            frmINVMedicine.txtDrugAmt4 = IIf(IsNumeric(!CLMDrugAmt4), !CLMDrugAmt4, 0)
            frmINVMedicine.txtDrugAmt5 = IIf(IsNumeric(!CLMDrugAmt5), !CLMDrugAmt5, 0)
            frmINVMedicine.txtDrugAmt6 = IIf(IsNumeric(!CLMDrugAmt6), !CLMDrugAmt6, 0)
            frmINVMedicine.txtDrugAmt7 = IIf(IsNumeric(!CLMDrugAmt7), !CLMDrugAmt7, 0)
            frmINVMedicine.CboProc1.Text = IIf(Len(!CLMProc1), !CLMProc1, "")
            frmINVMedicine.CboProc2.Text = IIf(Len(!CLMProc2), !CLMProc2, "")
            frmINVMedicine.CboProc3.Text = IIf(Len(!CLMProc3), !CLMProc3, "")
            frmINVMedicine.txtProcAmt1 = IIf(IsNumeric(!CLMProcAmt1), !CLMProcAmt1, 0)
            frmINVMedicine.txtProcAmt2 = IIf(IsNumeric(!CLMProcAmt2), !CLMProcAmt2, 0)
            frmINVMedicine.txtProcAmt3 = IIf(IsNumeric(!CLMProcAmt3), !CLMProcAmt3, 0)
            frmINVMedicine.TxtLabSAmt1 = IIf(IsNumeric(!CLMLabAmt1), !CLMLabAmt1, 0)
            frmINVMedicine.TxtLabSAmt2 = IIf(IsNumeric(!CLMLabAmt2), !CLMLabAmt2, 0)
            frmINVMedicine.TxtLabSAmt3 = IIf(IsNumeric(!CLMLabAmt3), !CLMLabAmt3, 0)
            frmINVMedicine.TxtLabSAmt4 = IIf(IsNumeric(!CLMLabAmt4), !CLMLabAmt4, 0)
            frmINVMedicine.TxtLabSAmt5 = IIf(IsNumeric(!CLMLabAmt5), !CLMLabAmt5, 0)
            
        End With
        Call EntriesEnable(False)
        If Option1.Value = True Then
            CmdEdit.Enabled = False
        End If
    End If
    CISCAS1.Close
    Set CISCAS1 = Nothing
End Sub

Private Sub SaveRecord()
Dim RecordSaved
Dim startTrans As Boolean
    
    Call CheckItemsAmt
    RecordSaved = MsgBox("Are you sure", vbYesNo + vbExclamation)
    If RecordSaved = vbYes Then
            MediConnCISWS.beginTrans
            startTrans = True
               
            Call getTotal
            Call SaveCase
            Call SaveInvoice
            Call saveActual
            Call saveActualOthers
'            Call SaveReason
            Call saveNotCover
        
            MediConnCISWS.CommitTrans
            startTrans = False
            If Option1.Value = True Then
                txtCaseNo = NewCaseNo
                MsgBox "Case Number : " & NewCaseNo, vbCritical
            ElseIf Option2.Value = True Then
                MsgBox "Record Updated ! ", vbOKOnly
            End If
            
            CmdClear.Enabled = True
'            Call ClearEntries
            If Option2.Value = True Then
                Call InvoiceINFO
                Call MBMInfo
            End If
            Call EntriesEnable(False)
            CmdEdit.Visible = True
            CmdEdit.Enabled = True
            EditFlag = False
    End If
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
    
    If EditFlag = False Then
        ClaimInvoice.Open "Select INVINdex , * from ClaimInvoice", MediConnCISWS, adOpenKeyset, adLockOptimistic
        ClaimInvoice.AddNew
        With ClaimInvoice
            !INVBordxed = 0
            !INVWSStatus = 0
            !INVEnteredDate = Date
            !INVEnteredBy = Logon
        End With
    Else
        strsql = "Select INVINdex , * from ClaimInvoice" & _
            " where INVIndex ='" & Trim(TxtIndex) & "'"
        ClaimInvoice.Open strsql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    End If
    
    With ClaimInvoice
        !INVNo = txtInvoiceNo
        !CASNumber = NewCaseNo
        !InvDate = txtInvoiceDate
        !MCDays = IIf(IsNumeric(txtMC), txtMC, 0)
        !INVLastUpdateBy = Logon
        !INVLastUpdateDate = Date
        If IsNumeric(IntInvAmt) = True Then
            !INVAmount = IntInvAmt
        Else
            !INVAmount = 0
        End If
        !INVPayee = Trim(Left(CboPayee, IIf(InStr(1, CboPayee, "-") - 1 < 0, Len(CboPayee), InStr(1, CboPayee, "-") - 1)))
        !INVPAYType = Mid(TxtInvPayType, 1, 1)
        !INVPayTo = Mid(CboPAYto, 1, 1)

    End With
    ClaimInvoice.Update
    
    If EditFlag = False Then
        Dim ClaimInvoiceLastIndex  As New ADODB.Recordset
        ClaimInvoiceLastIndex.Open "Select INVINdex from ClaimInvoice Order by INVIndex DESC", MediConnCISWS, adOpenKeyset, adLockOptimistic
        If Not ClaimInvoiceLastIndex.EOF Then
            intIndex = ClaimInvoiceLastIndex!InvINdex
        Else
            intIndex = 1
        End If
        ClaimInvoiceLastIndex.Close
        Set ClaimInvoiceLastIndex = Nothing
    Else
        intIndex = TxtIndex
    End If
    
    ClaimInvoice.Close
    Set ClaimInvoice = Nothing
End Sub
Private Sub saveActual()
    Dim ClaimActual As New ADODB.Recordset
    Dim strsql As String
 
    If EditFlag = False Then
        ClaimActual.Open "Select INVINdex , * from ClaimInvoiceActual", MediConnCISWS, adOpenKeyset, adLockOptimistic
        ClaimActual.AddNew
        ClaimActual!InvINdex = intIndex
    Else
        strsql = "Select INVINdex , * from ClaimInvoiceActual where INVIndex =" & intIndex
        ClaimActual.Open strsql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    End If
    
    With ClaimActual
        !CLMGPConsAmt = IIf(IsNumeric(A1), A1, 0)
        !CLMSPConsAmt = IIf(IsNumeric(A2), A2, 0)
        !CLMMedicine = IIf(IsNumeric(A3), A3, 0)
        !CLMProcedures = IIf(IsNumeric(A4), A4, 0)
        !CLMLab = IIf(IsNumeric(A5), A5, 0)
        !CLMXray = IIf(IsNumeric(A6), A6, 0)
        !CLMDental = IIf(IsNumeric(A7), A7, 0)
        !CLMChineseMed = IIf(IsNumeric(A8), A8, 0)
        !CLMTraditionalMEd = IIf(IsNumeric(A9), A9, 0)
        !CLMphysioTheraphy = IIf(IsNumeric(A10), A10, 0)
        !CLMAccident = IIf(IsNumeric(A11), A11, 0)
        !CLMPRVCheckup = IIf(IsNumeric(A12), A12, 0)
        !CLMVision = IIf(IsNumeric(A13), A13, 0)
        !CLMID = IIf(IsNumeric(A14), A14, 0)
        !CLMTSSTO = IIf(IsNumeric(A15), A15, 0)
        !CLMOTHERS1 = IIf(IsNumeric(A16), A16, 0)
        !CLMOTHERS2 = IIf(IsNumeric(A17), A17, 0)
        !CLMOthers1Desc = txtOtherDesc1
        !CLMOthers2Desc = txtOtherDesc2
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
        !SPFee = IIf(Len(B1), B1, 0)
        !Medicine = IIf(Len(B2), B2, 0)
        !Dental = IIf(Len(B3), B3, 0)
        !Physio = IIf(Len(B4), B4, 0)
        !Maternity = IIf(Len(B5), B5, 0)
        !PRVCheckup = IIf(Len(B6), B6, 0)
        !Vision = IIf(Len(B7), B7, 0)
    End With
    ClaimNotCover.Update
    
    ClaimNotCover.Close
    Set ClaimNotCover = Nothing
End Sub

Private Sub SaveCase()
Dim CISCAS  As New ADODB.Recordset
Dim CISSql As String
    
    Hospitals = ""
    
    If Option1.Value = True Then
        CISSql = "SELECT * FROM CISCAS "
        CISCAS.Open CISSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        CISCAS.AddNew
        NewCaseNo = GenerateCISCaseNo(TxtPayor)
    ElseIf Option2.Value = True Then
        CISSql = "SELECT * FROM CISCAS where CASNumber = '" & Trim(CASENumber) & "'"
        CISCAS.Open CISSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        NewCaseNo = CASENumber
    End If
    
    With CISCAS
        If Option1.Value = True Then
            !CASNumber = NewCaseNo
            !MBMNumber = txtMBMNo
            !CoverID = Mid(CboPatientID, 1, 2)
            !CASStatus = "O"
            !Claimability = "A"
            !Pending = "N"
            !REgisterDate = Date
            !REgisterBy = Logon
        End If
        !ChkItemCat = strBillChkItem
        !Remarks = TxtRemarks
        !ClinicCode = Trim(Left(Cboclinic, IIf(InStr(1, Cboclinic, "-") - 1 < 0, Len(Cboclinic), InStr(1, Cboclinic, "-") - 1)))
        !DateVisit = txtDateVisit
        !CASStatus = Mid(cboStatus, 1, 1)
        !Claimability = Mid(cboClaimability, 1, 1)
        !Pending = Mid(cboPending, 1, 1)
        !CASDoctor = Trim(txtDoctor)
        !NatureCLM = Trim(Left(CboNatureClm, IIf(InStr(1, CboNatureClm, "-") - 1 < 0, Len(CboNatureClm), InStr(1, CboNatureClm, "-") - 1)))
        !CASTreatReason1 = Trim(Left(CboReason1, IIf(InStr(1, CboReason1, "-") - 1 < 0, Len(CboReason1), InStr(1, CboReason1, "-") - 1)))
        !CASTreatReason2 = Trim(Left(CboReason2, IIf(InStr(1, CboReason2, "-") - 1 < 0, Len(CboReason2), InStr(1, CboReason2, "-") - 1)))
        !CASTreatReason3 = Trim(Left(CboReason3, IIf(InStr(1, CboReason3, "-") - 1 < 0, Len(CboReason3), InStr(1, CboReason3, "-") - 1)))
        If Mid(cboClaimability, 1, 1) = "R" Then
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
            !CASREPCode1 = Trim(Mid(ChkStr(CboRepudiate1), 1, 7))
            !CASREPCode2 = Trim(Mid(ChkStr(CboRepudiate2), 1, 7))
            !CASREPCode3 = Trim(Mid(ChkStr(CboRepudiate3), 1, 7))
            
            !CASREPIssuedBy = Logon
            !CASREPDate = Date
        End With
        CASREP.Update
       
    CASREP.Close
    Set CASREP = Nothing
End Sub

Private Sub EntriesEnable(status As Boolean)
    FrameInvoice.Enabled = status
    FrameCase.Enabled = status
    If Mid(cboClaimability, 1, 1) = "R" Then
        FrameRepudiate.Enabled = True
    End If
End Sub

Private Sub ClearInvoice()
On Error Resume Next
    Dim ctl As Control
    For Each ctl In Me.Controls
        If ctl.TabIndex >= 19 And ctl.TabIndex <= 64 Then
            If TypeOf ctl Is TextBox Then
                If ctl.Name <> "txtCaseNo" And ctl.Name <> "txtMBMNo" Then
                    ctl.Text = ""
                End If
            ElseIf TypeOf ctl Is ComboBox Then
                 'If ctl.Name = "CboPayType" Then
                 '   CboPayType.ListIndex = 1
                'Else
                '    ctl.Text = ""
                'End If
            ElseIf TypeOf ctl Is MaskEdBox Then
                ctl.mask = ""
                ctl.Text = ""
                ctl.mask = "##/##/####"
            End If
        End If
    Next ctl
    TxtInvPayType = "Cashless"
    CboPAYto.Text = "C - Clinic"
    CboNatureClm.Text = ""
    CboPayee.Text = ""
    Cboclinic.Text = ""
    CboReason1.Text = ""
    CboReason2.Text = ""
    CboReason3.Text = ""
    CboRepudiate1.Text = ""
    CboRepudiate2.Text = ""
    CboRepudiate3.Text = ""
    strBillChkItem = "001"
    strMEdChkItem = "004"
    CmdUpdate.Visible = False
    CmdEdit.Visible = True
    CmdEdit.Enabled = False
End Sub

Private Sub txtOtherDesc1_LostFocus()
    txtOtherDesc1 = ChkStr(txtOtherDesc1)
End Sub

Private Sub txtOtherDesc2_LostFocus()
    txtOtherDesc2 = ChkStr(txtOtherDesc2)
End Sub

Private Sub TxtRemarks_LostFocus()
    TxtRemarks = ChkStr(TxtRemarks)
End Sub

Private Sub CheckItemsAmt()
Dim ChkSql As String
Dim ChKItemRec As New ADODB.Recordset

    ChkSql = "SELECT Catcode, Amount FROM CheckItems WHERE CatCode = '" & strBillChkItem & _
            "' OR CatCode ='" & strMEdChkItem & "'"
    ChKItemRec.Open ChkSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    Do Until ChKItemRec.EOF
        If ChKItemRec!CatCode >= "001" And ChKItemRec!CatCode <= "003" Then
            strBillChkAmt = ChKItemRec!Amount
        Else
            strMEdChkAmt = ChKItemRec!Amount
        End If
    ChKItemRec.MoveNext
    Loop
    ChKItemRec.Close
    Set ChKItemRec = Nothing

    If IsNumeric(A3) Then
        If CDbl(A3) > CDbl(strMEdChkAmt) Then
            lblMedicine.Visible = True
        Else
            lblMedicine.Visible = False
        End If
    End If
    Call getTotal
    If CDbl(IntInvAmt) > CDbl(strBillChkAmt) Then
        lblTotAmt.Visible = True
    Else
        lblTotAmt.Visible = False
    End If
    If IsDate(txtInvoiceDate) Then
        If (Format(Date, "yyyymmdd") - Format(txtInvoiceDate, "yyyymmdd")) > 45 Then
            Lbl45days.Visible = True
        Else
            Lbl45days.Visible = False
        End If
    End If
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
        !CLMDrug1 = IIf(Len(frmINVMedicine.CboDrug1), Mid(frmINVMedicine.CboDrug1, 1, 3), "")
        !CLMDrug2 = IIf(Len(frmINVMedicine.CboDrug2), Mid(frmINVMedicine.CboDrug2, 1, 3), "")
        !CLMDrug3 = IIf(Len(frmINVMedicine.CboDrug3), Mid(frmINVMedicine.CboDrug3, 1, 3), "")
        !CLMDrug4 = IIf(Len(frmINVMedicine.CboDrug4), Mid(frmINVMedicine.CboDrug4, 1, 3), "")
        !CLMDrug5 = IIf(Len(frmINVMedicine.CboDrug5), Mid(frmINVMedicine.CboDrug5, 1, 3), "")
        !CLMDrug6 = IIf(Len(frmINVMedicine.CboDrug6), Mid(frmINVMedicine.CboDrug6, 1, 3), "")
        !CLMDrug7 = IIf(Len(frmINVMedicine.CboDrug7), Mid(frmINVMedicine.CboDrug7, 1, 3), "")
        !CLMDrugAmt1 = IIf(IsNumeric(frmINVMedicine.txtDrugAmt1), frmINVMedicine.txtDrugAmt1, 0)
        !CLMDrugAmt2 = IIf(IsNumeric(frmINVMedicine.txtDrugAmt2), frmINVMedicine.txtDrugAmt2, 0)
        !CLMDrugAmt3 = IIf(IsNumeric(frmINVMedicine.txtDrugAmt3), frmINVMedicine.txtDrugAmt3, 0)
        !CLMDrugAmt4 = IIf(IsNumeric(frmINVMedicine.txtDrugAmt4), frmINVMedicine.txtDrugAmt4, 0)
        !CLMDrugAmt5 = IIf(IsNumeric(frmINVMedicine.txtDrugAmt5), frmINVMedicine.txtDrugAmt5, 0)
        !CLMDrugAmt6 = IIf(IsNumeric(frmINVMedicine.txtDrugAmt6), frmINVMedicine.txtDrugAmt6, 0)
        !CLMDrugAmt7 = IIf(IsNumeric(frmINVMedicine.txtDrugAmt7), frmINVMedicine.txtDrugAmt7, 0)
        !CLMProc1 = IIf(Len(frmINVMedicine.CboProc1), Mid(frmINVMedicine.CboProc1, 1, 3), "")
        !CLMProc2 = IIf(Len(frmINVMedicine.CboProc2), Mid(frmINVMedicine.CboProc2, 1, 3), "")
        !CLMProc3 = IIf(Len(frmINVMedicine.CboProc3), Mid(frmINVMedicine.CboProc3, 1, 3), "")
        !CLMProcAmt1 = IIf(IsNumeric(frmINVMedicine.txtProcAmt1), frmINVMedicine.txtProcAmt1, 0)
        !CLMProcAmt2 = IIf(IsNumeric(frmINVMedicine.txtProcAmt2), frmINVMedicine.txtProcAmt2, 0)
        !CLMProcAmt3 = IIf(IsNumeric(frmINVMedicine.txtProcAmt3), frmINVMedicine.txtProcAmt3, 0)
        !CLMLabAmt1 = IIf(IsNumeric(frmINVMedicine.TxtLabSAmt1), frmINVMedicine.TxtLabSAmt1, 0)
        !CLMLabAmt2 = IIf(IsNumeric(frmINVMedicine.TxtLabSAmt2), frmINVMedicine.TxtLabSAmt2, 0)
        !CLMLabAmt3 = IIf(IsNumeric(frmINVMedicine.TxtLabSAmt3), frmINVMedicine.TxtLabSAmt3, 0)
        !CLMLabAmt4 = IIf(IsNumeric(frmINVMedicine.TxtLabSAmt4), frmINVMedicine.TxtLabSAmt4, 0)
        !CLMLabAmt5 = IIf(IsNumeric(frmINVMedicine.TxtLabSAmt5), frmINVMedicine.TxtLabSAmt5, 0)
    End With
    ClmActOth.Update
    
    ClmActOth.Close
    Set ClmActOth = Nothing
End Sub
