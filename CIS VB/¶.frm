VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInvoiceAdj 
   Caption         =   "Invoice Adjustment"
   ClientHeight    =   8520
   ClientLeft      =   0
   ClientTop       =   240
   ClientWidth     =   12780
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   12780
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   120
      TabIndex        =   88
      Top             =   2640
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   4
      TabHeight       =   520
      TabCaption(0)   =   "General Information"
      TabPicture(0)   =   "¶.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label113"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FrameMedINjDiag"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FrameInv"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtRepReason"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TxtRemark"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtInternalRemark"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtPriDiag"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Charges"
      TabPicture(1)   =   "¶.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameDet"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Mgmt Note"
      TabPicture(2)   =   "¶.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtCASRemark"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Claim Documents"
      TabPicture(3)   =   "¶.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label2(6)"
      Tab(3).Control(1)=   "Label2(8)"
      Tab(3).Control(2)=   "Label2(9)"
      Tab(3).Control(3)=   "List1"
      Tab(3).Control(4)=   "List2"
      Tab(3).Control(5)=   "List3"
      Tab(3).ControlCount=   6
      Begin VB.TextBox txtPriDiag 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1550
         MultiLine       =   -1  'True
         TabIndex        =   44
         Top             =   3390
         Width           =   3135
      End
      Begin VB.TextBox txtInternalRemark 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   8520
         MultiLine       =   -1  'True
         TabIndex        =   45
         Top             =   960
         Width           =   3735
      End
      Begin VB.ListBox List3 
         Height          =   840
         ItemData        =   "¶.frx":0070
         Left            =   -67680
         List            =   "¶.frx":0072
         TabIndex        =   301
         ToolTipText     =   "right click to randomize contents"
         Top             =   1080
         Width           =   3615
      End
      Begin VB.ListBox List2 
         Height          =   840
         ItemData        =   "¶.frx":0074
         Left            =   -70920
         List            =   "¶.frx":0076
         TabIndex        =   299
         ToolTipText     =   "right click to randomize contents"
         Top             =   1080
         Width           =   3015
      End
      Begin VB.ListBox List1 
         Height          =   840
         ItemData        =   "¶.frx":0078
         Left            =   -74880
         List            =   "¶.frx":007A
         TabIndex        =   297
         ToolTipText     =   "right click to randomize contents"
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox TxtRemark 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   240
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   276
         Top             =   3900
         Width           =   4365
      End
      Begin VB.TextBox txtRepReason 
         Height          =   315
         Left            =   4800
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   275
         Top             =   4500
         Width           =   2820
      End
      Begin VB.TextBox txtCASRemark 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4185
         Left            =   -74880
         MaxLength       =   95
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   274
         Top             =   730
         Width           =   12135
      End
      Begin VB.Frame FrameDet 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4545
         Left            =   -74880
         TabIndex        =   158
         Top             =   680
         Width           =   12255
         Begin VB.TextBox txtBTBGRemark 
            Height          =   315
            Left            =   8400
            TabIndex        =   327
            Top             =   4200
            Width           =   3735
         End
         Begin VB.TextBox A32_GST 
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
            Height          =   315
            Left            =   7560
            MaxLength       =   8
            TabIndex        =   56
            Top             =   3360
            Width           =   735
         End
         Begin VB.TextBox A33_GST 
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
            Height          =   315
            Left            =   7560
            MaxLength       =   8
            TabIndex        =   66
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox A24_GST 
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
            Height          =   315
            Left            =   7560
            MaxLength       =   8
            TabIndex        =   314
            Top             =   1785
            Width           =   735
         End
         Begin VB.TextBox A25_GST 
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
            Height          =   315
            Left            =   7560
            MaxLength       =   8
            TabIndex        =   315
            Top             =   2040
            Width           =   735
         End
         Begin VB.TextBox A26_GST 
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
            Height          =   315
            Left            =   7560
            MaxLength       =   8
            TabIndex        =   316
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox A27_GST 
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
            Height          =   315
            Left            =   7560
            MaxLength       =   8
            TabIndex        =   317
            Top             =   2550
            Width           =   735
         End
         Begin VB.TextBox A19_GST 
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
            Height          =   315
            Left            =   7560
            MaxLength       =   8
            TabIndex        =   319
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox A20_GST 
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
            Height          =   315
            Left            =   7560
            MaxLength       =   8
            TabIndex        =   320
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox A21_GST 
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
            Height          =   315
            Left            =   7560
            MaxLength       =   8
            TabIndex        =   321
            Top             =   855
            Width           =   735
         End
         Begin VB.TextBox A22_GST 
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
            Height          =   315
            Left            =   7560
            MaxLength       =   8
            TabIndex        =   322
            Top             =   1095
            Width           =   735
         End
         Begin VB.TextBox A23_GST 
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
            Height          =   315
            Left            =   7560
            MaxLength       =   8
            TabIndex        =   323
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox A28_GST 
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
            Height          =   315
            Left            =   7560
            MaxLength       =   8
            TabIndex        =   318
            Top             =   2800
            Width           =   735
         End
         Begin VB.TextBox A7_GST 
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
            Height          =   315
            Left            =   4525
            MaxLength       =   8
            TabIndex        =   313
            Top             =   3030
            Width           =   735
         End
         Begin VB.TextBox A8_GST 
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
            Height          =   315
            Left            =   4525
            MaxLength       =   8
            TabIndex        =   312
            Top             =   3270
            Width           =   735
         End
         Begin VB.TextBox A9_GST 
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
            Height          =   315
            Left            =   4525
            MaxLength       =   8
            TabIndex        =   311
            Top             =   3510
            Width           =   735
         End
         Begin VB.TextBox A30_GST 
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
            Height          =   315
            Left            =   4525
            MaxLength       =   8
            TabIndex        =   310
            Top             =   3765
            Width           =   735
         End
         Begin VB.TextBox A10_GST 
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
            Height          =   315
            Left            =   4525
            MaxLength       =   8
            TabIndex        =   309
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox A11_GST 
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
            Height          =   315
            Left            =   4525
            MaxLength       =   8
            TabIndex        =   308
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox A12_GST 
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
            Height          =   315
            Left            =   4525
            MaxLength       =   8
            TabIndex        =   307
            Top             =   855
            Width           =   735
         End
         Begin VB.TextBox A13_GST 
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
            Height          =   315
            Left            =   4525
            MaxLength       =   8
            TabIndex        =   306
            Top             =   1095
            Width           =   735
         End
         Begin VB.TextBox A14_GST 
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
            Height          =   315
            Left            =   4525
            MaxLength       =   8
            TabIndex        =   305
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox A15_GST 
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
            Height          =   315
            Left            =   4525
            MaxLength       =   8
            TabIndex        =   304
            Top             =   1605
            Width           =   735
         End
         Begin VB.TextBox A16_GST 
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
            Height          =   315
            Left            =   4525
            MaxLength       =   8
            TabIndex        =   303
            Top             =   1890
            Width           =   735
         End
         Begin VB.TextBox A17_GST 
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
            Height          =   315
            Left            =   4525
            MaxLength       =   8
            TabIndex        =   302
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox A18_GST 
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
            Height          =   315
            Left            =   4525
            MaxLength       =   8
            TabIndex        =   67
            Top             =   2400
            Width           =   735
         End
         Begin VB.TextBox B5_GST 
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
            Height          =   315
            Left            =   1600
            MaxLength       =   8
            TabIndex        =   298
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox B8_GST 
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
            Height          =   315
            Left            =   1600
            MaxLength       =   8
            TabIndex        =   68
            Top             =   2400
            Width           =   735
         End
         Begin VB.TextBox B7_GST 
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
            Height          =   315
            Left            =   1600
            MaxLength       =   8
            TabIndex        =   296
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox B6_GST 
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
            Height          =   315
            Left            =   1600
            MaxLength       =   8
            TabIndex        =   295
            Top             =   2880
            Width           =   735
         End
         Begin VB.TextBox C2_GST 
            Alignment       =   1  'Right Justify
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
            Left            =   1600
            MaxLength       =   8
            TabIndex        =   294
            Top             =   3120
            Width           =   735
         End
         Begin VB.TextBox TExcess 
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
            Height          =   315
            Left            =   10425
            MaxLength       =   8
            TabIndex        =   261
            Top             =   2400
            Width           =   735
         End
         Begin VB.Frame Frame1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1980
            Left            =   9120
            TabIndex        =   175
            Top             =   120
            Width           =   3015
            Begin VB.TextBox A29 
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
               Height          =   315
               Left            =   885
               MaxLength       =   8
               TabIndex        =   271
               Top             =   720
               Width           =   855
            End
            Begin VB.Label lblCLNBillAmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   885
               TabIndex        =   273
               Top             =   1005
               Width           =   855
            End
            Begin VB.Label lblDispCLNBillAmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   885
               TabIndex        =   272
               Top             =   420
               Width           =   855
            End
            Begin VB.Label Label101 
               Caption         =   "Actual"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   14
               Left            =   2040
               TabIndex        =   187
               Top             =   120
               Width           =   615
            End
            Begin VB.Label Label101 
               Caption         =   "Clinic"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   13
               Left            =   960
               TabIndex        =   186
               Top             =   120
               Width           =   615
            End
            Begin VB.Label lblActCoPay 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   1920
               TabIndex        =   185
               Top             =   720
               Width           =   975
            End
            Begin VB.Line Line1 
               Index           =   0
               X1              =   1920
               X2              =   2880
               Y1              =   1560
               Y2              =   1560
            End
            Begin VB.Label lblPay2CLN 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   1920
               TabIndex        =   184
               Top             =   1605
               Width           =   975
            End
            Begin VB.Label Label76 
               Caption         =   "Amt Pay"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   285
               Index           =   3
               Left            =   75
               TabIndex        =   183
               Top             =   1620
               Width           =   1095
            End
            Begin VB.Label lblFFS 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   1920
               TabIndex        =   182
               Top             =   1275
               Width           =   975
            End
            Begin VB.Label Label76 
               Caption         =   "Less:FFS"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   210
               Index           =   2
               Left            =   75
               TabIndex        =   181
               Top             =   1260
               Width           =   855
            End
            Begin VB.Label Label76 
               Caption         =   "Bill Amt"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   210
               Index           =   1
               Left            =   75
               TabIndex        =   180
               Top             =   975
               Width           =   735
            End
            Begin VB.Label lblBillAmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   1920
               TabIndex        =   179
               Top             =   1005
               Width           =   975
            End
            Begin VB.Label Label101 
               Caption         =   "Co-Pymt"
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
               Index           =   0
               Left            =   75
               TabIndex        =   178
               Top             =   690
               Width           =   735
            End
            Begin VB.Label lblDispBillAmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   1920
               TabIndex        =   177
               Top             =   420
               Width           =   975
            End
            Begin VB.Label Label76 
               Caption         =   "Total Bill "
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   210
               Index           =   0
               Left            =   75
               TabIndex        =   176
               Top             =   420
               Width           =   735
            End
         End
         Begin VB.TextBox txtPOthDesc 
            Height          =   285
            Left            =   2400
            MaxLength       =   50
            TabIndex        =   159
            Top             =   2430
            Width           =   1380
         End
         Begin VB.TextBox txtLOthDesc 
            Height          =   285
            Left            =   5355
            MaxLength       =   50
            TabIndex        =   148
            Top             =   1300
            Width           =   1455
         End
         Begin VB.TextBox A32 
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
            Height          =   315
            Left            =   6800
            MaxLength       =   8
            TabIndex        =   135
            Top             =   3360
            Width           =   735
         End
         Begin VB.TextBox txtExcess 
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
            Height          =   315
            Left            =   10425
            MaxLength       =   8
            TabIndex        =   133
            Top             =   3840
            Width           =   855
         End
         Begin VB.TextBox A31 
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
            Left            =   9960
            MaxLength       =   12
            TabIndex        =   190
            ToolTipText     =   "54"
            Top             =   480
            Width           =   570
         End
         Begin VB.TextBox txtMBMAdjAnnLmt 
            Alignment       =   1  'Right Justify
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
            Left            =   9240
            MaxLength       =   12
            TabIndex        =   189
            ToolTipText     =   "55"
            Top             =   720
            Width           =   570
         End
         Begin VB.TextBox txtIOthDesc 
            Height          =   285
            Left            =   5355
            MaxLength       =   50
            TabIndex        =   191
            Top             =   2800
            Width           =   1455
         End
         Begin VB.TextBox A33 
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
            Height          =   315
            Left            =   6800
            MaxLength       =   8
            TabIndex        =   188
            Top             =   3600
            Width           =   735
         End
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
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   153
            Top             =   360
            Width           =   735
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
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   168
            Top             =   645
            Width           =   735
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
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   169
            Top             =   900
            Width           =   735
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
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   170
            Top             =   1150
            Width           =   735
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
            Height          =   315
            Left            =   960
            MaxLength       =   8
            TabIndex        =   174
            Top             =   1800
            Visible         =   0   'False
            Width           =   375
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
            Height          =   315
            Left            =   960
            MaxLength       =   8
            TabIndex        =   173
            Top             =   1800
            Visible         =   0   'False
            Width           =   495
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
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   154
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox B8 
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
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   161
            Top             =   2400
            Width           =   735
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
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   162
            Top             =   2640
            Width           =   735
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
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   163
            Top             =   2880
            Width           =   735
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
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   167
            Top             =   3600
            Visible         =   0   'False
            Width           =   735
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
            Height          =   315
            Left            =   480
            MaxLength       =   8
            TabIndex        =   166
            Top             =   3600
            Visible         =   0   'False
            Width           =   495
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
            Height          =   315
            Left            =   960
            MaxLength       =   8
            TabIndex        =   165
            Top             =   3600
            Visible         =   0   'False
            Width           =   495
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
            Height          =   255
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   172
            Top             =   1800
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox C2 
            Alignment       =   1  'Right Justify
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
            Left            =   840
            MaxLength       =   8
            TabIndex        =   164
            Top             =   3120
            Width           =   735
         End
         Begin VB.TextBox C1 
            Alignment       =   1  'Right Justify
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
            Left            =   840
            MaxLength       =   8
            TabIndex        =   171
            Top             =   1440
            Width           =   735
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
            Height          =   315
            Left            =   3765
            MaxLength       =   8
            TabIndex        =   136
            Top             =   360
            Width           =   735
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
            Height          =   315
            Left            =   3765
            MaxLength       =   8
            TabIndex        =   149
            Top             =   600
            Width           =   735
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
            Height          =   315
            Left            =   3765
            MaxLength       =   8
            TabIndex        =   150
            Top             =   855
            Width           =   735
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
            Height          =   315
            Left            =   3765
            MaxLength       =   8
            TabIndex        =   151
            Top             =   1095
            Width           =   735
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
            Height          =   315
            Left            =   3765
            MaxLength       =   8
            TabIndex        =   152
            Top             =   1320
            Width           =   735
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
            Height          =   315
            Left            =   3765
            MaxLength       =   8
            TabIndex        =   155
            Top             =   1605
            Width           =   735
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
            Height          =   315
            Left            =   3765
            MaxLength       =   8
            TabIndex        =   156
            Top             =   1890
            Width           =   735
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
            Height          =   315
            Left            =   3765
            MaxLength       =   8
            TabIndex        =   157
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox A18 
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
            Height          =   315
            Left            =   3765
            MaxLength       =   8
            TabIndex        =   69
            Top             =   2400
            Width           =   735
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
            Height          =   315
            Left            =   3765
            MaxLength       =   8
            TabIndex        =   197
            Top             =   3030
            Width           =   735
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
            Height          =   315
            Left            =   3765
            MaxLength       =   8
            TabIndex        =   146
            Top             =   3270
            Width           =   735
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
            Height          =   315
            Left            =   3765
            MaxLength       =   8
            TabIndex        =   147
            Top             =   3510
            Width           =   735
         End
         Begin VB.TextBox A30 
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
            Height          =   315
            Left            =   3765
            MaxLength       =   8
            TabIndex        =   160
            Top             =   3765
            Width           =   735
         End
         Begin VB.TextBox A19 
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
            Height          =   315
            Left            =   6800
            MaxLength       =   8
            TabIndex        =   137
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox A20 
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
            Height          =   315
            Left            =   6800
            MaxLength       =   8
            TabIndex        =   138
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox A21 
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
            Height          =   315
            Left            =   6800
            MaxLength       =   8
            TabIndex        =   139
            Top             =   855
            Width           =   735
         End
         Begin VB.TextBox A22 
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
            Height          =   315
            Left            =   6800
            MaxLength       =   8
            TabIndex        =   140
            Top             =   1095
            Width           =   735
         End
         Begin VB.TextBox A23 
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
            Height          =   315
            Left            =   6800
            MaxLength       =   8
            TabIndex        =   141
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox A24 
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
            Height          =   315
            Left            =   6800
            MaxLength       =   8
            TabIndex        =   142
            Top             =   1785
            Width           =   735
         End
         Begin VB.TextBox A25 
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
            Height          =   315
            Left            =   6800
            MaxLength       =   8
            TabIndex        =   143
            Top             =   2040
            Width           =   735
         End
         Begin VB.TextBox A26 
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
            Height          =   315
            Left            =   6800
            MaxLength       =   8
            TabIndex        =   144
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox A27 
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
            Height          =   315
            Left            =   6800
            MaxLength       =   8
            TabIndex        =   145
            Top             =   2550
            Width           =   735
         End
         Begin VB.TextBox A28 
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
            Height          =   315
            Left            =   6800
            MaxLength       =   8
            TabIndex        =   192
            Top             =   2800
            Width           =   735
         End
         Begin VB.TextBox txtRBAmt 
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
            Height          =   315
            Left            =   10425
            MaxLength       =   8
            TabIndex        =   134
            Top             =   2760
            Width           =   855
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
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
            Left            =   10425
            MaxLength       =   8
            TabIndex        =   193
            Top             =   3000
            Width           =   855
         End
         Begin VB.TextBox txtRBAmtExc 
            Alignment       =   1  'Right Justify
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
            Left            =   10425
            MaxLength       =   8
            TabIndex        =   194
            Top             =   3240
            Width           =   855
         End
         Begin VB.TextBox txtRBCoPay 
            Alignment       =   1  'Right Justify
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
            Left            =   10425
            MaxLength       =   8
            TabIndex        =   195
            Top             =   3525
            Width           =   855
         End
         Begin VB.TextBox txtLmtExceeded 
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
            Height          =   315
            Left            =   10425
            MaxLength       =   8
            TabIndex        =   196
            Top             =   2115
            Width           =   735
         End
         Begin VB.ComboBox cmbPharmacy 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6800
            TabIndex        =   254
            Top             =   3870
            Width           =   855
         End
         Begin VB.TextBox txtBTBG 
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
            Height          =   315
            Left            =   6800
            MaxLength       =   8
            TabIndex        =   325
            Top             =   4150
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "BTBG Remark"
            Height          =   375
            Left            =   7800
            TabIndex        =   328
            Top             =   4125
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "BTBG"
            Height          =   255
            Left            =   5350
            TabIndex        =   326
            Top             =   4200
            Width           =   1215
         End
         Begin VB.Label Label73 
            Caption         =   "GST"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   7680
            TabIndex        =   324
            ToolTipText     =   "GST Amount"
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label73 
            Caption         =   "GST"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   4680
            TabIndex        =   300
            ToolTipText     =   "GST Amount"
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label73 
            Caption         =   "GST"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   1800
            TabIndex        =   70
            ToolTipText     =   "GST Amount"
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label Label85 
            Caption         =   "Excess"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   9120
            TabIndex        =   262
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label lblPharmacy 
            Caption         =   "Pharmacy"
            Height          =   255
            Left            =   5350
            TabIndex        =   255
            Top             =   3870
            Width           =   735
         End
         Begin VB.Label Label73 
            Caption         =   "GP"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   242
            ToolTipText     =   "General Practitioner"
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label114 
            Caption         =   "Others"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   241
            Top             =   2760
            Width           =   735
         End
         Begin VB.Label Label74 
            Caption         =   "Fee + LT Med."
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
            Index           =   0
            Left            =   120
            TabIndex        =   240
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label32 
            Caption         =   "DTP&&Polio"
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
            Left            =   5350
            TabIndex        =   239
            Top             =   1785
            Width           =   1095
         End
         Begin VB.Label Label65 
            Caption         =   "Pap Smear"
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
            Left            =   5350
            TabIndex        =   238
            Top             =   1035
            Width           =   855
         End
         Begin VB.Label Label82 
            Caption         =   "Urine"
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
            Left            =   5350
            TabIndex        =   237
            Top             =   825
            Width           =   735
         End
         Begin VB.Label Label83 
            Caption         =   "B. Glucose"
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
            Left            =   5350
            TabIndex        =   236
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label84 
            Caption         =   "Laboratory"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5350
            TabIndex        =   235
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label88 
            Caption         =   "B. Drawing"
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
            Left            =   5350
            TabIndex        =   234
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label85 
            Caption         =   "MMR"
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
            Index           =   0
            Left            =   5350
            TabIndex        =   233
            Top             =   2550
            Width           =   1095
         End
         Begin VB.Label Label86 
            Caption         =   "Hepatitis B"
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
            Left            =   5350
            TabIndex        =   232
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label Label89 
            Caption         =   "Immunization"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5350
            TabIndex        =   231
            Top             =   1570
            Width           =   1455
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
            Left            =   2400
            TabIndex        =   230
            Top             =   3000
            Width           =   735
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
            Left            =   2400
            TabIndex        =   229
            Top             =   3255
            Width           =   735
         End
         Begin VB.Label Label30 
            Caption         =   "PRV Chk."
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
            Left            =   2400
            TabIndex        =   228
            Top             =   3555
            Width           =   735
         End
         Begin VB.Label Label79 
            Caption         =   "Procedures"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   227
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label43 
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
            Left            =   2400
            TabIndex        =   226
            Top             =   3840
            Width           =   735
         End
         Begin VB.Label Label73 
            Caption         =   "Specialist"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   225
            ToolTipText     =   "General Practitioner"
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label73 
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   224
            ToolTipText     =   "General Practitioner"
            Top             =   1480
            Width           =   495
         End
         Begin VB.Label Label73 
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   223
            ToolTipText     =   "General Practitioner"
            Top             =   3200
            Width           =   495
         End
         Begin VB.Label Label74 
            Caption         =   "Med."
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
            Index           =   1
            Left            =   120
            TabIndex        =   222
            Top             =   650
            Width           =   495
         End
         Begin VB.Label Label74 
            Caption         =   "Inj."
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
            Index           =   2
            Left            =   120
            TabIndex        =   221
            Top             =   900
            Width           =   495
         End
         Begin VB.Label Label74 
            Caption         =   "LT Med."
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
            Index           =   3
            Left            =   120
            TabIndex        =   220
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label74 
            Caption         =   "Fee + LT Med."
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
            Index           =   4
            Left            =   120
            TabIndex        =   219
            Top             =   2160
            Width           =   255
         End
         Begin VB.Label Label74 
            Caption         =   "Med."
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
            Index           =   5
            Left            =   120
            TabIndex        =   218
            Top             =   2450
            Width           =   495
         End
         Begin VB.Label Label74 
            Caption         =   "Inj."
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
            Index           =   6
            Left            =   120
            TabIndex        =   217
            Top             =   2700
            Width           =   495
         End
         Begin VB.Label Label74 
            Caption         =   "LT Med."
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
            Index           =   7
            Left            =   120
            TabIndex        =   216
            Top             =   2920
            Width           =   735
         End
         Begin VB.Label Label74 
            Caption         =   "X-Ray"
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
            Index           =   8
            Left            =   2400
            TabIndex        =   215
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label74 
            Caption         =   "Utr Sound"
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
            Index           =   9
            Left            =   2400
            TabIndex        =   214
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label74 
            Caption         =   "ECG"
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
            Index           =   10
            Left            =   2400
            TabIndex        =   213
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label74 
            Caption         =   "Dressing"
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
            Index           =   11
            Left            =   2400
            TabIndex        =   212
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label74 
            Caption         =   "Nubulizer"
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
            Index           =   12
            Left            =   2400
            TabIndex        =   211
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label74 
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
            Index           =   13
            Left            =   2400
            TabIndex        =   210
            Top             =   1605
            Width           =   495
         End
         Begin VB.Label Label74 
            Caption         =   "T&&S/Sto"
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
            Index           =   14
            Left            =   2400
            TabIndex        =   209
            Top             =   1875
            Width           =   735
         End
         Begin VB.Label Label74 
            Caption         =   "M. Surg."
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
            Index           =   15
            Left            =   2400
            TabIndex        =   208
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label Label87 
            Caption         =   "DTP&&Polio&&HIB"
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
            Left            =   5350
            TabIndex        =   207
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label90 
            Caption         =   "Typhoid Jab"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   5350
            TabIndex        =   206
            Top             =   3360
            Width           =   915
         End
         Begin VB.Label txtOther 
            Caption         =   "Others"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   5350
            TabIndex        =   205
            Top             =   3600
            Width           =   1095
         End
         Begin VB.Label Label73 
            Caption         =   "No Limit Deduction"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   5350
            TabIndex        =   204
            ToolTipText     =   "Geberal Pratitioner"
            Top             =   3075
            Width           =   1875
         End
         Begin VB.Label Label85 
            Caption         =   "Limit Exceeded"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   9120
            TabIndex        =   203
            Top             =   2145
            Width           =   1335
         End
         Begin VB.Label Label85 
            Caption         =   "Other Excess"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   9120
            TabIndex        =   202
            Top             =   3270
            Width           =   1575
         End
         Begin VB.Label Label85 
            Caption         =   "Room && Board"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   9120
            TabIndex        =   201
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label85 
            Caption         =   "R&&B Co-Pay"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   9120
            TabIndex        =   200
            Top             =   3555
            Width           =   1575
         End
         Begin VB.Label Label85 
            Caption         =   "R&&B Excess"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   9120
            TabIndex        =   199
            Top             =   3015
            Width           =   1575
         End
         Begin VB.Label Label85 
            Caption         =   "Total Excess"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   9120
            TabIndex        =   198
            Top             =   3840
            Width           =   1095
         End
      End
      Begin VB.Frame FrameInv 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   120
         TabIndex        =   89
         Top             =   660
         Width           =   8295
         Begin VB.TextBox txtBillNo 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   840
            MaxLength       =   20
            TabIndex        =   95
            Top             =   170
            Width           =   1305
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
            Height          =   300
            Left            =   4680
            MaxLength       =   3
            TabIndex        =   94
            Top             =   170
            Width           =   1095
         End
         Begin VB.TextBox txtCLNBillAmt 
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
            Height          =   315
            Left            =   840
            MaxLength       =   20
            TabIndex        =   92
            Top             =   405
            Width           =   1305
         End
         Begin VB.ComboBox cboMemoClaim 
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
            TabIndex        =   91
            Top             =   170
            Width           =   1155
         End
         Begin VB.ComboBox cboRefLetter 
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
            TabIndex        =   90
            Top             =   465
            Width           =   1155
         End
         Begin MSMask.MaskEdBox txtVisitDate 
            Height          =   300
            Left            =   3060
            TabIndex        =   96
            Top             =   165
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
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
            Height          =   300
            Left            =   3060
            TabIndex        =   97
            Top             =   405
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
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
         Begin VB.ComboBox cboFFS 
            Height          =   315
            Left            =   4680
            TabIndex        =   93
            Top             =   440
            Width           =   1125
         End
         Begin VB.Label Label15 
            Caption         =   "MC"
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
            Top             =   200
            Width           =   375
         End
         Begin VB.Label Label66 
            Caption         =   "Vst Time"
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
            Left            =   2280
            TabIndex        =   104
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "Vst Date"
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
            Left            =   2280
            TabIndex        =   103
            Top             =   200
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "Bill No"
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
            TabIndex        =   102
            Top             =   200
            Width           =   975
         End
         Begin VB.Label Label101 
            Caption         =   "FFS"
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
            Index           =   9
            Left            =   4320
            TabIndex        =   101
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label101 
            Caption         =   "Bill Amt"
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
            Index           =   15
            Left            =   120
            TabIndex        =   100
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label70 
            Caption         =   "Accepted"
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
            Index           =   1
            Left            =   6000
            TabIndex        =   99
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label70 
            Caption         =   "Referal Letter"
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
            Index           =   7
            Left            =   6000
            TabIndex        =   98
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.Frame FrameMedINjDiag 
         Height          =   3555
         Left            =   120
         TabIndex        =   106
         Top             =   1450
         Width           =   12255
         Begin VB.TextBox txtRepAmt 
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
            Left            =   10320
            MaxLength       =   8
            TabIndex        =   283
            Top             =   3000
            Width           =   690
         End
         Begin VB.TextBox txtRepAction 
            Height          =   285
            Left            =   7560
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   281
            Top             =   3000
            Width           =   2685
         End
         Begin VB.Frame FrameMed 
            Caption         =   "Medication && Injection"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2535
            Left            =   4680
            TabIndex        =   129
            Top             =   120
            Width           =   7335
            Begin VB.TextBox txtMedType1 
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
               Index           =   0
               Left            =   240
               MaxLength       =   50
               TabIndex        =   132
               Top             =   405
               Width           =   4335
            End
            Begin VB.TextBox txtMedDay1 
               Alignment       =   2  'Center
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
               Index           =   0
               Left            =   4560
               MaxLength       =   3
               TabIndex        =   130
               Top             =   405
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
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
               Index           =   0
               Left            =   5400
               MaxLength       =   3
               TabIndex        =   131
               Top             =   405
               Width           =   855
            End
            Begin VB.TextBox txtMedDay1 
               Alignment       =   2  'Center
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
               Index           =   1
               Left            =   4560
               MaxLength       =   3
               TabIndex        =   71
               Top             =   690
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
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
               Index           =   1
               Left            =   5400
               MaxLength       =   3
               TabIndex        =   72
               Top             =   690
               Width           =   855
            End
            Begin VB.TextBox txtMedDay1 
               Alignment       =   2  'Center
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
               Index           =   2
               Left            =   4560
               MaxLength       =   3
               TabIndex        =   73
               Top             =   945
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
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
               Index           =   2
               Left            =   5400
               MaxLength       =   3
               TabIndex        =   85
               Top             =   945
               Width           =   855
            End
            Begin VB.TextBox txtMedDay1 
               Alignment       =   2  'Center
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
               Index           =   3
               Left            =   4560
               MaxLength       =   3
               TabIndex        =   87
               Top             =   1200
               Width           =   855
            End
            Begin VB.TextBox txtMedDay1 
               Alignment       =   2  'Center
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
               Index           =   4
               Left            =   4560
               MaxLength       =   3
               TabIndex        =   243
               Top             =   1485
               Width           =   855
            End
            Begin VB.TextBox txtMedDay1 
               Alignment       =   2  'Center
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
               Index           =   5
               Left            =   4560
               MaxLength       =   3
               TabIndex        =   245
               Top             =   1740
               Width           =   855
            End
            Begin VB.TextBox txtMedDay1 
               Alignment       =   2  'Center
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
               Index           =   6
               Left            =   4560
               MaxLength       =   3
               TabIndex        =   247
               Top             =   2030
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
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
               Index           =   3
               Left            =   5400
               MaxLength       =   3
               TabIndex        =   57
               Top             =   1230
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
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
               Index           =   4
               Left            =   5400
               MaxLength       =   3
               TabIndex        =   244
               Top             =   1485
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
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
               Index           =   5
               Left            =   5400
               MaxLength       =   3
               TabIndex        =   246
               Top             =   1740
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
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
               Index           =   6
               Left            =   5400
               MaxLength       =   3
               TabIndex        =   248
               Top             =   2030
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
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
               Index           =   7
               Left            =   6240
               TabIndex        =   60
               Top             =   405
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
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
               Index           =   8
               Left            =   6240
               TabIndex        =   59
               Top             =   690
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
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
               Index           =   9
               Left            =   6240
               TabIndex        =   58
               Top             =   945
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
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
               Index           =   10
               Left            =   6240
               TabIndex        =   52
               Top             =   1230
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
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
               Index           =   11
               Left            =   6240
               TabIndex        =   51
               Top             =   1485
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
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
               Index           =   12
               Left            =   6240
               TabIndex        =   50
               Top             =   1740
               Width           =   855
            End
            Begin VB.TextBox txtMedQty1 
               Alignment       =   2  'Center
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
               Index           =   13
               Left            =   6240
               TabIndex        =   46
               Top             =   2030
               Width           =   855
            End
            Begin VB.TextBox txtMedType1 
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
               Index           =   1
               Left            =   240
               TabIndex        =   249
               Top             =   690
               Width           =   4335
            End
            Begin VB.TextBox txtMedType1 
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
               Index           =   2
               Left            =   240
               TabIndex        =   250
               Top             =   960
               Width           =   4335
            End
            Begin VB.TextBox txtMedType1 
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
               Index           =   3
               Left            =   240
               TabIndex        =   251
               Top             =   1200
               Width           =   4335
            End
            Begin VB.TextBox txtMedType1 
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
               Index           =   4
               Left            =   240
               TabIndex        =   252
               Top             =   1490
               Width           =   4335
            End
            Begin VB.TextBox txtMedType1 
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
               Index           =   5
               Left            =   240
               TabIndex        =   253
               Top             =   1750
               Width           =   4335
            End
            Begin VB.TextBox txtMedType1 
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
               Index           =   6
               Left            =   240
               TabIndex        =   256
               Top             =   2040
               Width           =   4335
            End
            Begin VB.Label Label74 
               Caption         =   "Price"
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
               Index           =   17
               Left            =   6480
               TabIndex        =   61
               Top             =   150
               Width           =   495
            End
            Begin VB.Label Label26 
               Caption         =   "2"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   257
               Top             =   720
               Width           =   135
            End
            Begin VB.Label Label26 
               Caption         =   "1"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   258
               Top             =   480
               Width           =   135
            End
            Begin VB.Label Label26 
               Caption         =   "4"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   259
               Top             =   1280
               Width           =   975
            End
            Begin VB.Label Label26 
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
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   260
               Top             =   1000
               Width           =   855
            End
            Begin VB.Label Label26 
               Caption         =   "5"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   263
               Top             =   1515
               Width           =   975
            End
            Begin VB.Label Label71 
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
               Left            =   4800
               TabIndex        =   264
               Top             =   150
               Width           =   375
            End
            Begin VB.Label Label72 
               Caption         =   "Qty"
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
               TabIndex        =   265
               Top             =   150
               Width           =   375
            End
            Begin VB.Label Labe26 
               Caption         =   "6"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   266
               Top             =   1800
               Width           =   255
            End
            Begin VB.Label Label26 
               Caption         =   "7"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   6
               Left            =   120
               TabIndex        =   267
               Top             =   2080
               Width           =   255
            End
         End
         Begin VB.Frame FrameDiag 
            Caption         =   "Diagnosis"
            Height          =   1815
            Left            =   120
            TabIndex        =   107
            Top             =   120
            Width           =   4455
            Begin VB.CommandButton cmdSearchICD2 
               Caption         =   ">"
               Height          =   255
               Left            =   4125
               TabIndex        =   112
               Top             =   585
               Width           =   255
            End
            Begin VB.CommandButton cmdSearchICD3 
               Caption         =   ">"
               Height          =   255
               Left            =   4125
               TabIndex        =   111
               Top             =   870
               Width           =   255
            End
            Begin VB.CommandButton cmdSearchICD1 
               Caption         =   ">"
               Height          =   255
               Left            =   4125
               TabIndex        =   110
               Top             =   300
               Width           =   255
            End
            Begin VB.CommandButton cmdSearchICD4 
               Caption         =   ">"
               Height          =   255
               Left            =   4125
               TabIndex        =   109
               Top             =   1155
               Width           =   255
            End
            Begin VB.CommandButton cmdSearchICD5 
               Caption         =   ">"
               Height          =   255
               Left            =   4125
               TabIndex        =   108
               Top             =   1440
               Width           =   255
            End
            Begin VB.TextBox TxtDiagDesc1 
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
               Left            =   885
               MaxLength       =   10
               TabIndex        =   118
               Top             =   285
               Width           =   3195
            End
            Begin VB.TextBox TxtDiagDesc2 
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
               Left            =   885
               MaxLength       =   10
               TabIndex        =   119
               Top             =   555
               Width           =   3195
            End
            Begin VB.TextBox TxtDiagDesc3 
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
               Left            =   885
               MaxLength       =   10
               TabIndex        =   120
               Top             =   840
               Width           =   3195
            End
            Begin VB.TextBox TxtDiagDesc4 
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
               Left            =   885
               MaxLength       =   10
               TabIndex        =   121
               Top             =   1125
               Width           =   3195
            End
            Begin VB.TextBox TxtDiagnosisOth 
               Enabled         =   0   'False
               Height          =   285
               Left            =   885
               MaxLength       =   100
               TabIndex        =   123
               Top             =   1395
               Width           =   3195
            End
            Begin VB.TextBox TxtDiagDesc5 
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
               Left            =   885
               MaxLength       =   10
               TabIndex        =   122
               Top             =   1395
               Width           =   2115
            End
            Begin VB.TextBox txtDiagnosis1 
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
               Left            =   240
               MaxLength       =   10
               TabIndex        =   113
               Top             =   285
               Width           =   615
            End
            Begin VB.TextBox txtDiagnosis2 
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
               Left            =   240
               MaxLength       =   10
               TabIndex        =   114
               Top             =   555
               Width           =   615
            End
            Begin VB.TextBox txtDiagnosis3 
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
               Left            =   240
               MaxLength       =   10
               TabIndex        =   115
               Top             =   840
               Width           =   615
            End
            Begin VB.TextBox txtDiagnosis4 
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
               Left            =   240
               MaxLength       =   10
               TabIndex        =   116
               Top             =   1125
               Width           =   615
            End
            Begin VB.TextBox TxtDiagnosis5 
               Height          =   285
               Left            =   240
               MaxLength       =   100
               TabIndex        =   117
               Top             =   1395
               Width           =   615
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
               TabIndex        =   128
               Top             =   360
               Width           =   135
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
               TabIndex        =   127
               Top             =   600
               Width           =   255
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
               TabIndex        =   126
               Top             =   900
               Width           =   255
            End
            Begin VB.Label Label6 
               Caption         =   "4"
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
               Top             =   1185
               Width           =   255
            End
            Begin VB.Label Label28 
               Caption         =   "5"
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
               TabIndex        =   124
               Top             =   1455
               Width           =   255
            End
         End
         Begin VB.Label lblPriDiag 
            Caption         =   "Preliminary Diag"
            Height          =   255
            Left            =   100
            TabIndex        =   62
            Top             =   1930
            Width           =   1155
         End
         Begin VB.Label Label75 
            Caption         =   "Remark"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   63
            Top             =   2190
            Width           =   1935
         End
         Begin VB.Label Label75 
            Caption         =   "Rep Amt"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   10320
            TabIndex        =   284
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label Label75 
            Caption         =   "Action"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   7560
            TabIndex        =   282
            Top             =   2760
            Width           =   1935
         End
         Begin VB.Label Label18 
            Caption         =   "Repudiation Reason"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4680
            TabIndex        =   279
            Top             =   2760
            Width           =   1935
         End
      End
      Begin VB.Label Label3 
         Caption         =   "Internal Remarks (LG)"
         Height          =   255
         Left            =   8520
         TabIndex        =   64
         Top             =   720
         Width           =   1755
      End
      Begin VB.Label Label2 
         Caption         =   "REIMBURSEMENT Documents"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   -67680
         TabIndex        =   268
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "SCAN Documents"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   -70920
         TabIndex        =   269
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "GL Documents"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   -74880
         TabIndex        =   270
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Repudiation Reason"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   278
         Top             =   4020
         Width           =   1935
      End
      Begin VB.Label Label113 
         Caption         =   "Remark"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   277
         Top             =   4020
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Height          =   550
      Left            =   0
      TabIndex        =   26
      Top             =   240
      Width           =   5985
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
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
         Left            =   4980
         TabIndex        =   40
         Top             =   200
         Width           =   660
      End
      Begin VB.CommandButton CmdNext 
         Caption         =   "&Next"
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
         Left            =   4440
         TabIndex        =   39
         Top             =   200
         Width           =   540
      End
      Begin VB.CommandButton CmdPrev 
         Caption         =   "Pre&v"
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
         Left            =   3900
         TabIndex        =   38
         Top             =   200
         Width           =   540
      End
      Begin VB.CommandButton cmdGo 
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
         Height          =   285
         Left            =   3360
         TabIndex        =   1
         Top             =   200
         Width           =   540
      End
      Begin VB.TextBox txtCaseNo 
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
         Left            =   840
         TabIndex        =   0
         Top             =   200
         Width           =   2415
      End
      Begin VB.Label Label42 
         Caption         =   "Case No."
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
         TabIndex        =   31
         Top             =   255
         Width           =   735
      End
   End
   Begin VB.Frame FrameCAS 
      Height          =   550
      Left            =   6120
      TabIndex        =   27
      Top             =   240
      Width           =   6615
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
         TabIndex        =   37
         Text            =   "--Select--"
         Top             =   160
         Width           =   975
      End
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
         ItemData        =   "¶.frx":007C
         Left            =   2760
         List            =   "¶.frx":007E
         TabIndex        =   35
         Text            =   "--Select--"
         Top             =   160
         Width           =   1455
      End
      Begin VB.ComboBox cboStatus 
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
         Left            =   600
         TabIndex        =   36
         Text            =   "--Select--"
         Top             =   160
         Width           =   1095
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
         Left            =   4320
         TabIndex        =   30
         Top             =   240
         Width           =   660
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
         Left            =   1800
         TabIndex        =   29
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label107 
         Caption         =   "Status"
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
         TabIndex        =   28
         Top             =   240
         Width           =   585
      End
   End
   Begin MSComctlLib.ListView lstErr 
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   8280
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
         Text            =   "Patient Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Policy No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Patient IC"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Eff Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Exp. Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Bal. Limit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Visit Lmt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Err Msg"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame FrameDis 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   9720
      TabIndex        =   43
      Top             =   960
      Width           =   3045
      Begin VB.TextBox txtAnnLmt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   47
         Top             =   165
         Width           =   900
      End
      Begin VB.TextBox txtAccumAmt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   48
         Top             =   420
         Width           =   900
      End
      Begin VB.TextBox TxtBalLimit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   49
         Top             =   675
         Width           =   900
      End
      Begin VB.TextBox txtSpecBal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   287
         Top             =   935
         Width           =   900
      End
      Begin VB.TextBox txtSpecAccLmt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   289
         Top             =   1180
         Width           =   900
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFF00&
         Caption         =   "Spec Ann. Lmt"
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
         Left            =   120
         TabIndex        =   290
         Top             =   1220
         Width           =   1215
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFF00&
         Caption         =   "Spec Bal. Lmt"
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
         Left            =   120
         TabIndex        =   288
         Top             =   970
         Width           =   1215
      End
      Begin VB.Label Label40 
         BackColor       =   &H00FFFF00&
         Caption         =   "Bal. Lmt"
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
         Left            =   120
         TabIndex        =   55
         Top             =   715
         Width           =   615
      End
      Begin VB.Label Label34 
         BackColor       =   &H00FFFF00&
         Caption         =   "Ann. Lmt"
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
         Left            =   120
         TabIndex        =   54
         Top             =   210
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF00&
         Caption         =   "Acc. Amt"
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
         Left            =   120
         TabIndex        =   53
         Top             =   460
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   5040
      TabIndex        =   76
      Top             =   960
      Width           =   4515
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
         Height          =   315
         Left            =   840
         TabIndex        =   77
         Top             =   200
         Width           =   3585
      End
      Begin VB.TextBox txtMBMName 
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
         Left            =   840
         TabIndex        =   79
         Top             =   480
         Width           =   3585
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
         Height          =   315
         Left            =   840
         TabIndex        =   78
         Top             =   770
         Width           =   3585
      End
      Begin VB.TextBox txtDept 
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
         Left            =   840
         TabIndex        =   80
         Top             =   1050
         Width           =   3585
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
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
         Index           =   0
         Left            =   120
         TabIndex        =   84
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label54 
         Caption         =   "MBM No"
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
         TabIndex        =   83
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblPatientIC 
         Caption         =   "IC No."
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
         TabIndex        =   82
         Top             =   800
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Dept"
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
         TabIndex        =   81
         Top             =   1080
         Width           =   615
      End
   End
   Begin VB.Frame Frame7 
      Height          =   555
      Left            =   120
      TabIndex        =   22
      Top             =   7920
      Width           =   12600
      Begin VB.TextBox txtpaymentstatus 
         Height          =   285
         Left            =   11880
         TabIndex        =   65
         Top             =   120
         Width           =   600
      End
      Begin VB.TextBox LblBatchNo 
         Height          =   285
         Left            =   2160
         TabIndex        =   285
         Top             =   195
         Width           =   1080
      End
      Begin VB.TextBox Label22 
         Height          =   285
         Left            =   3960
         TabIndex        =   286
         Top             =   195
         Width           =   960
      End
      Begin VB.TextBox LblBordx 
         Height          =   285
         Left            =   720
         TabIndex        =   292
         Top             =   195
         Width           =   600
      End
      Begin VB.CommandButton cmdCLMBordxApp 
         Caption         =   "Bordx App"
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
         Left            =   5040
         TabIndex        =   280
         Top             =   160
         Width           =   975
      End
      Begin VB.CommandButton cmdExclusions 
         Caption         =   "&Exclusions"
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
         Left            =   6000
         TabIndex        =   86
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
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
         Left            =   7680
         TabIndex        =   14
         Top             =   160
         Width           =   615
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
         Left            =   8280
         TabIndex        =   15
         Top             =   160
         Width           =   975
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
         Left            =   9840
         TabIndex        =   16
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
         Left            =   10680
         TabIndex        =   17
         Top             =   160
         Width           =   615
      End
      Begin VB.CommandButton cmdCompute 
         Caption         =   "&Compute"
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
         Left            =   6840
         TabIndex        =   13
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "E&dit"
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
         Left            =   9240
         TabIndex        =   11
         Top             =   160
         Width           =   615
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
         Left            =   9240
         TabIndex        =   12
         Top             =   160
         Width           =   615
      End
      Begin VB.Label Label21 
         Caption         =   "Posting"
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
         Index           =   1
         Left            =   11280
         TabIndex        =   293
         Top             =   160
         Width           =   615
      End
      Begin VB.Label Label21 
         Caption         =   "Payment"
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
         Index           =   0
         Left            =   3240
         TabIndex        =   291
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Label33 
         Caption         =   "Batch No."
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
         Left            =   1320
         TabIndex        =   34
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Label108 
         Caption         =   "Bordx ?"
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
         TabIndex        =   25
         Top             =   195
         Width           =   615
      End
   End
   Begin VB.Frame FrameMBM 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   18
      Top             =   960
      Width           =   4890
      Begin VB.TextBox txtClinicCode 
         Height          =   330
         Left            =   795
         TabIndex        =   2
         Top             =   180
         Width           =   580
      End
      Begin VB.TextBox txtClinicName 
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
         Height          =   330
         Left            =   1360
         TabIndex        =   3
         Top             =   180
         Width           =   3210
      End
      Begin VB.TextBox txtNPClinic 
         Height          =   330
         Left            =   1360
         TabIndex        =   4
         Top             =   180
         Visible         =   0   'False
         Width           =   2610
      End
      Begin VB.TextBox txtInvNo 
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
         Left            =   795
         MaxLength       =   30
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtInvDate 
         Height          =   315
         Left            =   795
         TabIndex        =   7
         Top             =   765
         Width           =   1575
         _ExtentX        =   2778
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
      Begin MSMask.MaskEdBox txtInvRecDate 
         Height          =   315
         Left            =   3240
         TabIndex        =   8
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.ComboBox cboProcMth 
         Height          =   315
         Left            =   3240
         TabIndex        =   10
         Top             =   765
         Width           =   1335
      End
      Begin VB.ComboBox CboPayto 
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
         Left            =   795
         TabIndex        =   5
         Text            =   "C - Clinic"
         Top             =   1050
         Width           =   1575
      End
      Begin VB.ComboBox CboInvRecMode 
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
         Left            =   3240
         TabIndex        =   9
         Top             =   1050
         Width           =   1335
      End
      Begin VB.Label Label101 
         Caption         =   "Mth"
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
         Index           =   5
         Left            =   2520
         TabIndex        =   75
         ToolTipText     =   "Processing Month"
         Top             =   795
         Width           =   735
      End
      Begin VB.Label lblClnType 
         Caption         =   "lblClnType"
         Height          =   255
         Left            =   3240
         TabIndex        =   74
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label110 
         Caption         =   "RecMode"
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
         TabIndex        =   42
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label78 
         Caption         =   "RecDate"
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
         TabIndex        =   41
         Top             =   525
         Width           =   735
      End
      Begin VB.Label Label109 
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
         Left            =   120
         TabIndex        =   33
         Top             =   1095
         Width           =   615
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
         TabIndex        =   21
         Top             =   260
         Width           =   735
      End
      Begin VB.Label Label67 
         Caption         =   "InvDate"
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
         TabIndex        =   20
         Top             =   795
         Width           =   735
      End
      Begin VB.Label Label35 
         Caption         =   "InvNo"
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
         TabIndex        =   19
         Top             =   555
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView lstInvoice 
      Height          =   945
      Left            =   120
      TabIndex        =   32
      Top             =   8160
      Visible         =   0   'False
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   1667
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
      NumItems        =   489
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
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Bordxed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Bill No."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Stat. Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Visit Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Visit Time"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "MC"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Clinic"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "Bill Amt"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "PreparedBy"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "101"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "102"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "103"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   18
         Text            =   "104"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   19
         Text            =   "105"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   20
         Text            =   "106"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   21
         Text            =   "Dental"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   22
         Text            =   "Maternity"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   23
         Text            =   "PRV checkup"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   24
         Text            =   "X-Ray"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   25
         Text            =   "Ultra Sound"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   26
         Text            =   "ECG"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   27
         Text            =   "Dressing"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   28
         Text            =   "Nebulizer"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   29
         Text            =   "I & D"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   30
         Text            =   "T&S/STO"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   31
         Text            =   "Minor Surgery"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   32
         Text            =   "Proc Others"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   33
         Text            =   "Blood Drawing"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   34
         Text            =   "Blood Glucose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   35
         Text            =   "Urine"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   36
         Text            =   "Pap Smear"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   37
         Text            =   "Lab-Others"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   38
         Text            =   "DTP"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   39
         Text            =   "DTP Polio HIB"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   40
         Text            =   "Hepatitis B"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   41
         Text            =   "MMR"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   42
         Text            =   "Imm Others"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   43
         Text            =   "Co-Payment"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   44
         Text            =   "POth Desc"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   45
         Text            =   "LOth Desc"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   46
         Text            =   "IOth Desc"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   47
         Text            =   "MedInjType1"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(49) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   48
         Text            =   "MedInjType2"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(50) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   49
         Text            =   "MedInjType3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(51) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   50
         Text            =   "MedInjType4"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(52) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   51
         Text            =   "MedInjType5"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(53) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   52
         Text            =   "MedInjDay1"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(54) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   53
         Text            =   "MedInjDay2"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(55) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   54
         Text            =   "MedInjDay3"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(56) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   55
         Text            =   "MedInjDay4"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(57) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   56
         Text            =   "MedInjDay5"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(58) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   57
         Text            =   "MedInjQty1"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(59) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   58
         Text            =   "MedInjQty2"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(60) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   59
         Text            =   "MedInjQty3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(61) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   60
         Text            =   "MedInjQty4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(62) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   61
         Text            =   "MedInjQty5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(63) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   62
         Text            =   "Diagnosis1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(64) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   63
         Text            =   "Diagnosis2"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(65) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   64
         Text            =   "Diagnosis3"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(66) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   65
         Text            =   "Diagnosis4"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(67) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   66
         Text            =   "Payor"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(68) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   67
         Text            =   "CLN Plan"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(69) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   68
         Text            =   "Cln Ins"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(70) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   69
         Text            =   "HLT Code"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(71) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   70
         Text            =   "Bal Limit"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(72) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   71
         Text            =   "Visit Lmt"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(73) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   72
         Text            =   "Card Type"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(74) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   73
         Text            =   "Visit Ind"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(75) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   74
         Text            =   "Annual Ind"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(76) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   75
         Text            =   "Co-pay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(77) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   76
         Text            =   "Pol No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(78) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   77
         Text            =   "Patient IC"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(79) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   78
         Text            =   "Eff Date"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(80) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   79
         Text            =   "Exp Date"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(81) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   80
         Text            =   "Payto"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(82) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   81
         Text            =   "Payee"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(83) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   82
         Text            =   "Supp Lmt Type"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(84) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   83
         Text            =   "GP Fee"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(85) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   84
         Text            =   "GP+LT Med"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(86) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   85
         Text            =   "GP + Inj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(87) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   86
         Text            =   "GP + Med"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(88) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   87
         Text            =   "SP Fee"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(89) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   88
         Text            =   "SP+ LT Med"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(90) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   89
         Text            =   "SP + Inj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(91) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   90
         Text            =   "SP + Med"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(92) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   91
         Text            =   "total Gp(c1)"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(93) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   92
         Text            =   "total SP(C2)"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(94) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   93
         Text            =   "DiagDesc5"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(95) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   94
         Text            =   "DiagOthInd"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(96) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   95
         Text            =   "DiagDesc"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(97) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   96
         Text            =   "Vision"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(98) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   97
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(99) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   98
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(100) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   99
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(101) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   100
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(102) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   101
         Text            =   "MedType1(5)"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(103) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   102
         Text            =   "MedType1(6)"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(104) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   103
         Text            =   "MedDay1(5)"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(105) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   104
         Text            =   "MedDay1(6)"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(106) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   105
         Text            =   "MedQty1(5)"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(107) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   106
         Text            =   "MedQty1(6)"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(108) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   107
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(109) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   108
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(110) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   109
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(111) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   110
         Text            =   "InvRecDate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(112) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   111
         Text            =   "InvRecMode"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(113) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   112
         Text            =   "remark"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(114) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   113
         Text            =   "AdjAnnLmt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(115) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   114
         Text            =   "NonGPFee"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(116) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   115
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(117) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   116
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(118) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   117
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(119) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   118
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(120) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   119
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(121) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   120
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(122) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   121
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(123) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   122
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(124) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   123
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(125) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   124
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(126) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   125
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(127) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   126
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(128) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   127
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(129) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   128
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(130) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   129
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(131) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   130
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(132) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   131
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(133) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   132
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(134) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   133
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(135) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   134
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(136) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   135
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(137) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   136
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(138) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   137
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(139) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   138
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(140) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   139
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(141) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   140
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(142) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   141
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(143) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   142
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(144) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   143
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(145) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   144
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(146) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   145
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(147) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   146
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(148) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   147
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(149) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   148
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(150) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   149
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(151) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   150
         Text            =   "CLMOpenDate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(152) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   151
         Text            =   "CLMOpenBy"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(153) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   152
         Text            =   "CLMUpdateBy"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(154) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   153
         Text            =   "CLMUpdateDate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(155) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   154
         Text            =   "CLMBordxStatus"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(156) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   155
         Text            =   "CLMStatus"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(157) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   156
         Text            =   "CLMCloseBy"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(158) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   157
         Text            =   "CLMCloseDate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(159) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   158
         Text            =   "TotFCLN"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(160) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   159
         Text            =   "TotCLNExc"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(161) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   160
         Text            =   "TotFinalAct"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(162) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   161
         Text            =   "TotFinalAPP"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(163) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   162
         Text            =   "TotFinalMem"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(164) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   163
         Text            =   "CLMBalLmt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(165) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   164
         Text            =   "CLMPrevBal"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(166) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   165
         Text            =   "CLMPaidByM2M"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(167) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   166
         Text            =   "CLMPaidByM2C"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(168) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   167
         Text            =   "CLMActGPBMed"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(169) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   168
         Text            =   "CLMActGPMedInj"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(170) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   169
         Text            =   "CLMActGPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(171) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   170
         Text            =   "CLMActSPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(172) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   171
         Text            =   "CLMActSPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(173) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   172
         Text            =   "CLMActSPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(174) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   173
         Text            =   "CLMActXRay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(175) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   174
         Text            =   "CLMActUltraSound"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(176) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   175
         Text            =   "CLMActECG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(177) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   176
         Text            =   "CLMActDressing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(178) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   177
         Text            =   "CLMActNebulizer"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(179) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   178
         Text            =   "CLMActID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(180) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   179
         Text            =   "CLMActTSSTO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(181) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   180
         Text            =   "CLMActMinorSurgery"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(182) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   181
         Text            =   "CLMActPOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(183) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   182
         Text            =   "CLMActBloodDrawing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(184) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   183
         Text            =   "CLMActBloodGlucose"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(185) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   184
         Text            =   "CLMActUrine"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(186) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   185
         Text            =   "CLMActPapSmear"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(187) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   186
         Text            =   "CLMActLOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(188) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   187
         Text            =   "CLMActDTPPolio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(189) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   188
         Text            =   "CLMActDTPPolioHIB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(190) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   189
         Text            =   "CLMActHepatitisB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(191) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   190
         Text            =   "CLMActMMR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(192) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   191
         Text            =   "CLMActIOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(193) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   192
         Text            =   "CLMActDental"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(194) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   193
         Text            =   "CLMActMaternity"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(195) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   194
         Text            =   "CLMActPRVCheckup"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(196) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   195
         Text            =   "CLMActVision"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(197) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   196
         Text            =   "CLMActTyphoid"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(198) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   197
         Text            =   "CLMActOthersBalNoDed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(199) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   198
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(200) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   199
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(201) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   200
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(202) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   201
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(203) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   202
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(204) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   203
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(205) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   204
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(206) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   205
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(207) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   206
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(208) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   207
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(209) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   208
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(210) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   209
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(211) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   210
         Text            =   "CLMContGPBMed"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(212) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   211
         Text            =   "CLMContGPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(213) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   212
         Text            =   "CLMContGPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(214) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   213
         Text            =   "CLMContSPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(215) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   214
         Text            =   "CLMContSPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(216) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   215
         Text            =   "CLMContSPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(217) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   216
         Text            =   "CLMContXRay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(218) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   217
         Text            =   "CLMContUltraSound"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(219) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   218
         Text            =   "CLMContECG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(220) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   219
         Text            =   "CLMContDressing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(221) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   220
         Text            =   "CLMContNebulizer"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(222) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   221
         Text            =   "CLMContID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(223) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   222
         Text            =   "CLMContTSSTO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(224) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   223
         Text            =   "CLMContMinorSurgery"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(225) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   224
         Text            =   "CLMContPOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(226) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   225
         Text            =   "CLMContBloodDrawing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(227) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   226
         Text            =   "CLMContBloodGlucose"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(228) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   227
         Text            =   "CLMContUrine"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(229) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   228
         Text            =   "CLMContPapSmear"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(230) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   229
         Text            =   "CLMContLOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(231) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   230
         Text            =   "CLMContDTPPolio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(232) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   231
         Text            =   "CLMContDTPPolioHIB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(233) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   232
         Text            =   "CLMContHepatitisB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(234) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   233
         Text            =   "CLMContMMR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(235) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   234
         Text            =   "CLMContIOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(236) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   235
         Text            =   "CLMContDental"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(237) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   236
         Text            =   "CLMContMaternity"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(238) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   237
         Text            =   "CLMContPRVCheckup"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(239) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   238
         Text            =   "CLMContVision"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(240) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   239
         Text            =   "CLMTyphoid"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(241) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   240
         Text            =   "CLMOthersBalNoDed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(242) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   241
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(243) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   242
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(244) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   243
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(245) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   244
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(246) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   245
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(247) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   246
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(248) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   247
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(249) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   248
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(250) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   249
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(251) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   250
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(252) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   251
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(253) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   252
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(254) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   253
         Text            =   "CLMAppGPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(255) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   254
         Text            =   "CLMAppGPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(256) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   255
         Text            =   "CLMAppGPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(257) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   256
         Text            =   "CLMAppSPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(258) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   257
         Text            =   "CLMAppSPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(259) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   258
         Text            =   "CLMAppSPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(260) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   259
         Text            =   "CLMAppXRay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(261) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   260
         Text            =   "CLMAppUltraSound"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(262) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   261
         Text            =   "CLMAppECG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(263) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   262
         Text            =   "CLMAppDressing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(264) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   263
         Text            =   "CLMAppNebulizer"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(265) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   264
         Text            =   "CLMAppID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(266) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   265
         Text            =   "CLMAppTSSTO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(267) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   266
         Text            =   "CLMAppMinorSurgery"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(268) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   267
         Text            =   "CLMAppPOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(269) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   268
         Text            =   "CLMAppBloodDrawing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(270) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   269
         Text            =   "CLMAppBloodGlucose"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(271) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   270
         Text            =   "CLMAppUrine"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(272) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   271
         Text            =   "CLMAppPapSmear"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(273) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   272
         Text            =   "CLMAppLOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(274) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   273
         Text            =   "CLMAppDTPPolio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(275) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   274
         Text            =   "CLMAppDTPPolioHIB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(276) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   275
         Text            =   "CLMAppHepatitisB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(277) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   276
         Text            =   "CLMAppMMR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(278) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   277
         Text            =   "CLMAppIOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(279) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   278
         Text            =   "CLMAppDental"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(280) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   279
         Text            =   "CLMAppMaternity"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(281) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   280
         Text            =   "CLMAppPRVCheckup"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(282) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   281
         Text            =   "CLMAppVision"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(283) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   282
         Text            =   "CLMAppTyphoidBalNoDed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(284) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   283
         Text            =   "CLMAppOthersNoBalDed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(285) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   284
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(286) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   285
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(287) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   286
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(288) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   287
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(289) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   288
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(290) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   289
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(291) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   290
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(292) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   291
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(293) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   292
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(294) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   293
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(295) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   294
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(296) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   295
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(297) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   296
         Text            =   "CLMBnfGPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(298) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   297
         Text            =   "CLMBnfGPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(299) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   298
         Text            =   "CLMBnfGPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(300) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   299
         Text            =   "CLMBnfSPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(301) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   300
         Text            =   "CLMBnfSPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(302) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   301
         Text            =   "CLMBnfSPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(303) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   302
         Text            =   "CLMBnfXray"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(304) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   303
         Text            =   "CLMBnfUltraSound"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(305) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   304
         Text            =   "CLMBnfECG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(306) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   305
         Text            =   "CLMBnfDressing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(307) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   306
         Text            =   "CLMBnfNebulizer"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(308) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   307
         Text            =   "CLMBnfID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(309) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   308
         Text            =   "CLMBnfTSSTO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(310) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   309
         Text            =   "CLMBnfMinorSurgery"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(311) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   310
         Text            =   "CLMBnfPOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(312) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   311
         Text            =   "CLMBnfBloodDrawing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(313) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   312
         Text            =   "CLMBnfBloodGlucose"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(314) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   313
         Text            =   "CLMBnfUrine"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(315) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   314
         Text            =   "CLMBnfPapSmear"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(316) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   315
         Text            =   "CLMBnfLOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(317) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   316
         Text            =   "CLMBnfDTPPolio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(318) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   317
         Text            =   "CLMBnfDTPPolioHIB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(319) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   318
         Text            =   "CLMBnfHepatitisB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(320) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   319
         Text            =   "CLMBnfMMR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(321) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   320
         Text            =   "CLMBnfIOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(322) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   321
         Text            =   "CLMBnfDental"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(323) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   322
         Text            =   "CLMBnfMaternity"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(324) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   323
         Text            =   "CLMBnfPRVCheckup"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(325) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   324
         Text            =   "CLMBnfVision"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(326) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   325
         Text            =   "CLMBNFTyPhoidBalNoDed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(327) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   326
         Text            =   "CLMBNFOthersBalNoDed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(328) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   327
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(329) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   328
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(330) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   329
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(331) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   330
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(332) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   331
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(333) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   332
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(334) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   333
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(335) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   334
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(336) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   335
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(337) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   336
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(338) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   337
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(339) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   338
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(340) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   339
         Text            =   "CLMMBMExcGPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(341) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   340
         Text            =   "CLMMBMExcGPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(342) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   341
         Text            =   "CLMMBMExcGPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(343) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   342
         Text            =   "CLMMBMExcSPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(344) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   343
         Text            =   "CLMMBMExcSPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(345) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   344
         Text            =   "CLMMBMExcSPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(346) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   345
         Text            =   "CLMMBMExcXRay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(347) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   346
         Text            =   "CLMMBMExcUltraSound"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(348) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   347
         Text            =   "CLMMBMExcECG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(349) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   348
         Text            =   "CLMMBMExcDressing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(350) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   349
         Text            =   "CLMMBMExcNebulizer"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(351) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   350
         Text            =   "CLMMBMExcID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(352) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   351
         Text            =   "CLMMBMExcTSSTO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(353) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   352
         Text            =   "CLMMBMExcMinorSurgery"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(354) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   353
         Text            =   "CLMMBMExcPOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(355) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   354
         Text            =   "CLMMBMExcBloodDrawing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(356) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   355
         Text            =   "CLMMBMExcBloodGlucose"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(357) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   356
         Text            =   "CLMMBMExcUrine"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(358) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   357
         Text            =   "CLMMBMExcPapSmear"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(359) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   358
         Text            =   "CLMMBMExcLOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(360) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   359
         Text            =   "CLMMBMExcDTPPolio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(361) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   360
         Text            =   "CLMMBMExcDTPPolioHIB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(362) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   361
         Text            =   "CLMMBMExcHepatitisB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(363) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   362
         Text            =   "CLMMBMExcMMR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(364) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   363
         Text            =   "CLMMBMExcIOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(365) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   364
         Text            =   "CLMMBMExcDental"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(366) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   365
         Text            =   "CLMMBMExcMaternity"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(367) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   366
         Text            =   "CLMMBMExcPRVCheckup"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(368) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   367
         Text            =   "CLMMBMExcVision"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(369) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   368
         Text            =   "CLMMBMExcTyphoid"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(370) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   369
         Text            =   "CLMMBMExcOthersNoBalDed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(371) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   370
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(372) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   371
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(373) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   372
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(374) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   373
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(375) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   374
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(376) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   375
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(377) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   376
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(378) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   377
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(379) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   378
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(380) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   379
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(381) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   380
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(382) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   381
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(383) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   382
         Text            =   "CLMNPayCLNGPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(384) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   383
         Text            =   "CLMNPayCLNGPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(385) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   384
         Text            =   "CLMNPayCLNGPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(386) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   385
         Text            =   "CLMNPayCLNSPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(387) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   386
         Text            =   "CLMNPayCLNSPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(388) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   387
         Text            =   "CLMNPayCLNSPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(389) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   388
         Text            =   "CLMNPayCLNXRay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(390) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   389
         Text            =   "CLMNPayCLNUltraSound"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(391) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   390
         Text            =   "CLMNPayCLNECG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(392) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   391
         Text            =   "CLMNPayCLNDressing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(393) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   392
         Text            =   "CLMNPayCLNNebulizer"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(394) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   393
         Text            =   "CLMNPayCLNID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(395) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   394
         Text            =   "CLMNPayCLNTSSTO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(396) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   395
         Text            =   "CLMNPayCLNMinorSurgery"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(397) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   396
         Text            =   "CLMNPayCLNPThers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(398) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   397
         Text            =   "CLMNPayCLNBloodDrawing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(399) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   398
         Text            =   "CLMNPayCLNBloodGlucose"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(400) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   399
         Text            =   "CLMNPayCLNUrine"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(401) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   400
         Text            =   "CLMNPayCLNPapSmear"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(402) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   401
         Text            =   "CLMNPayCLNLOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(403) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   402
         Text            =   "CLMNPayCLNDTPolio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(404) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   403
         Text            =   "CLMNPayCLNDTPPolioHIB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(405) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   404
         Text            =   "CLMNPayCLNHepatitisB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(406) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   405
         Text            =   "CLMNPayCLNMMR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(407) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   406
         Text            =   "CLMNPayCLNIOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(408) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   407
         Text            =   "CLMNPayCLNDental"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(409) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   408
         Text            =   "CLMNPayCLNMaternity"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(410) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   409
         Text            =   "CLMNPayCLNPRVCheckup"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(411) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   410
         Text            =   "CLMNPayCLNVision"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(412) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   411
         Text            =   "CLMNPayCLNTyphoid"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(413) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   412
         Text            =   "CLMNPayCLNOthersNoBalDed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(414) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   413
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(415) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   414
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(416) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   415
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(417) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   416
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(418) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   417
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(419) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   418
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(420) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   419
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(421) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   420
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(422) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   421
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(423) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   422
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(424) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   423
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(425) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   424
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(426) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   425
         Text            =   "CLMPayCLNGPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(427) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   426
         Text            =   "CLMPayCLNGPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(428) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   427
         Text            =   "CLMPayCLNGPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(429) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   428
         Text            =   "CLMPayCLNSPBMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(430) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   429
         Text            =   "CLMPayCLNSPMedInj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(431) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   430
         Text            =   "CLMPayCLNSPLTMed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(432) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   431
         Text            =   "CLMPayCLNXRay"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(433) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   432
         Text            =   "CLMPayCLNUltraSound"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(434) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   433
         Text            =   "CLMPayCLNECG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(435) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   434
         Text            =   "CLMPayCLNDressing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(436) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   435
         Text            =   "CLMPayCLNNebulizer"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(437) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   436
         Text            =   "CLMPayCLNID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(438) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   437
         Text            =   "CLMPayCLNTSSTO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(439) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   438
         Text            =   "CLMPayCLNMinorSurgery"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(440) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   439
         Text            =   "CLMPayCLNPOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(441) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   440
         Text            =   "CLMPayCLNBloodDrawing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(442) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   441
         Text            =   "CLMPayCLNBloodGlucose"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(443) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   442
         Text            =   "CLMPayCLNUrine"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(444) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   443
         Text            =   "CLMPayCLNPapSmear"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(445) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   444
         Text            =   "CLMPayCLNLOthers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(446) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   445
         Text            =   "CLMPayCLNDTPPolio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(447) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   446
         Text            =   "CLMPayCLNDTPPolioHIB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(448) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   447
         Text            =   "CLMPayCLNHepatitisB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(449) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   448
         Text            =   "CLMPayCLNMMR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(450) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   449
         Text            =   "CLMPayCLNIOhters"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(451) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   450
         Text            =   "CLMPayCLNDental"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(452) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   451
         Text            =   "CLMPayCLNMaternity"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(453) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   452
         Text            =   "CLMPayCLNPRVCheckup"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(454) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   453
         Text            =   "CLMPayCLNVision"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(455) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   454
         Text            =   "CLMPayCLNTyphoid"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(456) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   455
         Text            =   "CLMPayCLNOthersNoBalDed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(457) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   456
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(458) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   457
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(459) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   458
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(460) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   459
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(461) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   460
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(462) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   461
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(463) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   462
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(464) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   463
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(465) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   464
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(466) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   465
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(467) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   466
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(468) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   467
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(469) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   468
         Text            =   "MCOProcFee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(470) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   469
         Text            =   "MCOProcFeeStatus"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(471) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   470
         Text            =   "InvProcPeriod"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(472) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   471
         Text            =   "CopayCLNColl"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(473) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   472
         Text            =   "BillAmt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(474) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   473
         Text            =   "TyphoidNoBalDed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(475) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   474
         Text            =   "OthersNoBalDed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(476) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   475
         Text            =   "BordxFees"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(477) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   476
         Text            =   "Accpeted"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(478) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   477
         Text            =   "RepReason"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(479) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   478
         Text            =   "Ref Letter"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(480) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   479
         Text            =   "Mgmt Note"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(481) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   480
         Text            =   "Limit Exceeded"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(482) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   481
         Text            =   "MEDCHARGES1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(483) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   482
         Text            =   "MEDCHARGES2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(484) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   483
         Text            =   "MEDCHARGES3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(485) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   484
         Text            =   "MEDCHARGES4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(486) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   485
         Text            =   "MEDCHARGES5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(487) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   486
         Text            =   "MEDCHARGES6"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(488) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   487
         Text            =   "MEDCHARGES7"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(489) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   488
         Text            =   "TEST"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblInvoice 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Invoice Adjustment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "frmInvoiceAdj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bDataStatus As Boolean
Dim bExit As Boolean
'Dim IntIndex As String
Dim EditFlag As Boolean
Dim ScanCode As String
Dim AA, CC As Integer
Dim strPayTo, strPayee As String
Dim strMBMno As String
Dim IntInvAmt As Double
Dim CASENumber, NewCaseNo As String
Dim MembershipNo As String
Dim CovPersonID As String
Dim ASOInd As String
Dim UtilisedAmt As Double
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
Dim strHLTCode As String
Dim strVisitIND, strAnnualInd As String
Dim strCoPay As String
Dim strSpecStatus As String
Dim strMEPStatus As String
Dim StrDentLmtStatus As String
Dim StrSpecLmtStatus As String
Dim strBNFStatus As String
Dim strDentExtLmt As String
Dim str1stSpecLmt As String
Dim strMCStatus As String
'Dim StrInvIndex As String
Dim strGRPMCOStatus As String
Dim strPLNGrossStatus As String
Dim strMBMDepartment As String
Dim strMBMMCDaysBal As String
Dim strGRPPostageFees As String
Dim strAccumAmt As String
Dim MBM As New ADODB.Recordset
Dim lstErrRec As ListItem
Dim VPlan, VExpDate, VEffDate, VDiagCode, VDiagIND As String
Dim SuppLmtType As String
Dim Z1, Z2, Z3, Z4, Z5, Z6, Z7, Z8 As Double
Dim CISCAS1 As New ADODB.Recordset
Dim ReasonOtherIND As String
Dim StrDentStatus As String
Dim tmpFinalApp As Double
Dim tmpPrevAdjAmt As Double
Dim tmpPrevNonGPFee As Double
Public lstRec As ListItem
Dim tmpCASStatus As String
Dim tmpClaimability As String
Dim tmpPending As String
Dim tmpLastUpdateDate As String
Dim tmpLastUpdateBy As String
Dim tmpClinicCode As String
Dim tmpNPClinic As String
Dim tmpInvoiceNo As String
Dim tmpInvoiceDate As String
Dim tmpDateVisit As String
Dim tmpMCDays As String
Dim tmpTimeVisit As String
Dim tmpMedInjType1 As String
Dim tmpMedInjType2 As String
Dim tmpMedInjType3 As String
Dim tmpMedInjType4 As String
Dim tmpMedInjType5 As String
Dim tmpMedInjDays1 As String
Dim tmpMedInjDays2 As String
Dim tmpMedInjDays3 As String
Dim tmpMedInjDays4 As String
Dim tmpMedInjDays5 As String
Dim tmpMedInjQty1 As String
Dim tmpMedInjQty2 As String
Dim tmpMedInjQty3 As String
Dim tmpMedInjQty4 As String
Dim tmpMedInjQty5 As String
Dim tmpMedInjType6 As String
Dim tmpMedInjType7 As String
Dim tmpMedInjDays6 As String
Dim tmpMedInjDays7 As String
Dim tmpMedInjQty6 As String
Dim tmpMedInjQty7 As String
Dim tmpInvoiceReceiptDate As String
Dim tmpInvoiceReceiptMode As String
Dim tmpCasRemark As String
Dim tmpCASTreatReason1 As String
Dim tmpCASTreatReason2 As String
Dim tmpCASTreatReason3 As String
Dim tmpCASTreatReason4 As String
Dim tmpCASTreatReason5 As String
Dim tmpCASDiagnosis5 As String
Dim tmpMCStatus As String
Dim tmpNoOfDiag As String
Dim tmpNoOfMedication As String
Dim tmpCASnumber As String
Dim tmpINVLastUpdateBy As String
Dim tmpINVLastUpdateDate As String
Dim tmpBillNo As String
Dim tmpBillDate As String
Dim tmpINVAmount As Double
Dim tmpINVPayTo As String
Dim tmpINVPayee As String
Dim tmpINVBillAmt As Double
Dim tmpCLMGPFeeBasicMed As Double
Dim tmpCLMGPFeeLTMed As Double
Dim tmpCLMGPFeeMedInj As Double
Dim tmpCLMSPFeeBasicMed As Double
Dim tmpCLMSPFeeLTMed As Double
Dim tmpCLMSPFeeMedInj As Double
Dim tmpCLMDental As Double
Dim tmpCLMMaternity As Double
Dim tmpCLMPRVCheckup As Double
Dim tmpCLMXRay As Double
Dim tmpCLMUltraSound As Double
Dim tmpCLMECG As Double
Dim tmpCLMDressing As Double
Dim tmpCLMNebulizer As Double
Dim tmpCLMID As Double
Dim tmpCLMTSSTO As Double
Dim tmpCLMMinorSurgery As Double
Dim tmpCLMPOthers As Double
Dim tmpCLMBloodDrawing As Double
Dim tmpCLMBloodGlucose As Double
Dim tmpCLMUrine As Double
Dim tmpCLMPapSmear As Double
Dim tmpCLMLOthers As Double
Dim tmpCLMDTPPolio As Double
Dim tmpCLMDTPPolioHIB As Double
Dim tmpCLMHepatitisB As Double
Dim tmpCLMMMR As Double
Dim tmpCLMIOthers As Double
Dim tmpCLMCoPayment As Double
Dim tmpCLMPOthersDesc As String
Dim tmpCLMLOthersDesc As String
Dim tmpCLMIOthersDesc As String
Dim tmpCLMVision As Double
Dim tmpCLMTyphiad As Double
Dim tmpCLMOtherNoBalDed As Double
Dim tmpGPFee As Double
Dim tmpGPLTMed As Double
Dim tmpSPFee As Double
Dim tmpSPLTMed As Double
Dim tmpSPInj As Double
Dim tmpSPMed As Double
Dim tmpGPInj As Double
Dim tmpGPMed As Double
Dim tmpClaimCloseDate As String
Dim tmpClaimClosedBy As String
Dim tmpClaimLastUpdateDate As String
Dim tmpClaimLastUpdateUser As String
Dim tmpCLMPayableByMedix2C As Double
Dim tmpCLMExcess As Double
Dim tmpCLMMCOProcFee As Double
Dim tmpCLMPayableByMedix2M As Double
Dim tmpCLMCLNExcess As Double
Dim tmpCLMTotalActual As Double
Dim tmpCLMFinalApprove As Double
Dim tmpCLMBalanceLimit As Double
Dim tmpCLMPreviousBalance As Double
Dim tmpCLMPaidBYMedix2M As Double
Dim tmpCLMPaidBYMedix2C As Double
Dim tmpCLMAdjAnnLmt As Double
Dim tmpClmNonGPFee As Double
Dim tmpNonGPFee As Double
Dim tmpCLMCoPayColl As Double
Dim tmpCLMMCOProcStatus As String
Dim tmpCLMProcMth As String
Dim tmpCLMProcYR As String
Dim tmpCLMCoPayCLNColl As Double
Dim tmpCLMBordxAmt As Double
Dim tmpCLMNo As String
Dim tmpGRPBordxFees As String
Dim strGRPBordxFees As String
Public tmpPayto As String
Dim strPLNMMPerc As String
Dim tmpLimitExceeded As Double
Dim bIbmUser As Boolean
Dim BSNUser As Boolean
Dim FFSAmount As Double
Dim FFSIndicator As Boolean
Dim tmpPymtStatus As Boolean
Dim tmpvisit As Double
Dim TmpSPVisitLimit As Double
Dim TmpMEVisitLimit As Double
Dim VisitInd As String
Dim Tmpprepostmatellimit As Double
Dim PDentalStatus As String
Dim tmpPDentalAnnLimit As String
Dim GPLimit As Double
Dim SPLimit As Double
Dim GPCoPay As Double
Dim SPCoPay As Double
Dim tmpTopupAmt  As Double
Dim tmpGrpCompany As String
Dim PictureFileName As String
Dim TxtOPBatchNo As String
Dim PostStatus As Boolean
Dim UpdateDia As Boolean
Dim txtEffDate As Date
Dim txtExpDate As Date
Dim strChkDoubleEntry As String
Public txtTabMBMNo, txtTabCoverID As String
Dim TxtPatientName, txtMBMStatus, TxtPayor, txtPlanCode, TxtInsType As String
Dim TxtCoPay, TxtCardType, txtPolStatus, txtLTM, TxtVisitLimit, txtDentStatus, txtInVIndex, lblGrpCompany As String
Dim strPlanCategory, strGLIssued As String
Dim bPLNBTBG As Boolean

Private Const SW_RESTORE        As Long = &H9&
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub A1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A29_Change()
Dim a As Double
Dim B As Double
Dim LimitValue As Double
Dim copay As Double
copay = 0
LimitValue = 0
Dim NolmtDid As Double
NolmtDid = 0

If IsNumeric(A32) Then
    NolmtDid = NolmtDid + CDbl(A32)
End If

If IsNumeric(A33) Then
    NolmtDid = NolmtDid + CDbl(A33)
End If


If tmpGrpCompany = "C014" Then    ' Group company for Prince Court
    If NolmtDid > 0 Then
        copay = 0
        A29 = 0
        lblActCoPay = 0
    ElseIf Trim(lblClnType) = "S" Or (Trim(txtClinicCode) = "NP") Or (Trim(txtClinicCode) = "NPS") Then
        copay = 0
        lblActCoPay = 0
        A29 = 0
    ElseIf Trim(txtClinicCode) = Trim("P481") Or (Trim(txtClinicCode) = "NP") Then
        LimitValue = 100
        copay = 5
    ElseIf Trim(txtClinicCode) = Trim("P432") Or (Trim(txtClinicCode) = "NPS") Then
        LimitValue = 250
        copay = 0
    End If
End If
If LimitValue > 0 Then

    If lblDispCLNBillAmt >= LimitValue Then
        a = IIf(IsNumeric(lblDispCLNBillAmt), lblDispCLNBillAmt, 0)
        B = lblDispCLNBillAmt - LimitValue
        lblActCoPay = B
        If Len(A29) Then
            If CDbl(A29) = CDbl(B) Then
            
            Else
                A29 = B
            End If
        Else
            A29 = B
        End If
        
        If CDbl(a) <> "0" Then
            lblCLNBillAmt = CDbl(a) - CDbl(B)
        End If
    Else
        a = IIf(IsNumeric(lblDispCLNBillAmt), lblDispCLNBillAmt, 0)
        
        If Len(A29) Then
            If CDbl(A29) = CDbl(copay) Then
            
            Else
                A29 = copay
            End If
        Else
            A29 = copay
        End If
        
        B = IIf(IsNumeric(A29), A29, 0)
        If CDbl(a) <> "0" Then
            lblCLNBillAmt = CDbl(a) - CDbl(B)
        End If
    End If
Else
    a = IIf(IsNumeric(lblDispCLNBillAmt), lblDispCLNBillAmt, 0)
    B = IIf(IsNumeric(A29), A29, 0)
    If CDbl(a) <> "0" Then
        If a < 0 And B < 0 Then
            lblCLNBillAmt = Format(Abs(CDbl(a)) - Abs(CDbl(B)), "-#,##0.00")
        Else
            lblCLNBillAmt = Format(CDbl(a) - CDbl(B), "#,##0.00")
        End If
    End If
End If

    a = IIf(IsNumeric(lblDispCLNBillAmt), lblDispCLNBillAmt, 0)
    B = IIf(IsNumeric(A29), A29, 0)
    If CDbl(a) <> "0" Then
        If a < 0 And B < 0 Then
            lblCLNBillAmt = Format(Abs(CDbl(a)) - Abs(CDbl(B)), "-#,##0.00")
        Else
            lblCLNBillAmt = Format(CDbl(a) - CDbl(B), "#,##0.00")
            lblBillAmt = Format(CDbl(a) - CDbl(B), "#,##0.00") '11-07-2016
        End If
        lblPay2CLN = Format(CDbl(a) - CDbl(B) - lblFFS, "#,##0.00") '11-07-2016
    End If
'End If
    
End Sub

Private Sub A29_LostFocus()
    
    If InStr(1, A29, "-") Then
         A29 = IIf(IsNumeric(A29), A29, 0)
         A29 = A29 * (-1)
    Else
           A29 = IIf(IsNumeric(A29), A29, 0)
    End If

    lblActCoPay = A29
End Sub

Private Sub A32_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A32_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A33_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A33_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub txtBTBG_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub txtBTBG_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub cboRefLetter_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboRefLetter.Text, cboRefLetter.SelStart) + Chr(KeyAscii) + Mid(cboRefLetter.Text, cboRefLetter.SelStart + cboRefLetter.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboRefLetter.ListCount - 1
       
        If NewText = LCase(Left(cboRefLetter.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = cboRefLetter.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            cboRefLetter.Text = ValidValue
            cboRefLetter.SelStart = 0
            cboRefLetter.SelLength = Len(ValidValue)
        End If
    End If

End Sub

Private Sub cmdCLMBordxApp_Click()
Const METHOD_NAME = "cmdCLMBordxApp_Click"
On Error GoTo debugs
Dim CAS As New ADODB.Recordset
Dim strsql As String
    strsql = "Select CASNumber, CASClaimsApp from CISCAS " & _
            " where CASNumber ='" & Trim(txtCaseNo) & "'"
    CAS.Open strsql, MediConnCISDB, adOpenForwardOnly, adLockOptimistic
    Do While Not CAS.EOF
        CAS!CASClaimsApp = "Y"
        CAS.update
        CAS.MoveNext
    Loop
    CAS.Close
    Set CAS = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdExclusions_Click()
    If Trim(lblGrpCompany) = "" Then
        MsgBox "Please Check Group Company! ", vbOKOnly
    Else
        Load frmGeneralBenefitExcEnq
        frmGeneralBenefitExcEnq.cboGRPCompany = Trim(lblGrpCompany)
        frmGeneralBenefitExcEnq.cmdSearch_Click
    End If
End Sub

Private Sub lblActCoPay_Change()
Dim a As Double
Dim B As Double

    'a = IIf(IsNumeric(lblDispBillAmt), lblDispBillAmt, 0)
    'B = IIf(IsNumeric(lblActCoPay), lblActCoPay, 0)
    'If CDbl(a) <> "0" Then
      '  'lblBillAmt = Format(CDbl(a) + (CDbl(B) * -1), "#,##0.00")
     '   lblBillAmt = Format(CDbl(a) - CDbl(B), "#,##0.00")
    'End If
     a = IIf(IsNumeric(lblDispBillAmt), lblDispBillAmt, 0)
    B = IIf(IsNumeric(lblActCoPay), lblActCoPay, 0)
    If CDbl(a) <> "0" Then
    If a < 0 And B < 0 Then
        lblBillAmt = Format(Abs(CDbl(a)) - Abs(CDbl(B)), "-#,##0.00")
    End If
    Else
        lblBillAmt = Format(CDbl(a) - CDbl(B), "#,##0.00")
    End If
End Sub

Private Sub A3_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A30_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A31_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A31_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A7_LostFocus()
    Call DisplayBillAmt
End Sub
Private Sub A8_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A9_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A10_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A11_LostFocus()
    Call DisplayBillAmt
End Sub
Private Sub A12_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A13_LostFocus()
    Call DisplayBillAmt
End Sub
Private Sub A14_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A15_LostFocus()
    Call DisplayBillAmt
End Sub
Private Sub A16_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A17_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A18_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A19_LostFocus()
    Call DisplayBillAmt
End Sub
Private Sub A20_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A21_LostFocus()
    Call DisplayBillAmt
End Sub
Private Sub A22_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A23_LostFocus()
    Call DisplayBillAmt
End Sub
Private Sub A24_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A25_LostFocus()
    Call DisplayBillAmt
End Sub
Private Sub A26_LostFocus()
    Call DisplayBillAmt
End Sub

Private Sub A27_LostFocus()
    Call DisplayBillAmt
End Sub
Private Sub A28_LostFocus()
    Call DisplayBillAmt
End Sub
Private Sub A30_LostFocus()
    Call DisplayBillAmt
End Sub
Private Sub B1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub B1_LostFocus()
    Z1 = IIf(IsNumeric(B1), B1, 0)
    Z2 = IIf(IsNumeric(B2), B2, 0)
    Z3 = IIf(IsNumeric(B3), B3, 0)
    Z4 = IIf(IsNumeric(B4), B4, 0)
    C1 = CDbl(Z1) + CDbl(Z2) + CDbl(Z3) + CDbl(Z4)
    Call DisplayBillAmt
End Sub

Private Sub B2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub B2_LostFocus()
    Z1 = IIf(IsNumeric(B1), B1, 0)
    Z2 = IIf(IsNumeric(B2), B2, 0)
    Z3 = IIf(IsNumeric(B3), B3, 0)
    Z4 = IIf(IsNumeric(B4), B4, 0)
    C1 = CDbl(Z1) + CDbl(Z2) + CDbl(Z3) + CDbl(Z4)
    Call DisplayBillAmt
End Sub

Private Sub B3_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub B3_LostFocus()
    Z1 = IIf(IsNumeric(B1), B1, 0)
    Z2 = IIf(IsNumeric(B2), B2, 0)
    Z3 = IIf(IsNumeric(B3), B3, 0)
    Z4 = IIf(IsNumeric(B4), B4, 0)
    C1 = CDbl(Z1) + CDbl(Z2) + CDbl(Z3) + CDbl(Z4)
    Call DisplayBillAmt
End Sub

Private Sub B4_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub B4_LostFocus()
    Z1 = IIf(IsNumeric(B1), B1, 0)
    Z2 = IIf(IsNumeric(B2), B2, 0)
    Z3 = IIf(IsNumeric(B3), B3, 0)
    Z4 = IIf(IsNumeric(B4), B4, 0)
    C1 = CDbl(Z1) + CDbl(Z2) + CDbl(Z3) + CDbl(Z4)
    Call DisplayBillAmt
End Sub

Private Sub B5_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub B5_LostFocus()
    Z5 = IIf(IsNumeric(B5), B5, 0)
    Z6 = IIf(IsNumeric(B6), B6, 0)
    Z7 = IIf(IsNumeric(B7), B7, 0)
    Z8 = IIf(IsNumeric(B8), B8, 0)
    C2 = CDbl(Z5) + CDbl(Z6) + CDbl(Z7) + CDbl(Z8)
    Call DisplayBillAmt
End Sub

Private Sub B6_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub B6_LostFocus()
    Z5 = IIf(IsNumeric(B5), B5, 0)
    Z6 = IIf(IsNumeric(B6), B6, 0)
    Z7 = IIf(IsNumeric(B7), B7, 0)
    Z8 = IIf(IsNumeric(B8), B8, 0)
    C2 = CDbl(Z5) + CDbl(Z6) + CDbl(Z7) + CDbl(Z8)
    Call A29_Change
    
    Call DisplayBillAmt
End Sub

Private Sub B7_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub B7_LostFocus()
    Z5 = IIf(IsNumeric(B5), B5, 0)
    Z6 = IIf(IsNumeric(B6), B6, 0)
    Z7 = IIf(IsNumeric(B7), B7, 0)
    Z8 = IIf(IsNumeric(B8), B8, 0)
    C2 = CDbl(Z5) + CDbl(Z6) + CDbl(Z7) + CDbl(Z8)
    Call A29_Change
    
    Call DisplayBillAmt
    
End Sub

Private Sub B8_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub B8_LostFocus()
    Z5 = IIf(IsNumeric(B5), B5, 0)
    Z6 = IIf(IsNumeric(B6), B6, 0)
    Z7 = IIf(IsNumeric(B7), B7, 0)
    Z8 = IIf(IsNumeric(B8), B8, 0)
    C2 = CDbl(Z5) + CDbl(Z6) + CDbl(Z7) + CDbl(Z8)
    Call DisplayBillAmt
End Sub

Private Sub B5_GST_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub B5_GST_LostFocus()
    Z5 = IIf(IsNumeric(B5_GST), B5_GST, 0)
    Z6 = IIf(IsNumeric(B6_GST), B6_GST, 0)
    Z7 = IIf(IsNumeric(B7_GST), B7_GST, 0)
    Z8 = IIf(IsNumeric(B8_GST), B8_GST, 0)
    C2_GST = CDbl(Z5) + CDbl(Z6) + CDbl(Z7) + CDbl(Z8)
    Call DisplayBillAmt
End Sub

Private Sub B6_GST_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub B6_GST_GST_LostFocus()
    Z5 = IIf(IsNumeric(B5_GST), B5_GST, 0)
    Z6 = IIf(IsNumeric(B6_GST), B6_GST, 0)
    Z7 = IIf(IsNumeric(B7_GST), B7_GST, 0)
    Z8 = IIf(IsNumeric(B8_GST), B8_GST, 0)
    C2_GST = CDbl(Z5) + CDbl(Z6) + CDbl(Z7) + CDbl(Z8)
    Call A29_Change
    
    Call DisplayBillAmt
End Sub

Private Sub B7_GST_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub B7_GST_LostFocus()
    Z5 = IIf(IsNumeric(B5_GST), B5_GST, 0)
    Z6 = IIf(IsNumeric(B6_GST), B6_GST, 0)
    Z7 = IIf(IsNumeric(B7_GST), B7_GST, 0)
    Z8 = IIf(IsNumeric(B8_GST), B8_GST, 0)
    C2_GST = CDbl(Z5) + CDbl(Z6) + CDbl(Z7) + CDbl(Z8)
    Call A29_Change
    
    Call DisplayBillAmt
    
End Sub

Private Sub B8_GST_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub B8_GST_LostFocus()
    Z5 = IIf(IsNumeric(B5_GST), B5_GST, 0)
    Z6 = IIf(IsNumeric(B6_GST), B6_GST, 0)
    Z7 = IIf(IsNumeric(B7_GST), B7_GST, 0)
    Z8 = IIf(IsNumeric(B8_GST), B8_GST, 0)
    C2_GST = CDbl(Z5) + CDbl(Z6) + CDbl(Z7) + CDbl(Z8)
    Call DisplayBillAmt
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
Private Sub A18_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A19_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A21_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A22_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A23_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A24_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A25_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A26_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A27_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A28_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A29_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub CboInvRecMode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboInvRecMode.Text, CboInvRecMode.SelStart) + Chr(KeyAscii) + Mid(CboInvRecMode.Text, CboInvRecMode.SelStart + CboInvRecMode.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboInvRecMode.ListCount - 1
       
        If NewText = LCase(Left(CboInvRecMode.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboInvRecMode.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboInvRecMode.Text = ValidValue
            CboInvRecMode.SelStart = 0
            CboInvRecMode.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CmdClearAll_Click()
    Call ClearEntries
    Call ClearMBMInfo
    txtInvDate = "__/__/____"
    txtInvNo = ""
    txtClinicName.Visible = True
    txtClinicName = ""
    txtClinicCode = ""
    txtInvRecDate = "__/__/____"
    CboInvRecMode.Text = ""
    CmdUpdate.Visible = False
    CmdEdit.Visible = True
    CmdMBMInfo.Enabled = False
    txtNPClinic.Visible = False
    Call EntriesEnabled(False)
    lstInvoice.ListItems.Clear
    lstErr.ListItems.Clear
    CmdEdit.Enabled = False
    lblCLNBillAmt = ""
    lblBillAmt = ""
    lblPay2CLN = ""
    cmbPharmacy = ""
    lblFFS = ""
    
End Sub

Private Sub cmdCompute_Click()
    frmWorksheet.InvoiceClick = 2
    frmWorksheet.Show
End Sub

Private Sub cmdEdit_Click()
    If MembershipNo <> "" Then
        If PostStatus = True And tmpPymtStatus = False And Trim(LblBordx) = False Then
            CmdEdit.Visible = False
            CmdEdit.Enabled = False
            CmdUpdate.Visible = True
            CmdUpdate.Enabled = True
            UpdateDia = True
            FrameMedINjDiag.Enabled = True
            FrameMed.Enabled = False
            txtRepReason.Enabled = False
            txtRepAction.Enabled = False
            txtRepAmt.Enabled = False
            MsgBox "This case posted into SUN System Only can amend Diagnosis and Remark ! ", vbInformation
        ElseIf tmpPymtStatus = True And Trim(LblBordx) = "True" And PostStatus = "True" Then
            CmdEdit.Visible = False
            CmdEdit.Enabled = False
            CmdUpdate.Visible = True
            CmdUpdate.Enabled = True
            UpdateDia = True
            FrameMedINjDiag.Enabled = True
            FrameMed.Enabled = False
            txtRepReason.Enabled = False
            txtRepAction.Enabled = False
            txtRepAmt.Enabled = False
            MsgBox "This case posted into SUN System Only can amend Diagnosis and Remark ! ", vbInformation
        ElseIf tmpPymtStatus = True Or Trim(LblBordx) = "True" Then
            CmdEdit.Visible = True
            CmdEdit.Enabled = False
            CmdUpdate.Visible = False
            CmdUpdate.Enabled = False
            If Trim(LblBordx) = "True" Then
            MsgBox "This case already Bordx ! ", vbInformation
            ElseIf tmpPymtStatus = True Then
            MsgBox "This case already Payment done ! ", vbInformation
            ElseIf tmpPymtStatus = True And Trim(LblBordx) = "True" Then
            MsgBox "This case already Payment and Bordx Done ! ", vbInformation
            End If
        Else
            CmdEdit.Visible = False
            CmdEdit.Enabled = False
            CmdUpdate.Visible = True
            CmdUpdate.Enabled = True
            Call EntriesEnabled(True)
            A7.Enabled = False    '601 Dental
            A30.Enabled = False   '604 Vision
            If Trim(txtDentStatus) = "C" Then ' only for Combine Lmt (Dental)
                A7.Enabled = True    '601 Dental
                A30.Enabled = True   '604 Vision
            ElseIf Trim(txtDentStatus) = "Y" Then
                A7.Enabled = True    '601 Dental
                A30.Enabled = True
            End If
            EditFlag = True
            txtClinicCode.SetFocus
        End If
       'If Trim(LblBordx) = "False" Then
         '   CmdEdit.Visible = False
          '  CmdEdit.Enabled = False
           ' CmdUpdate.Visible = True
           ' CmdUpdate.Enabled = True
           ' Call EntriesEnabled(True)
           ' A7.Enabled = False    '601 Dental
           ' A30.Enabled = False   '604 Vision
           ' If Trim(txtDentStatus) = "C" Then ' only for Combine Lmt (Dental)
            '    A7.Enabled = True    '601 Dental
            '    A30.Enabled = True   '604 Vision
            'ElseIf Trim(txtDentStatus) = "Y" Then
             '   A7.Enabled = True    '601 Dental
              '  A30.Enabled = True
            'End If
            'EditFlag = True
            'txtClinicCode.SetFocus
       'End If
    End If
      Dim StrSqlMsgProm As String
            Dim MsgProm As ADODB.Recordset
            Dim lblsplnote As String
            Set MsgProm = New ADODB.Recordset
            
            StrSqlMsgProm = "Select  PLNSpecialRemark from MBMCLNOthers inner join cisdb..PLN on PLN.PLNCode = MBMCLNOthers.plntype where MBMNumber ='" + MembershipNo + "' "
            MsgProm.Open StrSqlMsgProm, medixmasterconn, adOpenForwardOnly, adLockReadOnly
            lblsplnote = IIf(Trim(Len(MsgProm!PLNSpecialRemark)), MsgProm!PLNSpecialRemark, "")
            If Len(lblsplnote) Then
            MsgBox lblsplnote, vbOKOnly + vbInformation
            End If
            MsgProm.Close
            Set MsgProm = Nothing
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CallLimit(tmpMBMTYpe As String, tmpAnnLmt As String, tmpVisitLmt As String, tmpSpecLmt As String, tmpLimitStatus As String)
    Const METHOD_NAME = "CallLimit"
    On Error GoTo debugs
    Dim sArr() As String
    Dim MBMNo As String
    Dim CovID As String
    
    Dim ASOqry As String
    Dim ASORS As New ADODB.Recordset
    Dim ASOlimit As Double
    Dim IPLimit As Double
    Dim OPLimit As Double
    Dim IPUtilization As Double
    Dim OPUtilization As Double
    Dim IsCombined As String
    
    If ASOInd = "Y" Then
        If Len(txtMBMNo) Then
           If InStr(1, txtMBMNo, "-") Then
              sArr = Split(txtMBMNo, "-")
              MBMNo = sArr(0)
              CovID = sArr(1)
           Else
              MBMNo = txtMBMNo
              CovID = "00"
           End If
        End If
        ASOqry = ""
        ASOqry = "Select MBMASOLimit,MBMIPLimit,MBMOPLimit,MBMIPUtilized,MBMOPUtilized ,IsIPOPCombined  from Tbl_ASOMemberDetails  where MBMNumber= '" + MBMNo + "' and MBMCoverID='" + CovID + "'"
        ASORS.Open ASOqry, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        Do Until ASORS.EOF
             ASOlimit = IIf(Len(ASORS!MBMASOLimit), ASORS!MBMASOLimit, 0)
             IPLimit = IIf(Len(ASORS!MBMIPLimit), ASORS!MBMIPLimit, 0)
             OPLimit = IIf(Len(ASORS!MBMOPLimit), ASORS!MBMOPLimit, 0)
             IPUtilization = IIf(Len(ASORS!MBMIPUtilized), ASORS!MBMIPUtilized, 0)
             OPUtilization = IIf(Len(ASORS!MBMOPUtilized), ASORS!MBMOPUtilized, 0)
             IsCombined = IIf(Len(ASORS!IsIPOPCombined), ASORS!IsIPOPCombined, "N")
        ASORS.MoveNext
        Loop
        ASORS.Close
        Set ASORS = Nothing
        
        If IsCombined = "Y" Then
            tmpAnnLmt = ASOlimit
            UtilisedAmt = IPUtilization + OPUtilization
        Else
            tmpAnnLmt = OPLimit
            UtilisedAmt = OPUtilization
        End If
    Else
        tmpAnnLmt = ""
        tmpAnnLmt = Format(GetAnnualLimit("H", strClnPln, "N", txtEffDate, tmpMBMTYpe, "DC", tmpLimitStatus), "#,##0.00")
    End If
    
    tmpSpecLmt = ""
    tmpSpecLmt = Format(GetSpecLimit("H", strClnPln, "N", txtEffDate, tmpLimitStatus, "DC"), "#,##0.00")
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Public Sub CmdGo_Click()
Const METHOD_NAME = "CmdGo_Click"
On Error GoTo debugs
Dim TotalSpecLimit As Double
Dim tmpRecCnt As Integer
Dim ZZ As String
Dim tmpMBMTYpe As String
Dim PictureFileName As String
Dim isPrinDepGPShared As Boolean

    Label22 = "False"
    
    tmpPymtStatus = False
    lstInvoice.ListItems.Clear
    SSTab1.Tab = 0
    Call ClearEntries
    Call InvoiceINFO
    txtMBMNo = ChkStr(txtMBMNo)
    frmWorksheet.INVCLNForm = 3
    If Trim(txtMBMNo) <> "" Then
        ZZ = InStr(1, txtMBMNo, "-")
        MembershipNo = Trim(Left(txtMBMNo, IIf(ZZ - 1 <= 0, Len(Trim(txtMBMNo)), ZZ - 1)))
        If ZZ = 0 Then
            CovPersonID = "00"
        Else
            CovPersonID = Trim(Right(txtMBMNo, IIf(ZZ - 1 <= 0, 0, Len(Trim(txtMBMNo)) - ZZ)))
        End If
        
        Dim TopupRec As New ADODB.Recordset
        tmpTopupAmt = 0
        TopupRec.Open "SELECT Top 1 sum(TopupAmt) as sumOfTopupAmt from MBMCISTopupHistory where MBMNumber='" & MembershipNo & "' group by MBMNumber", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        With TopupRec
            If TopupRec.EOF = False Then
                tmpTopupAmt = IIf(IsNumeric(TopupRec!sumOfTopupAmt), TopupRec!sumOfTopupAmt, 0)
            End If
        End With
        TopupRec.Close
        Set TopupRec = Nothing
        
        'get the SP and GP limit  'Trinath Code
        Dim GetGPSPLmt As New ADODB.Recordset
        Dim GPSPQry As String
        GPSPQry = ""
        
        If CovPersonID = "00" Then
            GPSPQry = "Select PLN.GRPCompany,PLN.PrinGPLimit as GPLimit,PLN.PrinSPLimit as SPLimit,CISDB..PLN.ASOInd, PLN.PLNBTBG, COALESCE(IsPrinDepGPShared,0) AS IsPrinDepGPShared from HISDB..MBMCLNOthers inner join CISDB..PLN on PLN.PLNCode = MBMCLNOthers.PLNType where HISDB..MBMCLNOthers.MBMNumber ='" + MembershipNo + "'"
        Else
            GPSPQry = "Select PLN.GRPCompany,PLN.DepGPLimit as GPLimit,PLN.DepSPLimit as SPLimit,CISDB..PLN.ASOInd, PLN.PLNBTBG, COALESCE(IsPrinDepGPShared,0) AS IsPrinDepGPShared from HISDB..MBMCoveredPersonsCLN  inner join CISDB..PLN on PLN.PLNCode = MBMCoveredPersonsCLN.MBMCCLNPLNCode    where HISDB..MBMCoveredPersonsCLN.MBMCNumber ='" + MembershipNo + "' and HISDB..MBMCoveredPersonsCLN.MBMCCoverID ='" + CovPersonID + "'"
        End If
        
        GetGPSPLmt.Open GPSPQry, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        If GetGPSPLmt.EOF = False Then
'                txtGPlimit = IIf(IsNumeric(GetGPSPLmt!GPLimit), GetGPSPLmt!GPLimit, 0)
'                txtSpLimit = IIf(IsNumeric(GetGPSPLmt!SPLimit), GetGPSPLmt!SPLimit, 0)
            tmpGrpCompany = IIf(Len(GetGPSPLmt!GRPCompany), GetGPSPLmt!GRPCompany, 0)
            ASOInd = IIf(Len(GetGPSPLmt!ASOInd), GetGPSPLmt!ASOInd, "N")
            bPLNBTBG = IIf(Len(GetGPSPLmt!PLNBTBG), GetGPSPLmt!PLNBTBG, False)
            isPrinDepGPShared = GetGPSPLmt!isPrinDepGPShared
        End If
        GetGPSPLmt.Close
        Set GetGPSPLmt = Nothing
        
        Call MemberSuppInfo(MembershipNo, CovPersonID, strBalLimit, strVisitLimit, strAccumAmt, _
                            StrDentStatus, strCoPay, strPayor, strHLTCode, strClnPln, _
                            strClnIns, strEffdate, strExpDate, strPolNo, strMBMDataStatus, _
                            strPOlStatus, strPatientName, strPatientDOB, strPatientIC, _
                            strMBMCLNLTM, strCardType, strVisitIND, strAnnualInd, strSpecStatus, _
                            strMEPStatus, strDentLmt, strSpecLmt, strMEPLmt, SuppLmtType, _
                            StrDentLmtStatus, StrSpecLmtStatus, strBNFStatus, strDentExtLmt, _
                            str1stSpecLmt, strGRPMCOStatus, strPLNGrossStatus, strMBMDepartment, _
                            strMCStatus, strMBMMCDaysBal, strGRPPostageFees, strGRPBordxFees, strPLNMMPerc)
        tmpMBMTYpe = "P"
        If CovPersonID <> "00" Then
            tmpMBMTYpe = "S"
        End If
        If Trim(MembershipNo) <> "" Then
            EditFlag = False
            CmdMBMInfo.Enabled = True
            CmdEdit.Visible = True
            CmdEdit.Enabled = True
            txtTabMBMNo = MembershipNo
            txtTabCoverID = CovPersonID
            txtMBMName = strPatientName
            TxtPayor = strPayor
            txtPlanCode = strClnPln
            TxtInsType = strClnIns
            TxtPatientName = strPatientName
            txtEffDate = strEffdate
            txtExpDate = strExpDate
            txtPolStatus = strPOlStatus
            txtMBMStatus = strMBMDataStatus
            txtPatientIC = strPatientIC
            txtDept = strMBMDepartment
            PictureFileName = LocSignPath & Trim(txtPatientIC) & ".jpg"
            On Error Resume Next
            TxtPatientName = strPatientName
            If ASOInd = "Y" Then
                txtAccumAmt = Format(UtilisedAmt, "#,##0.00")
            Else
                txtAccumAmt = Format(strAccumAmt, "#,##0.00")
            End If
            'txtAnnLmt = Format(GetAnnualLimit(strClnIns, strClnPln, strHLTCode, CDate(strEffdate), tmpMBMTYpe, strPayor, SuppLmtType), "#,##0.00")
            TxtVisitLimit = Format(GetVisitLimit(strClnIns, strClnPln, strHLTCode, CDate(strEffdate), tmpMBMTYpe, strPayor), "#,##0.00")
            TxtBalLimit = txtAnnLmt
            If IsNumeric(txtAnnLmt) Then
                TxtBalLimit = Format(strBalLimit, "#,##0.00")
            Else
                TxtBalLimit = Format(0, "#,##0.00")
            End If
            If IsNumeric(TxtVisitLimit) Then
                TxtVisitLimit = Format(strVisitLimit, "#,##0.00")
            End If

            ' Write code to get visit limit and values based on plan code
            Dim visitLmt As New ADODB.Recordset
            Dim VisitLimitquery As String
            
            VisitLimitquery = ""
            VisitLimitquery = "Select AnnualLimit.VisitLimit,prepostmatellimit,PDentalStatus,PDentalAnnLimit,SPVisitLimit,MEVisitLimit,PLN.VISLmtInd from  CISDB..AnnualLimit Inner join CISDB..pln ON pln.PLNCode=AnnualLimit.PLNCode where AnnualLimit.PLNCode='" & Trim(strClnPln) & "'"
            visitLmt.Open VisitLimitquery, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            If visitLmt.EOF = False Then
            tmpvisit = IIf(IsNumeric(visitLmt!VisitLimit), visitLmt!VisitLimit, 0)
            Tmpprepostmatellimit = IIf(IsNumeric(visitLmt!prepostmatellimit), visitLmt!prepostmatellimit, 0)
            TmpSPVisitLimit = IIf(IsNumeric(visitLmt!SPVisitLimit), visitLmt!SPVisitLimit, 0)
            TmpMEVisitLimit = IIf(IsNumeric(visitLmt!MEVisitLimit), visitLmt!MEVisitLimit, 0)
            VisitInd = IIf(Len(visitLmt!VisLmtIND), visitLmt!VisLmtIND, "")
            PDentalStatus = IIf(Len(visitLmt!PDentalStatus), visitLmt!PDentalStatus, "")
            tmpPDentalAnnLimit = IIf(IsNumeric(visitLmt!PDentalAnnLimit), visitLmt!PDentalAnnLimit, 0)
        End If
        visitLmt.Close
        Set visitLmt = Nothing
        
         '------Calculate Limt Start-------
        Dim SpecLmtRec As New ADODB.Recordset
        Dim tmpLimitStatus As String
        Dim tmpAnnLmt As String
        Dim tmpVisitLmt As String
        Dim tmpSpecLmt As String
        Dim tmpSpecAnnLmt As String
        Dim strAppend As String
        Dim tmpMaternityInd As String
        Dim tmpDentalInd As String
        Dim tmpSpecLimitStatus As String
        tmpLimitStatus = SuppLmtType
        strPlanCategory = SuppLmtType
        Dim tmpSpecAccLmt As String
        
        Dim MBMNo As String
        Dim CovID As String
        Dim sArr() As String
        Dim dblMeternity As Double
        Dim dblDental As Double
        Dim OtherDed As Double
        Dim thyroid As Double
            
         Call CallLimit(tmpMBMTYpe, tmpAnnLmt, tmpVisitLmt, tmpSpecLmt, tmpLimitStatus)
         txtAnnLmt = tmpAnnLmt
        ' txtSpecLmt = tmpSpecLmt
         tmpSpecAnnLmt = tmpSpecLmt
        
         txtAccumAmt = 0
         tmpSpecLimitStatus = ""
         SpecLmtRec.Open "SELECT SuppLimitStatus FROM SpecialistLimit WHERE PLNCode='" & Trim(strClnPln) & "'", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
         If SpecLmtRec.EOF = False Then
             tmpSpecLimitStatus = IIf(Len(SpecLmtRec!SuppLimitStatus), SpecLmtRec!SuppLimitStatus, "A")
         End If
         SpecLmtRec.Close
         Set SpecLmtRec = Nothing
            
        If Trim(tmpSpecAnnLmt) = "N/A" Then
            If Len(txtMBMNo) Then
                 Erase sArr
                 If InStr(1, txtMBMNo, "-") Then
                      sArr = Split(txtMBMNo, "-")
                      MBMNo = sArr(0)
                      CovID = sArr(1)
                  Else
                      MBMNo = txtMBMNo
                      CovID = CovPersonID
                  End If
                 
                 strAppend = ""
                 If isPrinDepGPShared Then
                    strAppend = strAppend & "  CISCas.MBMnumber= '" & Trim(MBMNo) & "'"
                 Else
                    If tmpLimitStatus = "A" Then
                        strAppend = strAppend & "  CISCas.MBMnumber= '" & Trim(MBMNo) & "'"
                    ElseIf tmpLimitStatus = "B" Then
                       If CovID = "00" Then
                           strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID = '00' "
                       Else
                           strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID <> '00' "
                       End If
                    ElseIf tmpLimitStatus = "C" Or tmpLimitStatus = "" Then
                        strAppend = strAppend & "  CISCas.MBMnumber= '" & Trim(MBMNo) & "' AND CoverID ='" + CovID + "' "
                    ElseIf tmpLimitStatus = "D" Then
                       If CovID = "00" Then
                           strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID = '00' "
                       Else
                           strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID <> '00' "
                       End If
                    End If
                End If
             End If
                
            tmpSpecLimitStatus = ""
            SpecLmtRec.Open "SELECT Top 1 PLNMaternityInd,PDentalStatus FROM PLN WHERE PLNCode='" & Trim(strClnPln) & "'", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            If SpecLmtRec.EOF = False Then
                tmpMaternityInd = IIf(Len(SpecLmtRec!PLNMaternityInd), SpecLmtRec!PLNMaternityInd, "")
                tmpDentalInd = IIf(Len(SpecLmtRec!PDentalStatus), SpecLmtRec!PDentalStatus, "")
            End If
            SpecLmtRec.Close
            Set SpecLmtRec = Nothing
                
            Dim cmd As New ADODB.Command
            Dim Prm1 As New ADODB.Parameter
            Dim rst2 As New ADODB.Recordset
                
            If Trim(strAppend) = "" Then
               ' tmpMBMFFS = 0
                'tmpMBMCLMAmt = 0
            Else
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "CIS_FFS2013"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                cmd.Parameters.Append Prm1
                rst2.CursorLocation = adUseClient
                rst2.CursorType = adOpenForwardOnly
                rst2.CacheSize = 100
                Set rst2 = cmd.Execute
                Do While Not rst2.EOF
                    dblMeternity = 0
                    dblDental = 0
                    OtherDed = 0
                    thyroid = 0
                   
                    If IsNumeric(rst2!CLMMaternity) Then
                        dblMeternity = CDbl(rst2!CLMMaternity)
                    Else
                        dblMeternity = 0
                    End If
                           
                    If IsNumeric(rst2!CLMDental) Then
                        dblDental = CDbl(rst2!CLMDental)
                    Else
                        dblDental = 0
                    End If
                          
                    If dblDental > 0 Then
                        dblDental = dblDental
                    End If
                        
                    If IsNumeric(rst2!CLMOthersNoDed) Then
                        OtherDed = CDbl(rst2!CLMOthersNoDed)
                    Else
                        OtherDed = 0
                    End If
                           
                    If IsNumeric(rst2!CLMTyphiadJabNoDed) Then
                        thyroid = CDbl(rst2!CLMTyphiadJabNoDed)
                    Else
                        thyroid = 0
                    End If
                         
                    If tmpMaternityInd = "Y" And tmpDentalInd = "Y" Then
                        txtAccumAmt = CDbl(txtAccumAmt) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (dblMeternity + OtherDed + thyroid + dblDental), 0)
                    ElseIf tmpMaternityInd = "Y" And tmpDentalInd <> "Y" Then
                        txtAccumAmt = CDbl(txtAccumAmt) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (dblMeternity + OtherDed + thyroid), 0)
                    ElseIf tmpMaternityInd <> "Y" And tmpDentalInd = "Y" Then
                        txtAccumAmt = CDbl(txtAccumAmt) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (OtherDed + thyroid + dblDental), 0)
                    ElseIf tmpMaternityInd <> "Y" And tmpDentalInd <> "Y" Then
                        txtAccumAmt = CDbl(txtAccumAmt) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (OtherDed + thyroid), 0)
                    End If
                    If ASOInd = "Y" Then
                        txtAccumAmt = UtilisedAmt
                    End If
                rst2.MoveNext
                Loop
                rst2.Close
                Set rst2 = Nothing
                Set cmd = Nothing
                Set Prm1 = Nothing
            End If
        Else
            If Len(MembershipNo) Then
                If Len(MembershipNo) Then
                   If InStr(1, MembershipNo, "-") Then
                      sArr = Split(MembershipNo, "-")
                      MBMNo = sArr(0)
                      CovID = sArr(1)
                   Else
                      MBMNo = MembershipNo
                      CovID = CovPersonID
                   End If
                End If

                strAppend = ""
                If isPrinDepGPShared Then
                    strAppend = strAppend & "  CISCas.MBMnumber= '" & Trim(MBMNo) & "'"
                Else
                    If tmpLimitStatus = "A" Then
                        strAppend = strAppend & "  CISCas.MBMnumber= '" & Trim(MBMNo) & "'"
                    ElseIf tmpLimitStatus = "B" Then
                        If CovID = "00" Then
                            strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID = '00' "
                        Else
                            strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID <> '00' "
                        End If
                    ElseIf tmpLimitStatus = "C" Or tmpLimitStatus = "" Then
                        strAppend = strAppend & "  CISCas.MBMnumber= '" & Trim(MBMNo) & "' AND CoverID ='" + CovID + "' "
                    ElseIf tmpLimitStatus = "D" Then
                        If CovID = "00" Then
                            strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID = '00' "
                        Else
                            strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID <> '00' "
                        End If
                    End If
                End If
            End If
                
            If Trim(strAppend) = "" Then
                'tmpMBMFFS = 0
                'tmpMBMCLMAmt = 0
            Else
                Set cmd.ActiveConnection = MediConnCISDB
                cmd.CommandText = "CIS_FFS2013"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                cmd.Parameters.Append Prm1
                rst2.CursorLocation = adUseClient
                rst2.CursorType = adOpenForwardOnly
                rst2.CacheSize = 100
                Set rst2 = cmd.Execute
                Dim DblValue As Double
            
                Do While Not rst2.EOF
                dblMeternity = 0
                dblDental = 0
                OtherDed = 0
                thyroid = 0

                If IsNumeric(rst2!CLMMaternity) Then
                    dblMeternity = CDbl(rst2!CLMMaternity)
                Else
                    dblMeternity = 0
                End If
                  
                If IsNumeric(rst2!CLMDental) Then
                    dblDental = CDbl(rst2!CLMDental)
                Else
                    dblDental = 0
                End If
                                
                If IsNumeric(rst2!CLMOthersNoDed) Then
                    OtherDed = CDbl(rst2!CLMOthersNoDed)
                Else
                    OtherDed = 0
                End If
                
                If IsNumeric(rst2!CLMTyphiadJabNoDed) Then
                    thyroid = CDbl(rst2!CLMTyphiadJabNoDed)
                Else
                    thyroid = 0
                End If
                        
            If rst2!PLNGPSPShared = "Y" Then
                If tmpMaternityInd = "Y" And tmpDentalInd = "Y" Then
                    txtAccumAmt = CDbl(txtAccumAmt) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (dblMeternity + OtherDed + thyroid + dblDental), 0)
                ElseIf tmpMaternityInd = "Y" And tmpDentalInd <> "Y" Then
                    txtAccumAmt = CDbl(txtAccumAmt) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (dblMeternity + OtherDed + thyroid), 0)
                ElseIf tmpMaternityInd <> "Y" And tmpDentalInd = "Y" Then
                    txtAccumAmt = CDbl(txtAccumAmt) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (OtherDed + thyroid + dblDental), 0)
                ElseIf tmpMaternityInd <> "Y" And tmpDentalInd <> "Y" Then
                    txtAccumAmt = CDbl(txtAccumAmt) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (OtherDed + thyroid), 0)
                End If
                If ASOInd = "Y" Then
                    txtAccumAmt = UtilisedAmt
                End If
            Else
                If rst2!CLNType <> "S" Then
                    If tmpMaternityInd = "Y" And tmpDentalInd = "Y" Then
                        txtAccumAmt = CDbl(txtAccumAmt) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (dblMeternity + OtherDed + thyroid + dblDental), 0)
                    ElseIf tmpMaternityInd = "Y" And tmpDentalInd <> "Y" Then
                        txtAccumAmt = CDbl(txtAccumAmt) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (dblMeternity + OtherDed + thyroid), 0)
                    ElseIf tmpMaternityInd <> "Y" And tmpDentalInd = "Y" Then
                        txtAccumAmt = CDbl(txtAccumAmt) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (OtherDed + thyroid + dblDental), 0)
                    ElseIf tmpMaternityInd <> "Y" And tmpDentalInd <> "Y" Then
                        txtAccumAmt = CDbl(txtAccumAmt) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (OtherDed + thyroid), 0)
                    End If
                    If ASOInd = "Y" Then
                        txtAccumAmt = UtilisedAmt
                    End If
                ElseIf rst2!CLNType = "S" Then
                    Dim bAccLmt As Boolean
                    bAccLmt = False
                    If tmpLimitStatus = "A" Then
                        bAccLmt = True
                    ElseIf (tmpLimitStatus = "B" And CovID = "00" And rst2!coverid = "00") Or (tmpLimitStatus = "B" And CovID <> "00" And rst2!coverid <> "00") Then
                        bAccLmt = True
                    ElseIf (tmpLimitStatus = "C" Or tmpLimitStatus = "") And rst2!coverid = CovID Then
                        bAccLmt = True
                    ElseIf (tmpLimitStatus = "D" And CovID = "00" And rst2!coverid = "00") Or (tmpLimitStatus = "D" And CovID <> "00" And rst2!coverid <> "00") Then
                        bAccLmt = True
                    End If
                    
                    If bAccLmt = True Then
                        If tmpMaternityInd = "Y" And tmpDentalInd = "Y" Then
                            TotalSpecLimit = CDbl(TotalSpecLimit) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (dblMeternity + OtherDed + thyroid + dblDental), 0)
                        ElseIf tmpMaternityInd = "Y" And tmpDentalInd <> "Y" Then
                            TotalSpecLimit = CDbl(TotalSpecLimit) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (dblMeternity + OtherDed + thyroid), 0)
                        ElseIf tmpMaternityInd <> "Y" And tmpDentalInd = "Y" Then
                            TotalSpecLimit = CDbl(TotalSpecLimit) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (OtherDed + thyroid + dblDental), 0)
                        ElseIf tmpMaternityInd <> "Y" And tmpDentalInd <> "Y" Then
                            TotalSpecLimit = CDbl(TotalSpecLimit) + IIf(IsNumeric(rst2!CLMBordxAmt), CDbl(rst2!CLMBordxAmt) - (OtherDed + thyroid), 0)
                        End If
                    End If
                    If ASOInd = "Y" Then
                        txtAccumAmt = UtilisedAmt
                    End If
                End If
            End If
            rst2.MoveNext
                Loop
                'txtAccumAmt = DblValue
                rst2.Close
                Set rst2 = Nothing
                Set cmd = Nothing
                Set Prm1 = Nothing
            End If
        End If
               
        tmpSpecAccLmt = Format(TotalSpecLimit, "#,###.00")
        
        If IsNumeric(tmpSpecLmt) Then
            tmpSpecAnnLmt = Format(tmpSpecAnnLmt, "#,##0.00")
            txtSpecAccLmt = tmpSpecAnnLmt
             If IsNumeric(tmpSpecAccLmt) Then
                txtSpecBal = Format((CDbl(tmpSpecAnnLmt) - CDbl(tmpSpecAccLmt)), "#,##0.00")
            End If
        Else
            txtSpecBal = tmpSpecAnnLmt
            txtSpecAccLmt = tmpSpecAccLmt
        End If
            
        If IsNumeric(tmpAnnLmt) Then
             txtAccumAmt = Format(txtAccumAmt, "#,##0.00")
             TxtBalLimit = Format(CDbl(tmpAnnLmt) + CDbl(tmpTopupAmt) - CDbl(txtAccumAmt), "#,##0.00")
        End If
        ' ------Limit calculations ------

        TxtCardType = strCardType
        txtLTM = strMBMCLNLTM
        txtDentStatus = StrDentStatus
        A7.Enabled = False    '601 Dental
        A30.Enabled = False   '604 Vision
        If Trim(txtDentStatus) = "C" Then ' only for Combine Lmt (Dental)
            A7.Enabled = True    '601 Dental
            A30.Enabled = True   '604 Vision
        ElseIf Trim(txtDentStatus) = "Y" Then
            A7.Enabled = True    '601 Dental
            A30.Enabled = True
        End If
        tmpGRPBordxFees = IIf(Len(strGRPBordxFees), strGRPBordxFees, "")
    End If
    Call AssignData2ListView
    Call WSInfo
        
    If Trim(txtClinicCode) = "NP" Or Trim(txtClinicCode) = "NPS" Then
        txtClinicName.Visible = False
        txtNPClinic.Visible = True
        txtNPClinic.SetFocus
    End If
    
    Call PaymentInfo
    Dim StrSqlSUN As String
    Dim CLNGST As New ADODB.Recordset
    StrSqlSUN = ""
    StrSqlSUN = "Select * from SUNPosting where InvoiceNo = '" & txtCaseNo & "'"
    CLNGST.Open StrSqlSUN, MediConnCISWS, adOpenKeyset, adLockOptimistic
    
    If CLNGST.EOF = False Then
       ' cmdEdit.Enabled = False
        PostStatus = True
        txtpaymentstatus = "True"
    Else
       ' cmdEdit.Enabled = False
        PostStatus = False
    End If
    CLNGST.Close
    Set CLNGST = Nothing
        
       ' Dim StrOpQry As String
       ' Dim RSOpQry As ADODB.Recordset
       '
       ' StrOpQry = "Select CISDB.dbo.CISCas.OPBatchNo from CISDB.dbo.CISCas where  CISDB.dbo.CISCas.CASNumber ='" + txtCaseNo + "'"
       ' Set RSOpQry = New ADODB.Recordset
       ' RSOpQry.Open StrOpQry, MediConnCISDB, adOpenKeyset, adLockOptimistic
       ' If Not RSOpQry.EOF Then
       '       If Len(RSOpQry!OPBatchNo) Then
       '           CmdEdit.Enabled = False
       '       Else
       '           CmdEdit.Enabled = True
       '       End If
       ' End If
    
    End If
If Trim(txtCaseNo) <> "" Then
    Call duplicatecasesIdentify(txtCaseNo)
End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Public Sub duplicatecasesIdentify(txtCaseNo As String)
Const METHOD_NAME = "duplicatecasesIdentify"
On Error GoTo debugs
    Dim Balcmd As New ADODB.Command
    Dim BalPrm1 As New ADODB.Parameter
    Dim BalRes As New ADODB.Recordset
        Set Balcmd.ActiveConnection = MediConnCISDB
        Balcmd.CommandText = "spMemberduplicatecases_Identify"
        Balcmd.CommandType = adCmdStoredProc
        Balcmd.CommandTimeout = 300
        Set BalPrm1 = Balcmd.CreateParameter("CasNumber", adVarChar, adParamInput, 2000, Trim(txtCaseNo))
        Balcmd.Parameters.Append BalPrm1
        BalRes.CursorLocation = adUseClient
        BalRes.CursorType = adOpenForwardOnly
        BalRes.CacheSize = 100
        Set BalRes = Balcmd.Execute
        Dim Message As String
        Dim casno As String
        Dim i As Integer
        i = 0
        Do While Not BalRes.EOF
            
            casno = Trim(casno) + BalRes!CasNumber + ","
            BalRes.MoveNext
             i = i + 1
        Loop
        If i >= 2 Then
        Message = "Multiple cases Found for this Member on same VistDate" + ":" + " " + casno
        MsgBox Message, vbCritical
        BalRes.Close
        Set BalRes = Nothing
        End If

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub


Private Sub CmdMBMInfo_Click()
    FrmINVMBMinfo.Source = 3
    FrmINVMBMinfo.Show vbModal
End Sub

Private Sub CmdNext_Click()
Const METHOD_NAME = "CmdNext_Click"
On Error GoTo debugs
Dim CAS As New ADODB.Recordset
Dim sql As String
    sql = " select CASNumber from VIEW_Cas_MBM Where MBMCLNDataStatus <> 'D' " & _
          " AND MBMCLNStatus <> 'D' AND (HSCLNInd = 'C' OR HSCLNInd = 'M' OR HSCLNInd = 'B')" & _
          " AND CaseType='C' AND CASNumber > '" & RemStr(txtCaseNo) & "' Order by CASNumber "
    
    CAS.MaxRecords = 1
    CAS.Open sql, MediConnCISDB, adOpenForwardOnly, adLockOptimistic
    
    If CAS.EOF = False Then
        txtCaseNo.Text = Trim(CAS!CasNumber)
        CAS.Close
        Set CAS = Nothing
        CmdGo_Click
    Else
        MsgBox " Last Record Reached! ", vbCritical + vbOKOnly
        CAS.Close
        Set CAS = Nothing
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdPrev_Click()
Const METHOD_NAME = "CmdPrev_Click"
On Error GoTo debugs
Dim CAS As New ADODB.Recordset
Dim sql As String
    sql = " select CASNumber from VIEW_Cas_MBM Where MBMCLNDataStatus <> 'D' " & _
          " AND MBMCLNStatus <> 'D' AND (HSCLNInd = 'C' OR HSCLNInd = 'M' OR HSCLNInd = 'B')" & _
          " AND CaseType='C' AND CASNumber < '" & RemStr(txtCaseNo) & "' Order by CASNumber DESC"
    
    CAS.MaxRecords = 1
    CAS.Open sql, MediConnCISDB, adOpenForwardOnly, adLockOptimistic
    
    If CAS.EOF = False Then
        txtCaseNo.Text = Trim(CAS!CasNumber)
        CAS.Close
        Set CAS = Nothing
        CmdGo_Click
    Else
        MsgBox "First Record Reached! " & vbNewLine & "Not a valid Case Number! ", vbCritical + vbOKOnly
        CAS.Close
        Set CAS = Nothing
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdPrint_Click()
    If lstInvoice.ListItems.Count > 0 Then
        Call ExportToExcelError(lstInvoice)
    End If
End Sub

Private Sub cmdSearch_Click()
    frmCasList.Source = 1
    frmCasList.Show vbModal
End Sub

Private Sub cmdSearchICD1_Click()
    frmSearchICD.Source = 11
    frmSearchICD.Show vbModal
End Sub
Private Sub cmdSearchICD1_LostFocus()
    If Len(Trim(txtDiagnosis1)) Then
        Call VerifyNonPayableDiag(txtDiagnosis1)
    End If
End Sub

Private Sub cmdSearchICD2_Click()
    frmSearchICD.Source = 12
    frmSearchICD.Show vbModal
End Sub

Private Sub cmdSearchICD2_LostFocus()
    If Len(Trim(txtDiagnosis2)) Then
        Call VerifyNonPayableDiag(txtDiagnosis2)
    End If
End Sub

Private Sub cmdSearchICD3_Click()
    frmSearchICD.Source = 13
    frmSearchICD.Show vbModal
End Sub

Private Sub cmdSearchICD3_LostFocus()
    If Len(Trim(txtDiagnosis3)) Then
        Call VerifyNonPayableDiag(txtDiagnosis3)
    End If
End Sub

Private Sub cmdSearchICD4_Click()
    frmSearchICD.Source = 14
    frmSearchICD.Show vbModal
End Sub

Private Sub cmdSearchICD4_LostFocus()
    If Len(Trim(txtDiagnosis4)) Then
        Call VerifyNonPayableDiag(txtDiagnosis4)
    End If
End Sub

Private Sub cmdSearchICD5_Click()
Const METHOD_NAME = "cmdSearchICD5_Click"
On Error GoTo debugs
Dim TreatRec As New ADODB.Recordset
Dim TreatSql As String

    frmSearchICD.Source = 15
    frmSearchICD.Show vbModal
    TxtDiagDesc5.Visible = True
    TxtDiagnosisOth.Enabled = False
    If Trim(TxtDiagnosis5) <> "" Then
        TreatSql = "Select ReasonOtherDiag from TreatmentReason WHERE ReasonCode='" & Trim(TxtDiagnosis5) & "'"
        TreatRec.Open TreatSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        Do While Not TreatRec.EOF
            ReasonOtherIND = IIf(Len(Trim(TreatRec!ReasonOtherDiag)), Trim(TreatRec!ReasonOtherDiag), "N")
            If ReasonOtherIND = "Y" Then
                TxtDiagDesc5.Visible = False
                TxtDiagnosisOth.Visible = True
                TxtDiagnosisOth.Enabled = True
                TxtDiagnosisOth.SetFocus
            Else
                TxtDiagnosisOth = ""
            End If
            TreatRec.MoveNext
        Loop
        TreatRec.Close
        Set TreatRec = Nothing
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdSearchICD5_LostFocus()
    If Len(Trim(TxtDiagnosis5)) Then
        Call VerifyNonPayableDiag(TxtDiagnosis5)
    End If
End Sub

Private Sub CmdUpdate_Click()
Const METHOD_NAME = "CmdUpdate_Click"
On Error GoTo debugs
Dim ChKItemRec As New ADODB.Recordset
Dim ChkItemAmt As Double
Dim ChkAmtInd As Boolean
Dim ChkSql As String
Dim tmpchkBal As Double
Dim tmpchkBill As Double
Dim isVisit As Boolean
Dim tmpLmt As String
tmpLmt = ""
isVisit = False

    'New Condition Start
    If UpdateDia = True Then
        Call DiagnosisUpdate(txtCaseNo)
            CmdEdit.Visible = True
            CmdEdit.Enabled = True
    Exit Sub
    End If
    'New Condition End
    Call A29_Change
    If VisitInd = "Y" Or PDentalStatus = "Y" Then
        If txtDiagnosis1 = "211" Or txtDiagnosis1 = "210" Or txtDiagnosis2 = "211" Or txtDiagnosis2 = "210" Or txtDiagnosis3 = "211" Or txtDiagnosis3 = "210" Or txtDiagnosis4 = "211" Or txtDiagnosis4 = "210" Or TxtDiagnosis5 = "211" Or TxtDiagnosis5 = "210" Then
            tmpLmt = Tmpprepostmatellimit
            If lblDispBillAmt <= Tmpprepostmatellimit Then
                isVisit = True
            Else
                Tmpprepostmatellimit = 0
                isVisit = True
            End If
        ElseIf Len(A8) And (A8) <> 0 Then      'Metarnity Limit
            tmpLmt = TmpMEVisitLimit
            If lblDispBillAmt <= TmpMEVisitLimit Then
                isVisit = True
            Else
                TmpMEVisitLimit = 0
                isVisit = True
            End If
        ElseIf txtDiagnosis1 = "303" Or txtDiagnosis2 = "303" Or txtDiagnosis3 = "303" Or txtDiagnosis4 = "303" Or TxtDiagnosis5 = "303" Then      'Dental Limit
            tmpLmt = tmpPDentalAnnLimit
            If lblDispBillAmt <= tmpPDentalAnnLimit Then
                isVisit = True
            Else
                tmpPDentalAnnLimit = 0
                isVisit = True
            End If
        ElseIf Mid(lblClnType, 1, 1) = "G" Then
            tmpLmt = tmpvisit
            If lblDispBillAmt <= tmpvisit Then
                isVisit = True
            Else
                'TmpVisit = 0
                isVisit = True
            End If
        ElseIf Mid(lblClnType, 1, 1) = "S" Then
            tmpLmt = TmpSPVisitLimit
            If lblDispBillAmt <= TmpSPVisitLimit Then
                isVisit = True
            Else
                TmpSPVisitLimit = 0
                isVisit = True
            End If
        End If
    Else
        isVisit = True
    End If
    

    
    If isVisit Then
        frmWorksheet.InvoiceClick = 1
        frmWorksheet.INVCLNForm = 3
        Call getTotal
        ChkAmtInd = False
        ChkItemAmt = 0
        tmpchkBill = 0
        tmpchkBal = 0
        'If IsNumeric(lblCLNBillAmt) Then
        '    tmpchkBal = CDbl(lblCLNBillAmt) - CDbl(lblActCoPay)
        'Else
        '    tmpchkBal = 0
        'End If
        
        tmpchkBill = CDbl(txtCLNBillAmt)
        If Trim(txtMBMNo) = "" Then
            MsgBox "Please keyin Member No ! ", vbCritical
            txtMBMNo.SetFocus
        ElseIf txtVisitDate = "__/__/____" Or txtVisitDate = "" Then
            MsgBox "Please Enter Visit Date! ", vbCritical
            txtVisitDate.SetFocus
        ElseIf Trim(txtClinicName) = "" Then
            MsgBox "Please keyin the Clinic code ! ", vbCritical
            txtClinicCode.SetFocus
        ElseIf Trim(txtClinicCode) = "NP" And Trim(txtNPClinic) = "" Then
            MsgBox "Please keyin the NP Clinic ! ", vbCritical
            txtNPClinic.SetFocus
        ElseIf Trim(txtClinicCode) = "NPS" And Trim(txtNPClinic) = "" Then
            MsgBox "Please keyin the NPS Clinic ! ", vbCritical
            txtNPClinic.SetFocus
        ElseIf Trim(ReasonOtherIND) = "Y" And Trim(TxtDiagnosisOth) = "" Then
            MsgBox "Please Keyin Others Diagnosis ! ", vbCritical
            TxtDiagnosisOth.Enabled = True
            TxtDiagnosisOth.SetFocus
        ElseIf Mid(CboPayto, 1, 1) = "C" And Trim(txtClinicCode) = "NP" Then
            MsgBox "Invalid Clinic Code", vbCritical
            txtClinicCode.SetFocus
        ElseIf Trim(txtCLNBillAmt) = "" And cboMemoClaim = "Y" Then
            MsgBox "Please Enter Bill Amount ! ", vbCritical
            txtCLNBillAmt.SetFocus
        'ElseIf txtCLNBillAmt = 0 And cboMemoClaim = "Y" Then
        '    MsgBox "Bill Amount Amount cannot be Zero ! ", vbCritical
        'ElseIf ((CDbl(txtCLNBillAmt) <> CDbl(lblDispCLNBillAmt)) And (txtCLNBillAmt <> 0)) Then
        '    MsgBox "Please check Bill Amount! ", vbCritical
        '    txtCLNBillAmt.SetFocus
        ElseIf ((CDbl(txtCLNBillAmt) <> CDbl(lblDispCLNBillAmt))) Then
            MsgBox "Please check Bill Amount! ", vbCritical
            txtCLNBillAmt.SetFocus
        ElseIf txtClinicCode = "NPS" And cboRefLetter = "" Then
            MsgBox "Please enter status for referral letter", vbOKOnly
        ElseIf Trim(txtInvNo) <> "" And (txtInvDate = "__/__/____" Or txtInvDate = "01/01/1900") Then
            MsgBox "Please enter invoice date", vbOKOnly
            txtInvDate = "__/__/____"
            txtInvDate.SetFocus
        ElseIf (CheckExceedLimit() = True) And (cboStatus.ListIndex = 1 And cboClaimability.ListIndex) And strGLIssued <> "Y" Then
            MsgBox "Bill Amount exceeded Balance Limit entitlement! Please check the charged Bill Amount should not more than Balance Limit.", vbCritical
        Else
            If IntInvAmt < 0 Then
                MsgBox "Total Invoice Amount cannot be Zero ! ", vbCritical
            End If
            If ChkAmtInd = True Then
                If CDbl(IntInvAmt) > CDbl(ChkItemAmt) Then
                    MsgBox "Invoice Amount > Consultation Medication - " & TxtCardType & "Card"
                End If
            End If
            
            If bPLNBTBG Then
                If txtPlanCode = "G018" Or txtPlanCode = "G019" Then
                    If (CDbl(IntInvAmt) - 60 <> CDbl(txtBTBG.Text)) And Trim(txtBTBGRemark.Text) = "" Then
                        MsgBox "Please enter BTBG Remark ! ", vbCritical
                        txtBTBGRemark.SetFocus
                        Exit Sub
                    End If
                End If
            End If
            
            'If CDbl(txtCLNBillAmt) = 0 Then
            '    If cboStatus.ListIndex = 2 And cboClaimability.ListIndex = 2 Then
            '    Else
            '        MsgBox "Bill Amount '0' Not Valid. Please Check! ", vbCritical
            '        txtCLNBillAmt.SetFocus
            '        Exit Sub
            '    End If
            'End If
            Dim bAllowSave As Boolean
            bAllowSave = True
            
            If Trim(lblClnType) = "S" And Trim(strPayTo) = "C" Then
                Call CheckDoubleEntry
                If strChkDoubleEntry = "Y" Then
                    bAllowSave = False
                    MsgBox "This Case Number information seems created before! Please check this Member Case History to avoid double entry!", vbCritical
                ElseIf strChkDoubleEntry = "C" Then
                    'temporary by pass to allow save
                    'MsgBox "This Invoice Number seems exist on other Case before! Please check this Member Case History to avoid double entry!"
                    bAllowSave = True
                Else
                    bAllowSave = True
                End If
            End If
            
            If bAllowSave Then
                Call ChkData
                If ASOInd = "Y" Then
                    Call UpdateASOLimit(txtCaseNo)
                End If
            End If
        End If
    Else
        MsgBox ("Please check claim amount more than visit limit " + tmpLmt)
    End If
    
    'CmdGo_Click
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub DiagnosisUpdate(txtCaseNo As String)
Const METHOD_NAME = "DiagnosisUpdate"
On Error GoTo debugs

Dim update
Dim strsql As String
Dim INV As New ADODB.Recordset

    MediConnCISDB.beginTrans
    strsql = ""
    strsql = " UPDATE CISDB..CISCAS SET CASTREATREASON1 ='" & Trim(txtDiagnosis1) & "',CASTREATREASON2='" & Trim(txtDiagnosis2) & "',CASTREATREASON3='" & Trim(txtDiagnosis3) & "',CASTREATREASON4='" & Trim(txtDiagnosis4) & "',CASTREATREASON5='" & Trim(TxtDiagnosis5) & "',CASDIAGNOSIS5='" & Trim(TxtDiagnosisOth) & "',CASREMARK='" & Trim(TxtRemark) & "' " & _
             " where CASNumber ='" & Trim(txtCaseNo) & "'"
              MediConnCISDB.Execute strsql
    MediConnCISDB.commitTrans
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
    
End Sub

Private Sub UpdateRecord(txtCaseNo As String)
Const METHOD_NAME = "UpdateRecord"
On Error GoTo debugs

Dim update
Dim strsql As String
Dim INV As New ADODB.Recordset

    MediConnCISDB.beginTrans
    strsql = ""
    strsql = "update CISDB..CISCas set OPBatchNo = '" & Trim(TxtOPBatchNo) & "', OPNoEnterUser =  '" & Logon & "', OPBatchDate = '" & Format(DateTime.Now, "YYYY/MM/DD") & "'" & _
             " where CASNumber ='" & Trim(txtCaseNo) & "'"
              MediConnCISDB.Execute strsql
    MediConnCISDB.commitTrans
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub UpdateEmptyOPbatchno(txtCaseNo As String)
Const METHOD_NAME = "UpdateEmptyOPbatchno"
On Error GoTo debugs

Dim update
Dim strsql As String


    MediConnCISDB.beginTrans
    strsql = ""
    strsql = "update CISDB..CISCas set OPBatchNo = NULL , OPNoEnterUser =  NULL, OPBatchDate = NULL " & _
             " where CASNumber ='" & Trim(txtCaseNo) & "'"
              MediConnCISDB.Execute strsql
    MediConnCISDB.commitTrans
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Function GenerateOPSeq(OPVar)
Const METHOD_NAME = "GenerateOPSeq"
On Error GoTo debugs

Dim OPRec As New ADODB.Recordset
Dim strsql As String
Dim SeqNo As String
Dim SEQ As New ADODB.Recordset
Dim SEQInsert As New ADODB.Recordset
Dim currentNo, i As Integer
Dim SEQNumber, SEQCode As String
Dim CharacterArray
Dim firstChar, currentChar As String
    SEQNumber = "0000001"
    CharacterArray = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    strsql = "Select OPCode,OPBatchSeq from OPBatch Where OPCode = '" & OPVar & "'"
    OPRec.Open strsql, MediConnCISDB, adOpenStatic, adLockOptimistic
    currentNo = 1
    currentChar = "A"
    
    If OPRec.RecordCount = 0 Then
        With OPRec
            .AddNew
                !OPCode = OPVar
                !OPBatchSeq = currentChar & "0000001"
                .update
                SEQNumber = currentChar & "0000001"
            .update
        End With
    Else
        currentNo = Mid(OPRec!OPBatchSeq, 2, 8)
        firstChar = Mid(OPRec!OPBatchSeq, 1, 1)
        Do Until i = 25
            If firstChar = CharacterArray(i) Then
                If currentNo <> "9999999" Then
                    currentChar = CharacterArray(i)
                    currentNo = currentNo + 1
                Else
                    currentChar = CharacterArray(i + 1)
                    currentNo = "0000001"
                End If
                Exit Do
            End If
            i = i + 1
        Loop
        SEQNumber = currentChar & Format(currentNo, "0000000")
        OPRec!OPBatchSeq = SEQNumber
        OPRec.update
    End If

    GenerateOPSeq = "OC" & SEQNumber
    OPRec.Close
    Set OPRec = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function

Public Sub UpdateASOLimit(CasNumber As String)
Const METHOD_NAME = "UpdateASOLimit"
On Error GoTo debugs
    Dim SqlQry As String
    Dim UpdateRS As New ADODB.Recordset
    SqlQry = ""
    SqlQry = "Exec  sp_ASOClaimsUtilization '" + CasNumber + "','NULL','OP'"
    UpdateRS.Open SqlQry, medixmasterconn, adOpenKeyset, adLockOptimistic
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
        End Select
    End If

End Sub

Private Sub Form_Load()
    bDataStatus = True
    bExit = False
    EditFlag = False
    Call ClearEntries
    Call EntriesEnabled(False)
    Call ComboListing
    CmdClearAll.Enabled = True
End Sub

Private Sub ClearEntries()
   On Error Resume Next
   Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            If ctl.Name <> "txtCaseNo" Then
               ctl.Text = ""
            End If
        ElseIf TypeOf ctl Is ComboBox Then
        ElseIf TypeOf ctl Is MaskEdBox Then
            If ctl.Name <> "txtInvDate" Then
                ctl.mask = ""
                ctl.Text = ""
                ctl.mask = "##/##/####"
            End If
            If ctl.Name = "txtVisitTime" Then
                ctl.mask = "##:##"
            End If
        End If
    Next ctl
    txtClinicName = ""
    txtNPClinic.Visible = False
    lblDispBillAmt = ""
    lblDispCLNBillAmt = ""
    lblCLNBillAmt = "0.00"
    lblBillAmt = "0.00"
    lblFFS = ""
    lblPay2CLN = "0.00"
    Z1 = 0
    Z2 = 0
    Z3 = 0
    Z4 = 0
    Z5 = 0
    Z6 = 0
    Z7 = 0
    Z8 = 0
End Sub

Private Sub InvoiceINFO()
Const METHOD_NAME = "InvoiceINFO"
On Error GoTo debugs
Dim CASSql As String
Dim tempCase As String
Dim tmpDiagDesc As String
Dim CISCAS1 As New ADODB.Recordset
Dim StrCLMSql As String
Dim CLmRec As New ADODB.Recordset
lblFFS = ""
Dim IsFrst As Boolean
IsFrst = True

    CASENumber = txtCaseNo
    
    CASSql = "Exec SearchCASOP '" & Trim(txtCaseNo) & "'"
    Set CISCAS1 = New ADODB.Recordset
    CISCAS1.Open CASSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
        
    'CASSql = "SELECT * FROM View_INV_Clm WHERE CASNumber='" & Trim(CASENumber) & "'"
    'CISCAS1.Open CASSql, MediConnCISWS, adOpenForwardOnly, adLockOptimistic
    If CISCAS1.EOF = True Then
        MsgBox "Not Record Found ! ", vbOKOnly + vbCritical, DialogBoxTitle
        tempCase = txtCaseNo
        lstInvoice.ListItems.Clear
        MembershipNo = ""
        CASENumber = ""
        lblGrpCompany = ""
        txtCaseNo = tempCase
        Call ClearMBMInfo
        Call ClearEntries
        txtCaseNo.SetFocus
    Else
        With CISCAS1
            If IsFrst Then
                Dim MBMNum As String
                Dim coverid As String
                MBMNum = !MBMNumber
                coverid = !coverid
                'get the SP and GP limit  'Trinath Code
                Dim GetGPSPLmt As New ADODB.Recordset
                Dim GPSPQry As String
                GPSPQry = ""
                If CovPersonID = "00" Then
                    GPSPQry = "Select PLN.GRPCompany,PLN.PrinGPLimit as GPLimit,PLN.PrinSPLimit as SPLimit  from HISDB..MBMCLNOthers inner join CISDB..PLN on PLN.PLNCode = MBMCLNOthers.PLNType where HISDB..MBMCLNOthers.MBMNumber ='" + MBMNum + "'"
                Else
                    GPSPQry = "Select PLN.GRPCompany,PLN.DepGPLimit as GPLimit,PLN.DepSPLimit as SPLimit  from HISDB..MBMCoveredPersonsCLN  inner join CISDB..PLN on PLN.PLNCode = MBMCoveredPersonsCLN.MBMCCLNPLNCode    where HISDB..MBMCoveredPersonsCLN.MBMCNumber ='" + MBMNum + "' and HISDB..MBMCoveredPersonsCLN.MBMCCoverID ='" + coverid + "'"
                End If
                GetGPSPLmt.Open GPSPQry, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
                If GetGPSPLmt.EOF = False Then
'                    txtGPlimit = IIf(IsNumeric(GetGPSPLmt!GPLimit), GetGPSPLmt!GPLimit, 0)
'                    txtSpLimit = IIf(IsNumeric(GetGPSPLmt!SPLimit), GetGPSPLmt!SPLimit, 0)
                    tmpGrpCompany = IIf(Len(GetGPSPLmt!GRPCompany), GetGPSPLmt!GRPCompany, 0)
                End If
                GetGPSPLmt.Close
                Set GetGPSPLmt = Nothing
                IsFrst = False
            End If
        
        
        
            StrCLMSql = "Select CLMCoPayColl, CLMFinalApprove, ClmNonGPFee, CLMAdjAnnLmt, CLMRefLetter,CLMLimitExceeded, " & _
            "CLMPayableByMedix2C,CLMMCOProcStatus,CLMMCOProcFee, CLMProcMth, CLMProcYR,CLMCoPayCLNColl,CLMBordxMemo,CLMPayableByMedix2M from Claim where CASNumber='" & Trim(CASENumber) & "'"
            CLmRec.Open StrCLMSql, MediConnCISWS, adOpenForwardOnly, adLockOptimistic
            Do While Not CLmRec.EOF
                A31 = IIf(IsNumeric(CLmRec!ClmNonGPFee), CLmRec!ClmNonGPFee, 0)
                txtMBMAdjAnnLmt = IIf(IsNumeric(CLmRec!CLMAdjAnnLmt), CLmRec!CLMAdjAnnLmt, 0)
                tmpPrevNonGPFee = IIf(IsNumeric(CLmRec!ClmNonGPFee), CLmRec!ClmNonGPFee, 0)
                tmpPrevAdjAmt = IIf(IsNumeric(CLmRec!CLMAdjAnnLmt), CLmRec!CLMAdjAnnLmt, 0)
                tmpFinalApp = IIf(IsNumeric(CLmRec!CLMFinalApprove), CLmRec!CLMFinalApprove, 0)
                'lblActCoPay = IIf(Len(CLmRec!CLMCoPayColl), Format(CLmRec!CLMCoPayColl, "#,##0.00"), 0)
                lblActCoPay = IIf(Len(CLmRec!CLMCoPayCLNColl), Format(CLmRec!CLMCoPayCLNColl, "#,##0.00"), 0)
                A29 = IIf(Len(CLmRec!CLMCoPayCLNColl), Format(CLmRec!CLMCoPayCLNColl, "#,##0.00"), 0)
                cboFFS = IIf(Len(CLmRec!CLMMCOProcStatus), CLmRec!CLMMCOProcStatus, "N")
                cboProcMth = IIf(Len(CLmRec!CLMProcMth), CLmRec!CLMProcMth & "/" & CLmRec!CLMProcYR, "")
                lblFFS = IIf(IsNumeric(CLmRec!CLMMCOProcFee), Format(CLmRec!CLMMCOProcFee, "#,##0.00"), 0)
                If !INVPayTo = "C" Then
                    lblPay2CLN = IIf(IsNumeric(CLmRec!CLMPayableByMedix2c), Format(CLmRec!CLMPayableByMedix2c, "#,##0.00"), 0)
                Else
                    lblPay2CLN = IIf(IsNumeric(CLmRec!CLMPayableByMedix2M), Format(CLmRec!CLMPayableByMedix2M, "#,##0.00"), 0)
                End If
                cboMemoClaim = IIf(Len(CLmRec!CLMBordxMemo), CLmRec!CLMBordxMemo, "")
                cboRefLetter = IIf(Len(CLmRec!CLMRefLetter), CLmRec!CLMRefLetter, "")
                txtLmtExceeded = IIf(Len(CLmRec!CLMLimitExceeded), CLmRec!CLMLimitExceeded, "")
                CLmRec.MoveNext
            Loop
            CLmRec.Close
            Set CLmRec = Nothing
            lblGrpCompany = IIf(Len(!MBMGroupCompany), !MBMGroupCompany, "")
            txtMBMNo = !MBMNumber & "-" & !coverid
            MembershipNo = !MBMNumber
            CovPersonID = !coverid
            CASENumber = !CasNumber
            strPatientID = !coverid
            txtInvDate = IIf(IsDate(!InvoiceDate), !InvoiceDate, "__/__/____")
            txtInvNo = IIf(Len(!InvoiceNo), !InvoiceNo, "")
            txtInvRecDate = IIf(Len(!InvoiceReceiptDate), !InvoiceReceiptDate, "__/__/____")
            CboInvRecMode = IIf(Len(!InvoiceReceiptMode), !InvoiceReceiptMode, "")
            txtVisitDate = !DateVisit
            txtMC = IIf(Len(!MCDays), !MCDays, "")
            lblDispBillAmt = IIf(IsNumeric(!INVAmount), Format(!INVAmount, "#,##0.00"), 0)
            lblDispCLNBillAmt = IIf(IsNumeric(!INVAmount), Format(!INVAmount, "#,##0.00"), 0)
            If Len(Trim(!TimeVisit)) Then
                txtVisitTime = Format(!TimeVisit, "hh:mm")
            End If
            txtBillNo = IIf(Len(Trim(!BillNo)), Trim(!BillNo), "")
            txtCLNBillAmt = IIf(IsNumeric(!INVBillAmt), !INVBillAmt, 0)
            LblBordx = !INVBordxed
            txtClinicName.Visible = True
            txtClinicName = IIf(Len(!CLNName), !CLNName, "")
            txtClinicCode = !ClinicCode
            txtNPClinic.Visible = False
            If Trim(!NPClinic) <> "" Then
                txtClinicName.Visible = False
                txtNPClinic.Visible = True
                txtNPClinic = !NPClinic
            End If
            lblClnType = IIf(Len(Trim(!CLNType)), Trim(!CLNType), "G")
            'txtInVIndex = !InvINdex
            If (Trim(!CLNFFSStatus) = Trim("Y")) Then
                  FFSIndicator = True
                  FFSAmount = !CLNMCOProcfeePerc
            Else
                  FFSIndicator = False
                  FFSAmount = 1
            End If
            LblBatchNo = IIf(Len(!OPBatchNo), !OPBatchNo, "")
            strPayTo = !INVPayTo
            CboPayto.ListIndex = 0
            If Trim(!INVPayTo) = "M" Then
                CboPayto.ListIndex = 1
                tmpPayto = "M"
            Else
                tmpPayto = "C"
            End If
            A1 = IIf(Len(!CLMGPFeeBasicMed), !CLMGPFeeBasicMed, 0)
            A2 = IIf(Len(!CLMGPFeeLTMed), !CLMGPFeeLTMed, 0)
            A3 = IIf(Len(!CLMGPFeeMedInj), !CLMGPFeeMedInj, 0)
            A4 = IIf(Len(!CLMSPFeeBasicMed), !CLMSPFeeBasicMed, 0)
            A5 = IIf(Len(!CLMSPFeeLTMed), !CLMSPFeeLTMed, 0)
            A6 = IIf(Len(!CLMSPFeeMedInj), !CLMSPFeeMedInj, 0)
            A7 = IIf(Len(!CLMDental), !CLMDental, 0)
            A8 = IIf(Len(!CLMMaternity), !CLMMaternity, 0)
            A9 = IIf(Len(!CLMPRVCheckup), !CLMPRVCheckup, 0)
            A10 = IIf(Len(!CLMXRay), !CLMXRay, 0)
            A11 = IIf(Len(!CLMUltraSound), !CLMUltraSound, 0)
            A12 = IIf(Len(!CLMECG), !CLMECG, 0)
            A13 = IIf(Len(!CLMDressing), !CLMDressing, 0)
            A14 = IIf(Len(!CLMNebulizer), !CLMNebulizer, 0)
            A15 = IIf(Len(!CLMID), !CLMID, 0)
            A16 = IIf(Len(!CLMTSSTO), !CLMTSSTO, 0)
            A17 = IIf(Len(!CLMMinorSurgery), !CLMMinorSurgery, 0)
            A18 = IIf(Len(!CLMPOthers), !CLMPOthers, 0)
            A19 = IIf(Len(!CLMBloodDrawing), !CLMBloodDrawing, 0)
            A20 = IIf(Len(!CLMBloodGlucose), !CLMBloodGlucose, 0)
            A21 = IIf(Len(!CLMUrine), !CLMUrine, 0)
            A22 = IIf(Len(!CLMPapSmear), !CLMPapSmear, 0)
            A23 = IIf(Len(!CLMLOthers), !CLMLOthers, 0)
            A24 = IIf(Len(!CLMDTPPolio), !CLMDTPPolio, 0)
            A25 = IIf(Len(!CLMDTPPolioHIB), !CLMDTPPolioHIB, 0)
            A26 = IIf(Len(!CLMHepatitisB), !CLMHepatitisB, 0)
            A27 = IIf(Len(!CLMMMR), !CLMMMR, 0)
            A28 = IIf(Len(!CLMIOthers), !CLMIOthers, 0)
            A30 = IIf(Len(!CLMVision), !CLMVision, 0)
            A32 = IIf(Len(!CLMTyphiadJabNoDed), !CLMTyphiadJabNoDed, 0)
            A33 = IIf(Len(!CLMOthersNoDed), !CLMOthersNoDed, 0)
            
            A7_GST = IIf(Len(!CLMDentalGSTAmt), !CLMDentalGSTAmt, 0)
            A8_GST = IIf(Len(!CLMMaternityGSTAmt), !CLMMaternityGSTAmt, 0)
            A9_GST = IIf(Len(!CLMPRVCheckupGSTAmt), !CLMPRVCheckupGSTAmt, 0)
            A10_GST = IIf(Len(!CLMXRayGSTAmt), !CLMXRayGSTAmt, 0)
            A11_GST = IIf(Len(!CLMUltraSoundGSTAmt), !CLMUltraSoundGSTAmt, 0)
            A12_GST = IIf(Len(!CLMECGGSTAmt), !CLMECGGSTAmt, 0)
            A13_GST = IIf(Len(!CLMDressingGSTAmt), !CLMDressingGSTAmt, 0)
            A14_GST = IIf(Len(!CLMNebulizerGSTAmt), !CLMNebulizerGSTAmt, 0)
            A15_GST = IIf(Len(!CLMIDGSTAmt), !CLMIDGSTAmt, 0)
            A16_GST = IIf(Len(!CLMTSSTOGSTAmt), !CLMTSSTOGSTAmt, 0)
            A17_GST = IIf(Len(!CLMMinorSurgeryGSTAmt), !CLMMinorSurgeryGSTAmt, 0)
            A18_GST = IIf(Len(!CLMPOthersGSTAmt), !CLMPOthersGSTAmt, 0)
            A19_GST = IIf(Len(!CLMBloodDrawingGSTAmt), !CLMBloodDrawingGSTAmt, 0)
            A20_GST = IIf(Len(!CLMBloodGlucoseGSTAmt), !CLMBloodGlucoseGSTAmt, 0)
            A21_GST = IIf(Len(!CLMUrineGSTAmt), !CLMUrineGSTAmt, 0)
            A22_GST = IIf(Len(!CLMPapSmearGSTAmt), !CLMPapSmearGSTAmt, 0)
            A23_GST = IIf(Len(!CLMLOthersGSTAmt), !CLMLOthersGSTAmt, 0)
            A24_GST = IIf(Len(!CLMDTPPolioGSTAmt), !CLMDTPPolioGSTAmt, 0)
            A25_GST = IIf(Len(!CLMDTPPolioHIBGSTAmt), !CLMDTPPolioHIBGSTAmt, 0)
            A26_GST = IIf(Len(!CLMHepatitisBGSTAmt), !CLMHepatitisBGSTAmt, 0)
            A27_GST = IIf(Len(!CLMMMRGSTAmt), !CLMMMRGSTAmt, 0)
            A28_GST = IIf(Len(!CLMIOthersGSTAmt), !CLMIOthersGSTAmt, 0)
            A30_GST = IIf(Len(!CLMVisionGSTAmt), !CLMVisionGSTAmt, 0)
            A32_GST = IIf(Len(!CLMTyphiadJabNoDedGSTAmt), !CLMTyphiadJabNoDedGSTAmt, 0)
            A33_GST = IIf(Len(!CLMOthersNoDedGSTAmt), !CLMOthersNoDedGSTAmt, 0)
            
            txtPOthDesc = IIf(Len(Trim(!CLMPOthersDesc)), !CLMPOthersDesc, "")
            txtLOthDesc = IIf(Len(Trim(!CLMLOthersDesc)), !CLMLOthersDesc, "")
            txtIOthDesc = IIf(Len(Trim(!CLMIOthersDesc)), !CLMIOthersDesc, "")
            txtMedType1(0) = IIf(Len(Trim(!MedInjType1)), !MedInjType1, "")
            txtMedType1(1) = IIf(Len(Trim(!MedInjType2)), !MedInjType2, "")
            txtMedType1(2) = IIf(Len(Trim(!MedInjType3)), !MedInjType3, "")
            txtMedType1(3) = IIf(Len(Trim(!MedInjType4)), !MedInjType4, "")
            txtMedType1(4) = IIf(Len(Trim(!MedInjType5)), !MedInjType5, "")
            txtMedType1(5) = IIf(Len(Trim(!MedInjType6)), !MedInjType6, "")
            txtMedType1(6) = IIf(Len(Trim(!MedInjType7)), !MedInjType7, "")
            txtMedDay1(0) = IIf(Len(Trim(!MedInjDays1)), !MedInjDays1, "")
            txtMedDay1(1) = IIf(Len(Trim(!MedInjDays2)), !MedInjDays2, "")
            txtMedDay1(2) = IIf(Len(Trim(!MedInjDays3)), !MedInjDays3, "")
            txtMedDay1(3) = IIf(Len(Trim(!MedInjDays4)), !MedInjDays4, "")
            txtMedDay1(4) = IIf(Len(Trim(!MedInjDays5)), !MedInjDays5, "")
            txtMedDay1(5) = IIf(Len(Trim(!MedInjDays6)), !MedInjDays6, "")
            txtMedDay1(6) = IIf(Len(Trim(!MedInjDays7)), !MedInjDays7, "")
            txtMedQty1(0) = IIf(Len(Trim(!MedInjQty1)), !MedInjQty1, "")
            txtMedQty1(1) = IIf(Len(Trim(!MedInjQty2)), !MedInjQty2, "")
            txtMedQty1(2) = IIf(Len(Trim(!MedInjQty3)), !MedInjQty3, "")
            txtMedQty1(3) = IIf(Len(Trim(!MedInjQty4)), !MedInjQty4, "")
            txtMedQty1(4) = IIf(Len(Trim(!MedInjQty5)), !MedInjQty5, "")
            txtMedQty1(5) = IIf(Len(Trim(!MedInjQty6)), !MedInjQty6, "")
            txtMedQty1(6) = IIf(Len(Trim(!MedInjQty7)), !MedInjQty7, "")
            ''Price
            txtMedQty1(7) = IIf(Len(Trim(!MEDCHARGES1)), !MEDCHARGES1, "")
            txtMedQty1(8) = IIf(Len(Trim(!MEDCHARGES2)), !MEDCHARGES2, "")
            txtMedQty1(9) = IIf(Len(Trim(!MEDCHARGES3)), !MEDCHARGES3, "")
            txtMedQty1(10) = IIf(Len(Trim(!MEDCHARGES4)), !MEDCHARGES4, "")
            txtMedQty1(11) = IIf(Len(Trim(!MEDCHARGES5)), !MEDCHARGES5, "")
            txtMedQty1(12) = IIf(Len(Trim(!MEDCHARGES6)), !MEDCHARGES6, "")
            txtMedQty1(13) = IIf(Len(Trim(!MEDCHARGES7)), !MEDCHARGES7, "")
            
            txtDiagnosis1 = IIf(Len(Trim(!CASTreatReason1)), !CASTreatReason1, "")
            txtDiagnosis2 = IIf(Len(Trim(!CASTreatReason2)), !CASTreatReason2, "")
            txtDiagnosis3 = IIf(Len(Trim(!CASTreatReason3)), !CASTreatReason3, "")
            txtDiagnosis4 = IIf(Len(Trim(!CASTreatReason4)), !CASTreatReason4, "")
            TxtDiagDesc5.Visible = True
            TxtDiagnosis5 = IIf(Len(Trim(!CASTreatReason5)), !CASTreatReason5, "")
            If Trim(txtDiagnosis1) <> "" Then
                Call DiagDesc(txtDiagnosis1, tmpDiagDesc)
                TxtDiagDesc1 = tmpDiagDesc
            End If
            If Trim(txtDiagnosis2) <> "" Then
                Call DiagDesc(txtDiagnosis2, tmpDiagDesc)
                TxtDiagDesc2 = tmpDiagDesc
            End If
            If Trim(txtDiagnosis3) <> "" Then
                Call DiagDesc(txtDiagnosis3, tmpDiagDesc)
                TxtDiagDesc3 = tmpDiagDesc
            End If
            If Trim(txtDiagnosis4) <> "" Then
                Call DiagDesc(txtDiagnosis4, tmpDiagDesc)
                TxtDiagDesc4 = tmpDiagDesc
            End If
            If Trim(TxtDiagnosis5) <> "" Then
                Call DiagDesc(TxtDiagnosis5, tmpDiagDesc)
                TxtDiagDesc5 = tmpDiagDesc
            End If
            ReasonOtherIND = "N"
            TxtDiagnosisOth.Visible = False
            If Trim(!CASDiagnosis5) <> "" Then
                ReasonOtherIND = "Y"
                TxtDiagDesc5.Visible = False
                TxtDiagnosisOth.Visible = True
                TxtDiagnosisOth.Enabled = True
                TxtDiagnosisOth = IIf(Len(Trim(!CASDiagnosis5)), !CASDiagnosis5, "")
            End If
            B1 = IIf(Len(!GPFee), !GPFee, 0)
            B2 = IIf(Len(!GPLTMed), !GPLTMed, 0)
            B3 = IIf(Len(!GPInj), !GPInj, 0)
            B4 = IIf(Len(!GPMed), !GPMed, 0)
            B5 = IIf(Len(!SPFee), !SPFee, 0)
            B6 = IIf(Len(!SPLTMed), !SPLTMed, 0)
            B7 = IIf(Len(!SPInj), !SPInj, 0)
            B8 = IIf(Len(!SPMed), !SPMed, 0)
            Z1 = IIf(Len(!GPFee), !GPFee, 0)
            Z2 = IIf(Len(!GPLTMed), !GPLTMed, 0)
            Z3 = IIf(Len(!GPInj), !GPInj, 0)
            Z4 = IIf(Len(!GPMed), !GPMed, 0)
            Z5 = IIf(Len(!SPFee), !SPFee, 0)
            Z6 = IIf(Len(!SPLTMed), !SPLTMed, 0)
            Z7 = IIf(Len(!SPInj), !SPInj, 0)
            Z8 = IIf(Len(!SPMed), !SPMed, 0)
            C1 = CDbl(B1) + CDbl(B2) + CDbl(B3) + CDbl(B4)
            C2 = CDbl(B5) + CDbl(B6) + CDbl(B7) + CDbl(B8)
            
            B5_GST = IIf(Len(!SPFeeGSTAmt), !SPFeeGSTAmt, 0)
            B6_GST = IIf(Len(!SPLTMedGSTAmt), !SPLTMedGSTAmt, 0)
            B7_GST = IIf(Len(!SPInjGSTAmt), !SPInjGSTAmt, 0)
            B8_GST = IIf(Len(!SPMedGSTAmt), !SPMedGSTAmt, 0)
            C2_GST = CDbl(B5_GST) + CDbl(B6_GST) + CDbl(B7_GST) + CDbl(B8_GST)
            
            If Trim(!CASStatus) = "O" Then
                cboStatus.ListIndex = 1
            ElseIf Trim(!CASStatus) = "C" Then
                cboStatus.ListIndex = 2
            ElseIf Trim(!CASStatus) = "D" Then
                cboStatus.ListIndex = 3
            End If
            
            If Len(!Claimability) Then
                If Trim(!Claimability) = "A" Then
                    cboClaimability.ListIndex = 1
                ElseIf Trim(!Claimability) = "R" Then
                    cboClaimability.ListIndex = 2
                ElseIf Trim(!Claimability) = "C" Then
                    cboClaimability.ListIndex = 3
                ElseIf Trim(!Claimability) = "S" Then
                    cboClaimability.ListIndex = 4
                End If
            End If
            
            If Trim(!Pending) = "Y" Then
                cboPending.Text = "Y - Yes"
            ElseIf Trim(!Pending) = "N" Then
                cboPending.Text = "N - No"
            End If
            TxtRemark = IIf(Len(Trim(!CasRemark)), Trim(!CasRemark), "")
            txtRepReason = IIf(Len(Trim(!RepudiationReason)), Trim(!RepudiationReason), "")
            txtCASRemark = IIf(Len(Trim(!CASMgmtNotes)), Trim(!CASMgmtNotes), "")
            txtRepAction = IIf(Len(Trim(!CASRepAction)), Trim(!CASRepAction), "")
            txtRepAmt = IIf(Len(Trim(!CASRepAmt)), Trim(!CASRepAmt), "")
            txtPriDiag.Enabled = True
            txtPriDiag = Trim(!PrilimDiag)
            txtPriDiag.Locked = True
            txtInternalRemark.Enabled = True
            txtInternalRemark = Trim(!InternalRemarks)
            txtInternalRemark.Locked = True
            txtBTBG.Text = !CLMBTBGAmt
            txtBTBGRemark.Text = IIf(Len(Trim(!BTBGRemarks)), Trim(!BTBGRemarks), "")
            strGLIssued = IIf(Len(Trim(!GLIssued)), Trim(!GLIssued), "N")
            
             If Trim(!IsPharmacy) = "Y" Then
                cmbPharmacy.Text = "Y - Yes"
            ElseIf Trim(!IsPharmacy) = "N" Then
                cmbPharmacy.Text = "N - No"
               ' Call ValidateIbmUser
                '    If bIbmUser Then
                 '         If A29 = "0.00" Then
                  '          Dim sDt As String
                   '         Dim SpValue As Double
                   '         sDt = Year(txtVisitDate)
                   '             If sDt = "2014" Then
                   '               A29 = "5.00"
                   '             End If
                   '     End If
                   ' End If
            End If
        End With
        Call EntriesEnabled(False)
    End If
    CISCAS1.Close
    Set CISCAS1 = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub ClearMBMInfo()
    strMBMno = ""
    strPayor = ""
    strClnPln = ""
    strClnIns = ""
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
    strCardType = ""
    strMBMCLNLTM = ""
    strCoPay = ""
    MembershipNo = ""
    CASENumber = ""
    txtMBMNo = ""
    txtTabMBMNo = ""
    txtMBMName = ""
    TxtPayor = ""
    txtPlanCode = ""
    TxtInsType = ""
    TxtPatientName = ""
    'txtEffDate.mask = "##/##/####"
    'txtExpDate.mask = "##/##/####"
    txtPolStatus = ""
    txtMBMStatus = ""
'    txtdob = ""
    txtPatientIC = ""
    TxtPatientName = ""
    TxtBalLimit = ""
    TxtVisitLimit = ""
    TxtCardType = ""
    txtLTM = ""
End Sub
Private Sub ComboListing()
Const METHOD_NAME = "ComboListing"
On Error GoTo debugs
Dim InvRec As New ADODB.Recordset
Dim InvMth As New ADODB.Recordset
    CboInvRecMode.Clear
    InvRec.Open "Select RecModeCode, RecModeDesc From InvRecMode ", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    Do While Not InvRec.EOF
        CboInvRecMode.AddItem InvRec!RecModeCode & " - " & Trim(InvRec!RecModeDesc)
        InvRec.MoveNext
    Loop
    InvRec.Close
    Set InvRec = Nothing
    
    InvMth.Open "Select INVProcMth, INVProcYR From INVProcMonth ", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    cboProcMth.Clear
    Do While Not InvMth.EOF
        cboProcMth.AddItem InvMth!INVProcMth & "/" & InvMth!INVProcYR
        InvMth.MoveNext
    Loop
    InvMth.Close
    Set InvMth = Nothing
    
    cboStatus.AddItem "--Select--"
    cboStatus.AddItem "O - Open"
    cboStatus.AddItem "C - Close"
    cboStatus.AddItem "D - Delete"
    
    cboClaimability.AddItem "--Select--"
    cboClaimability.AddItem "A - Accepted"
    cboClaimability.AddItem "R - Repudiated"
    cboClaimability.AddItem "C - Cancelled"
    cboClaimability.AddItem "S - Special Approve"
    
    cboPending.AddItem "Y - Yes"
    cboPending.AddItem "N - No"
    
    cmbPharmacy.AddItem "Y - Yes"
    cmbPharmacy.AddItem "N - No"
    
    CboPayto.AddItem "C - CLINIC"
    CboPayto.AddItem "M - MEMBER"
    
    cboFFS.Clear
    cboFFS.AddItem "Y" & " - " & "Yes"
    cboFFS.AddItem "N" & " - " & "No"
    
    cboMemoClaim.AddItem "Y" & "-" & "Yes"
    cboMemoClaim.AddItem "N" & "-" & "No"
    
    cboRefLetter.AddItem "N" & "-" & "No"
    cboRefLetter.AddItem "Y" & "-" & "Yes"
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lblDispBillAmt_Change()
Dim a As Double
Dim B As Double

    a = IIf(IsNumeric(lblDispBillAmt), lblDispBillAmt, 0)
    B = IIf(IsNumeric(lblActCoPay), lblActCoPay, 0)
        'If CDbl(a) <> "0" Then
        'lblBillAmt = Format(CDbl(a) + CDbl(B), "#,##0.00")
         lblBillAmt = Format(CDbl(a) - CDbl(B), "#,##0.00")
         lblPay2CLN = Format(CDbl(a) - CDbl(B) - lblFFS, "#,##0.00")
    'End If

End Sub

Public Sub GetClinicLimit()
Const METHOD_NAME = "GetClinicLimit"
On Error GoTo debugs
GPLimit = 0
SPLimit = 0
GPCoPay = 0
SPCoPay = 0

Dim SqlQry As String
Dim CLNLimit As New ADODB.Recordset
        SqlQry = "select GPMax , SPMax , CLNCode,GPCopay,SPCoPay  from Tbl_CLNGPLimits where CLNCode ='" & Trim(txtClinicCode) & "'"
        CLNLimit.Open SqlQry, MediConnCISDB, adOpenKeyset, adLockOptimistic
        Do While Not CLNLimit.EOF
            GPLimit = CLNLimit!GPMax
            SPLimit = CLNLimit!SPMax
            GPCoPay = CLNLimit!GPCoPay
            SPCoPay = CLNLimit!SPCoPay
        CLNLimit.MoveNext
        Loop
        CLNLimit.Close
        Set CLNLimit = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub lblDispCLNBillAmt_Change()
Dim a As Double
Dim B As Double

    a = IIf(IsNumeric(lblDispCLNBillAmt), lblDispCLNBillAmt, 0)
    
    'If InStr(1, A29, "-") Then
     '    B = IIf(IsNumeric(A29), A29, 0)
      '   B = B * (-1)
    'Else
     '      B = IIf(IsNumeric(A29), A29, 0)
    'End If
    
   ' B = IIf(IsNumeric(A29), A29, 0)
    If CDbl(a) <> "0" Then
        
        'lblCLNBillAmt = Format(CDbl(a) + CDbl(B), "#,##0.00")
        lblCLNBillAmt = Format(CDbl(a) - CDbl(A29), "#,##0.00")
        
    End If
'txtCLNBillAmt = lblCLNBillAmt
End Sub

Private Sub List1_Click()
On Error Resume Next
Dim lRet    As Long
Dim sFile   As String
    
    PictureFileName = DocPath & "Outpatient\" & "Guarantee Letter\" & txtCaseNo & ""
    '\Medical Rpt"
    crt PictureFileName & "\" & List1.List(List1.ListIndex)

    'If Index = 0 Then
        sFile = PictureFileName & "\" & List1.List(List1.ListIndex)
        lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
End Sub
Sub crt(pth As String)
Dim f
f = FreeFile
Open App.path + "\temp.html" For Output As #f
    Print #f, "<html><body><center><img src=""" & pth & """></center></body></html>"
Close #f
DoEvents
'WebBrowser1.Navigate App.Path & "\temp.html"
End Sub

Private Sub List2_Click()
On Error Resume Next
Dim lRet    As Long
Dim sFile   As String
    PictureFileName = DocPath & "Outpatient\" & "SP_ScannedDoc\" & txtCaseNo & ".pdf"
    'crt PictureFileName & "\" & List2.List(List2.ListIndex)
    sFile = PictureFileName
    lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
End Sub

Private Sub List3_Click()
On Error Resume Next
Dim lRet    As Long
Dim sFile   As String
    
    Dim ParentFolder As String
    ParentFolder = DocPath & "Outpatient\" & "Reim_ScannedDoc\"
    
    Dim CaseFolder As String
    CaseFolder = DocPath & "Outpatient\" & "Reim_ScannedDoc\" & txtCaseNo
    
    Dim myFSO As Object
    Dim FolderExists As String
    Set myFSO = CreateObject("Scripting.FileSystemObject")
    FolderExists = myFSO.FolderExists(CaseFolder)

    If FolderExists = True Then
        sFile = CaseFolder & "\" & List3.List(List3.ListIndex)
    Else
        sFile = ParentFolder & "\" & List3.List(List3.ListIndex)
    End If
    lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
 If SSTab1.Tab = 3 Then
        Call LoadDoc
        Call LoadIntDoc
        Call LoadWebDoc
    End If
End Sub
Private Sub LoadDoc()
On Error Resume Next
Dim B As String, f As Integer, d As Integer, e As Integer
Call addTheFiles
Exit Sub
End Sub
Private Sub LoadIntDoc()
On Error Resume Next
Dim B As String, f As Integer, d As Integer, e As Integer
Call addTheIntFiles
Exit Sub
End Sub
Private Sub LoadWebDoc()
On Error Resume Next
Dim B As String, f As Integer, d As Integer, e As Integer
Call addTheIntFiles1
Exit Sub
End Sub
Sub addTheIntFiles1()
Dim B As String
List3.Clear

'1. Create file name.
'2. Open folder
'3. check file exist in folder if True then add to list

Dim ParentFolder  As String
ParentFolder = DocPath & "Outpatient\" & "Reim_ScannedDoc\" & txtCaseNo

Dim myFSO As Object
Dim FolderExists As String
Set myFSO = CreateObject("Scripting.FileSystemObject")
FolderExists = myFSO.FolderExists(ParentFolder)

If FolderExists = True Then

    B = Dir(ParentFolder & "\*.*")
    
    Do While B > ""
      List3.AddItem B
      B = Dir()
    Loop

Else
    
    Dim fname As String
    fname = DocPath & "Outpatient\" & "Reim_ScannedDoc\" & txtCaseNo & ".pdf"
    If Dir(fname) <> "" Then
        Dim tmpFname As String
        tmpFname = txtCaseNo & ".pdf"
        List3.AddItem tmpFname
    End If
    fname = DocPath & "Outpatient\" & "Reim_ScannedDoc\" & txtCaseNo & ".jpg"
    If Dir(fname) <> "" Then
        tmpFname = txtCaseNo & ".jpg"
        List3.AddItem tmpFname
    End If
    
End If

'List1.Height = Me.Height - 2520 - 350
End Sub
Sub addTheFiles()
Dim B As String
List1.Clear
PictureFileName = DocPath & "Outpatient\" & "Guarantee Letter\" & txtCaseNo & ""
'\Medical Rpt"
B = Dir(PictureFileName & "\*.*")
'b = Dir(Dir1.Path & "\*.*")
If B <> "" Then
    If Right(B, 4) = ".jpg" Then
        List1.AddItem B
    End If
    If Right(B, 4) = ".tif" Then
        List1.AddItem B
    End If
    'If Right(b, 4) = ".art" Then
    '    List1.AddItem b
    'End If
    If Right(B, 4) = ".gif" Then
        List1.AddItem B
    End If
    
    If Right(B, 4) = ".JPG" Then
        List1.AddItem B
    End If
    If Right(B, 4) = ".TIF" Then
        List1.AddItem B
    End If
    'If Right(b, 4) = ".art" Then
    '    List1.AddItem b
    'End If
    If Right(B, 4) = ".GIF" Then
        List1.AddItem B
    End If
    
    If Right(B, 4) = ".doc" Then
        List1.AddItem B
    End If

    If Right(B, 4) = ".pdf" Then
        List1.AddItem B
    End If
    
    If UCase(Right(B, 4)) = ".PNG" Then
        List1.AddItem B
    End If
    
    If Right(B, 4) = ".msg" Then
        List1.AddItem B
    End If
    
Else
    Exit Sub
End If
Do: DoEvents
    B = Dir
    If B <> "" Then
        If Right(B, 4) = ".jpg" Then
            List1.AddItem B
        End If
        If Right(B, 4) = ".tif" Then
            List1.AddItem B
        End If
        'If Right(b, 4) = ".art" Then
        '    List1.AddItem b
        'End If
        If Right(B, 4) = ".gif" Then
            List1.AddItem B
        End If
        
        If Right(B, 4) = ".JPG" Then
            List1.AddItem B
        End If
        If Right(B, 4) = ".TIF" Then
            List1.AddItem B
        End If
        'If Right(b, 4) = ".art" Then
        '    List1.AddItem b
        'End If
        If Right(B, 4) = ".GIF" Then
            List1.AddItem B
        End If
        
        If Right(B, 4) = ".doc" Then
            List1.AddItem B
        End If
        
        If Right(B, 4) = ".pdf" Then
            List1.AddItem B
        End If
        If UCase(Right(B, 4)) = ".PNG" Then
            List1.AddItem B
        End If
        
         If Right(B, 4) = ".msg" Then
            List1.AddItem B
        End If
    
    Else
        Exit Do
    End If
Loop
'List1.Height = Me.Height - 2520 - 350
End Sub
Sub addTheIntFiles()
Dim B As String
List2.Clear

'1. Create file name.
'2. Open folder
'3. check file exist in folder if True then add to list
Dim fname As String
fname = DocPath & "Outpatient\" & "SP_ScannedDoc\" & txtCaseNo & ".pdf"
If Dir(fname) <> "" Then
    Dim tmpFname As String
    tmpFname = txtCaseNo & ".pdf"
    List2.AddItem tmpFname
End If
fname = DocPath & "Outpatient\" & "SP_ScannedDoc\" & txtCaseNo & ".jpg"
If Dir(fname) <> "" Then
    tmpFname = txtCaseNo & ".jpg"
    List2.AddItem tmpFname
End If
'List1.Height = Me.Height - 2520 - 350
End Sub

Private Sub txtCaseNo_GotFocus()
    cmdGo.Default = True
End Sub

Private Sub txtCaseNo_LostFocus()
    cmdGo.Default = False
    txtCaseNo = ChkStr(txtCaseNo)
End Sub

Private Sub TxtClinicCode_GotFocus()
    B1.Enabled = True
    B2.Enabled = True
    B3.Enabled = True
    B4.Enabled = True
    B5.Enabled = False
    B6.Enabled = False
    B7.Enabled = False
    B8.Enabled = False
    B5_GST.Enabled = False
    B6_GST.Enabled = False
    B7_GST.Enabled = False
    B8_GST.Enabled = False
    A7_GST.Enabled = False
    A8_GST.Enabled = False
    A9_GST.Enabled = False
    A10_GST.Enabled = False
    A11_GST.Enabled = False
    A12_GST.Enabled = False
    A13_GST.Enabled = False
    A14_GST.Enabled = False
    A15_GST.Enabled = False
    A16_GST.Enabled = False
    A17_GST.Enabled = False
    A18_GST.Enabled = False
    A19_GST.Enabled = False
    A20_GST.Enabled = False
    A21_GST.Enabled = False
    A22_GST.Enabled = False
    A23_GST.Enabled = False
    A24_GST.Enabled = False
    A25_GST.Enabled = False
    A26_GST.Enabled = False
    A27_GST.Enabled = False
    A28_GST.Enabled = False
    A30_GST.Enabled = False
    A32_GST.Enabled = False
    A33_GST.Enabled = False
    If Trim(lblClnType) = "S" Then
        B1.Enabled = False
        B2.Enabled = False
        B3.Enabled = False
        B4.Enabled = False
        B5.Enabled = True
        B6.Enabled = True
        B7.Enabled = True
        B8.Enabled = True
        B5_GST.Enabled = True
        B6_GST.Enabled = True
        B7_GST.Enabled = True
        B8_GST.Enabled = True
        A7_GST.Enabled = True
        A8_GST.Enabled = True
        A9_GST.Enabled = True
        A10_GST.Enabled = True
        A11_GST.Enabled = True
        A12_GST.Enabled = True
        A13_GST.Enabled = True
        A14_GST.Enabled = True
        A15_GST.Enabled = True
        A16_GST.Enabled = True
        A17_GST.Enabled = True
        A18_GST.Enabled = True
        A19_GST.Enabled = True
        A20_GST.Enabled = True
        A21_GST.Enabled = True
        A22_GST.Enabled = True
        A23_GST.Enabled = True
        A24_GST.Enabled = True
        A25_GST.Enabled = True
        A26_GST.Enabled = True
        A27_GST.Enabled = True
        A28_GST.Enabled = True
        A30_GST.Enabled = True
        A32_GST.Enabled = True
        A33_GST.Enabled = True
    End If
End Sub

Private Sub TxtClinicCode_LostFocus()
Const METHOD_NAME = "TxtClinicCode_LostFocus"
On Error GoTo debugs
Dim CLNRec As New ADODB.Recordset
Dim CLNSql As String
Dim UpCln
    
    txtClinicCode = ChkStr(txtClinicCode)
    txtNPClinic.Visible = False
    If Trim(txtClinicCode) <> "" Then
        txtClinicName = ""
        lblClnType = "G"
        
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        CLNSql = "SELECT CLNPanel, CLNCode, CLNName, CLNType,CLNEtiqaDebtor FROM CLN WHERE CLNCode ='" & Trim(txtClinicCode) & "'"

        CLNRec.Open CLNSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If CLNRec.RecordCount <= 0 Then
            MsgBox "No such Clinic Code ! ", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf Trim(CLNRec!CLNPanel) = "C" Or Trim(CLNRec!CLNPanel) = "Y" Then
            txtClinicName = IIf(Len(CLNRec!CLNName), CLNRec!CLNName, "")
            lblClnType = IIf(Len(Trim(CLNRec!CLNType)), Trim(CLNRec!CLNType), "G")
            If Trim(txtMBMNo) <> "" Then
                Call EntriesEnabled(True)
            End If

            If Trim(txtClinicCode) = "NP" Or Trim(txtClinicCode) = "NPS" Then
                txtClinicName.Visible = False
                txtNPClinic.Visible = True
                txtNPClinic.SetFocus
            End If
        ElseIf CLNRec!CLNEtiqaDebtor <> "" Then
            txtClinicName = IIf(Len(CLNRec!CLNName), CLNRec!CLNName, "")
            lblClnType = IIf(Len(Trim(CLNRec!CLNType)), Trim(CLNRec!CLNType), "G")
            If Trim(txtMBMNo) <> "" Then
                Call EntriesEnabled(True)
            End If

            If Trim(txtClinicCode) = "NP" Or Trim(txtClinicCode) = "NPS" Then
                txtClinicName.Visible = False
                txtNPClinic.Visible = True
                txtNPClinic.SetFocus
            End If
        Else
            UpCln = MsgBox("Non-Panel Clinic !, Press 'Yes' to Proceed, 'No' to Cancel", vbYesNo + vbExclamation)
            If UpCln = vbYes Then
                If CheckAccess(137) = True Then 'Plan
                    If FormLoadedByName("frmMaintenanceEntity") = True Then
                        frmMaintenanceEntity.SetFocus
                    Else
                        frmMaintenanceEntity.Show
                    End If
                Else
                   MsgBox "Access Denied!", vbOKOnly + vbCritical, DialogBoxTitle
                End If
            End If
        End If
        CLNRec.Close
        Set CLNRec = Nothing
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub txtDiagnosis1_LostFocus()
Dim tmpDesc As String
    tmpDesc = ""
    If Trim(txtDiagnosis1) <> "" Then
            Call VerifyDiagnosis(txtDiagnosis1, tmpDesc)
            If tmpDesc <> "" Then
                TxtDiagDesc1 = tmpDesc
            End If
            Else
            TxtDiagDesc1 = ""
    End If
End Sub

Private Sub txtDiagnosis2_LostFocus()
Dim tmpDesc As String
    
    tmpDesc = ""
    If Trim(txtDiagnosis2) <> "" Then
            Call VerifyDiagnosis(txtDiagnosis2, tmpDesc)
            If tmpDesc <> "" Then
                TxtDiagDesc2 = tmpDesc
            End If
            Else
            TxtDiagDesc2 = ""
    End If
End Sub

Private Sub txtDiagnosis3_LostFocus()
Dim tmpDesc As String
    tmpDesc = ""
    If Trim(txtDiagnosis3) <> "" Then
            Call VerifyDiagnosis(txtDiagnosis3, tmpDesc)
            If tmpDesc <> "" Then
                TxtDiagDesc3 = tmpDesc
            End If
            Else
            TxtDiagDesc3 = ""
    End If
End Sub

Private Sub txtDiagnosis4_LostFocus()
Dim tmpDesc As String
    tmpDesc = ""
    If Trim(txtDiagnosis4) <> "" Then
            Call VerifyDiagnosis(txtDiagnosis4, tmpDesc)
            If tmpDesc <> "" Then
                TxtDiagDesc4 = tmpDesc
            End If
            Else
            TxtDiagDesc4 = ""
    End If
End Sub

Private Sub TxtDiagnosis5_LostFocus()
Const METHOD_NAME = "TxtDiagnosis5_LostFocus"
On Error GoTo debugs
Dim tmpDesc As String
Dim TreatRec As New ADODB.Recordset
Dim TreatSql As String
    TxtDiagnosisOth.Enabled = False
    TxtDiagDesc5.Visible = True
    tmpDesc = ""
    If Trim(TxtDiagnosis5) = "" Then
        TxtDiagDesc5 = ""
        TxtDiagnosisOth = ""
    Else
            Call VerifyDiagnosis(TxtDiagnosis5, tmpDesc)
            If tmpDesc <> "" Then
                TxtDiagDesc5 = tmpDesc
            End If
            TreatSql = "Select ReasonOtherDiag from TreatmentReason WHERE ReasonCode='" & Trim(TxtDiagnosis5) & "'"
            TreatRec.Open TreatSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            If TreatRec.RecordCount Then
                ReasonOtherIND = IIf(Len(Trim(TreatRec!ReasonOtherDiag)), Trim(TreatRec!ReasonOtherDiag), "N")
                If ReasonOtherIND = "Y" Then
                    TxtDiagnosisOth.Visible = True
                    TxtDiagDesc5.Visible = False
                    TxtDiagnosisOth.Enabled = True
                    TxtDiagnosisOth.SetFocus
                Else
                    TxtDiagnosisOth.Enabled = True
                    TxtDiagnosisOth = ""
                End If
            End If
            TreatRec.Close
            Set TreatRec = Nothing
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub txtBillNo_LostFocus()
    txtBillNo = ChkStr(txtBillNo)
End Sub

Private Sub TxtDiagnosisOth_LostFocus()
    TxtDiagnosisOth = ChkStr(TxtDiagnosisOth)
End Sub

Private Sub txtIOthDesc_LostFocus()
    txtIOthDesc = ChkStr(txtIOthDesc)
End Sub

Private Sub txtLOthDesc_LostFocus()
    txtLOthDesc = ChkStr(txtLOthDesc)
End Sub

Private Sub txtMBMAdjAnnLmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtMBMAdjAnnLmt_LostFocus()
Dim tmpAdjAmt As String
Dim tmpBillAmt As String
Dim tmpBalLmt As String
Dim tmpNonGP As Double

    tmpAdjAmt = CDbl(IIf(IsNumeric(txtMBMAdjAnnLmt), txtMBMAdjAnnLmt, 0))
    If tmpAdjAmt <= 0 Then
        If IsNumeric(TxtBalLimit) Then
            tmpBillAmt = IIf(IsNumeric(lblDispBillAmt), lblDispBillAmt, 0)
            tmpNonGP = IIf(IsNumeric(A31), A31, 0)
            If CDbl(tmpBillAmt - tmpNonGP) > CDbl(TxtBalLimit) Then
                tmpAdjAmt = CDbl(CDbl(tmpBillAmt) - tmpNonGP - CDbl(TxtBalLimit))
                txtMBMAdjAnnLmt = tmpAdjAmt
            End If
        End If
    End If
End Sub

Private Sub txtMBMNo_GotFocus()
    cmdGo.Default = True
End Sub

Private Sub txtMBMNo_LostFocus()
    cmdGo.Default = False
End Sub

Private Sub EntriesEnabled(EStatus As Boolean)
    FrameCAS.Enabled = EStatus
    FrameMBM.Enabled = EStatus
    FrameInv.Enabled = EStatus
    FrameDet.Enabled = EStatus
    FrameMedINjDiag.Enabled = EStatus
    If EStatus = True Then
        B1.Enabled = True
        B2.Enabled = True
        B3.Enabled = True
        B4.Enabled = True
        B5.Enabled = False
        B6.Enabled = False
        B7.Enabled = False
        B8.Enabled = False
        B5_GST.Enabled = False
        B6_GST.Enabled = False
        B7_GST.Enabled = False
        B8_GST.Enabled = False
        A7_GST.Enabled = False
        A8_GST.Enabled = False
        A9_GST.Enabled = False
        A10_GST.Enabled = False
        A11_GST.Enabled = False
        A12_GST.Enabled = False
        A13_GST.Enabled = False
        A14_GST.Enabled = False
        A15_GST.Enabled = False
        A16_GST.Enabled = False
        A17_GST.Enabled = False
        A18_GST.Enabled = False
        A19_GST.Enabled = False
        A20_GST.Enabled = False
        A21_GST.Enabled = False
        A22_GST.Enabled = False
        A23_GST.Enabled = False
        A24_GST.Enabled = False
        A25_GST.Enabled = False
        A26_GST.Enabled = False
        A27_GST.Enabled = False
        A28_GST.Enabled = False
        A30_GST.Enabled = False
        A32_GST.Enabled = False
        A33_GST.Enabled = False
        If Trim(lblClnType) = "S" Then
            B1.Enabled = False
            B2.Enabled = False
            B3.Enabled = False
            B4.Enabled = False
            B5.Enabled = True
            B6.Enabled = True
            B7.Enabled = True
            B8.Enabled = True
            B5_GST.Enabled = True
            B6_GST.Enabled = True
            B7_GST.Enabled = True
            B8_GST.Enabled = True
            A7_GST.Enabled = True
            A8_GST.Enabled = True
            A9_GST.Enabled = True
            A10_GST.Enabled = True
            A11_GST.Enabled = True
            A12_GST.Enabled = True
            A13_GST.Enabled = True
            A14_GST.Enabled = True
            A15_GST.Enabled = True
            A16_GST.Enabled = True
            A17_GST.Enabled = True
            A18_GST.Enabled = True
            A19_GST.Enabled = True
            A20_GST.Enabled = True
            A21_GST.Enabled = True
            A22_GST.Enabled = True
            A23_GST.Enabled = True
            A24_GST.Enabled = True
            A25_GST.Enabled = True
            A26_GST.Enabled = True
            A27_GST.Enabled = True
            A28_GST.Enabled = True
            A30_GST.Enabled = True
            A32_GST.Enabled = True
            A33_GST.Enabled = True
        End If
    End If
End Sub


Private Sub txtMC_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub ChkData()
Const METHOD_NAME = "ChkData"
On Error GoTo debugs
Dim CISCasRec As New ADODB.Recordset
Dim CISCasSql, ChkDupKey As String
Dim XX, QQ, BB As Integer
Dim AddNewInv As Boolean
Dim DupInvoice
Dim ArrChk() As String
Dim RecSaved
Dim OPBatchNo_New As String
        Erase ArrChk
        ReDim ArrChk(0)
        QQ = 0
        
        
        
            Call UpdateLstView
            Load frmWorksheet
            Unload frmWorksheet
        CmdEdit.Visible = True
        CmdEdit.Enabled = True
        CmdMBMInfo.Enabled = False
        If lstInvoice.ListItems.Count <> 0 Then
        RecSaved = MsgBox("Are you sure ? ", vbYesNo + vbExclamation)
        If RecSaved = vbYes Then
            For AA = 1 To lstInvoice.ListItems.Count
                Call SaveRecord
            Next AA
            If Trim(txtInvNo) <> "" And LblBatchNo = "" And cboStatus.ListIndex = 1 And cboClaimability.ListIndex = 1 And Mid(cboPending, 1, 1) = "N" Then
                OPBatchNo_New = GenerateOPSeq("OC")
                TxtOPBatchNo = OPBatchNo_New
                Call UpdateRecord(txtCaseNo)
                LblBatchNo = OPBatchNo_New
            ElseIf Trim(txtInvNo) = "" Then
                Call UpdateEmptyOPbatchno(txtCaseNo)
                LblBatchNo = ""
            End If
            MsgBox "Record updated ! ", vbInformation
        End If
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Function getTotal() As Double
Dim ctl  As Control
Dim tmpBillAmt As Double
Dim tmpBalLmt As Double
Dim tmpAdjAmt As Double
Dim tmpNonGP As Double

    A1 = 0
    A2 = 0
    A3 = 0
    If Z2 > 0 Then
        A3 = C1
    Else
        If Z3 > 0 And Z4 > 0 Then
            A2 = C1
        Else
            A1 = C1
        End If
    End If
    
    A4 = 0
    A5 = 0
    A6 = 0
    If Z6 > 0 Then
        A6 = C2
    Else
        If Z7 > 0 And Z8 > 0 Then
            A5 = C2
        Else
            A4 = C2
        End If
    End If
    IntInvAmt = 0
    For Each ctl In Me.Controls
        If Mid(ctl.Name, 1, 1) = "A" And Len(ctl.Name) < 6 Then
            If IsNumeric(ctl.Text) = True Then
                IntInvAmt = IntInvAmt + CDbl(ctl.Text)
            End If
        End If
    Next
 '    If InStr(1, A29, "-") Then
 '        A29 = IIf(IsNumeric(A29), A29, 0)
 '        A29 = A29 * (-1)
 '   Else
 '        A29 = IIf(IsNumeric(A29), A29, 0)
 '   End If
    
    IntInvAmt = Format(CDbl(IntInvAmt) - CDbl(IIf(IsNumeric(A29), A29, 0)), "#,##0.00")
    getTotal = IntInvAmt
    
    txtMBMAdjAnnLmt.Enabled = False
    If IsNumeric(TxtBalLimit) Then
        tmpBillAmt = IIf(IsNumeric(IntInvAmt), IntInvAmt, 0)
        tmpNonGP = IIf(IsNumeric(A31), A31, 0)
        If (CDbl(tmpBillAmt) - CDbl(tmpNonGP)) > CDbl(TxtBalLimit) Then
            txtMBMAdjAnnLmt.Enabled = True
            tmpAdjAmt = CDbl(tmpBillAmt) - CDbl(tmpNonGP) - CDbl(TxtBalLimit)
            txtMBMAdjAnnLmt = Format(tmpAdjAmt, "###0.00")
        Else
            txtMBMAdjAnnLmt = ""
        End If
    End If
End Function

Private Sub AssignData2ListView()
Dim strPayee As String
    Set lstRec = lstInvoice.ListItems.Add(, , "")
    If strPayTo = "C" Then
        strPayee = txtClinicCode
    Else
        strPayee = "MEMBER"
    End If
       
    Call getTotal
    lstRec.SubItems(1) = txtCaseNo
    lstRec.SubItems(2) = MembershipNo
    lstRec.SubItems(3) = CovPersonID
    lstRec.SubItems(4) = TxtPatientName
    lstRec.SubItems(5) = IIf(Len(Trim(txtNPClinic)), Trim(txtNPClinic), "")
    lstRec.SubItems(6) = IIf(Len(Trim(txtInvNo)), Trim(txtInvNo), "")
    lstRec.SubItems(7) = IIf(Len(Trim(txtBillNo)), txtBillNo, "")
    lstRec.SubItems(8) = IIf(IsDate(txtInvDate), txtInvDate, "__/__/____")
    lstRec.SubItems(9) = IIf(IsDate(txtVisitDate), txtVisitDate, "__/__/____")
    lstRec.SubItems(10) = IIf(isTime(txtVisitTime), txtVisitTime, "__:__")
    lstRec.SubItems(11) = IIf(IsNumeric(txtMC), txtMC, 0)
    lstRec.SubItems(12) = txtClinicCode
    lstRec.SubItems(13) = IntInvAmt
    lstRec.SubItems(14) = Logon
    lstRec.SubItems(15) = IIf(IsNumeric(A1), A1, 0)
    lstRec.SubItems(16) = IIf(IsNumeric(A2), A2, 0)
    lstRec.SubItems(17) = IIf(IsNumeric(A3), A3, 0)
    lstRec.SubItems(18) = IIf(IsNumeric(A4), A4, 0)
    lstRec.SubItems(19) = IIf(IsNumeric(A5), A5, 0)
    lstRec.SubItems(20) = IIf(IsNumeric(A6), A6, 0)
    lstRec.SubItems(21) = IIf(IsNumeric(A7), A7, 0)
    lstRec.SubItems(22) = IIf(IsNumeric(A8), A8, 0)
    lstRec.SubItems(23) = IIf(IsNumeric(A9), A9, 0)
    lstRec.SubItems(24) = IIf(IsNumeric(A10), A10, 0)
    lstRec.SubItems(25) = IIf(IsNumeric(A11), A11, 0)
    lstRec.SubItems(26) = IIf(IsNumeric(A12), A12, 0)
    lstRec.SubItems(27) = IIf(IsNumeric(A13), A13, 0)
    lstRec.SubItems(28) = IIf(IsNumeric(A14), A14, 0)
    lstRec.SubItems(29) = IIf(IsNumeric(A15), A15, 0)
    lstRec.SubItems(30) = IIf(IsNumeric(A16), A16, 0)
    lstRec.SubItems(31) = IIf(IsNumeric(A17), A17, 0)
    lstRec.SubItems(32) = IIf(IsNumeric(A18), A18, 0)
    lstRec.SubItems(33) = IIf(IsNumeric(A19), A19, 0)
    lstRec.SubItems(34) = IIf(IsNumeric(A20), A20, 0)
    lstRec.SubItems(35) = IIf(IsNumeric(A21), A21, 0)
    lstRec.SubItems(36) = IIf(IsNumeric(A22), A22, 0)
    lstRec.SubItems(37) = IIf(IsNumeric(A23), A23, 0)
    lstRec.SubItems(38) = IIf(IsNumeric(A24), A24, 0)
    lstRec.SubItems(39) = IIf(IsNumeric(A25), A25, 0)
    lstRec.SubItems(40) = IIf(IsNumeric(A26), A26, 0)
    lstRec.SubItems(41) = IIf(IsNumeric(A27), A27, 0)
    lstRec.SubItems(42) = IIf(IsNumeric(A28), A28, 0)
    lstRec.SubItems(43) = IIf(IsNumeric(lblActCoPay), lblActCoPay, 0)
    lstRec.SubItems(44) = IIf(Len(Trim(txtPOthDesc)), Trim(txtPOthDesc), "")
    lstRec.SubItems(45) = IIf(Len(Trim(txtLOthDesc)), Trim(txtLOthDesc), "")
    lstRec.SubItems(46) = IIf(Len(Trim(txtIOthDesc)), Trim(txtIOthDesc), "")
    lstRec.SubItems(47) = IIf(Len(Trim(txtMedType1(0))), Trim(txtMedType1(0)), "")
    lstRec.SubItems(48) = IIf(Len(Trim(txtMedType1(1))), Trim(txtMedType1(1)), "")
    lstRec.SubItems(49) = IIf(Len(Trim(txtMedType1(2))), Trim(txtMedType1(2)), "")
    lstRec.SubItems(50) = IIf(Len(Trim(txtMedType1(3))), Trim(txtMedType1(3)), "")
    lstRec.SubItems(51) = IIf(Len(Trim(txtMedType1(4))), Trim(txtMedType1(4)), "")
    lstRec.SubItems(52) = IIf(Len(Trim(txtMedDay1(0))), Trim(txtMedDay1(0)), "")
    lstRec.SubItems(53) = IIf(Len(Trim(txtMedDay1(1))), Trim(txtMedDay1(1)), "")
    lstRec.SubItems(54) = IIf(Len(Trim(txtMedDay1(2))), Trim(txtMedDay1(2)), "")
    lstRec.SubItems(55) = IIf(Len(Trim(txtMedDay1(3))), Trim(txtMedDay1(3)), "")
    lstRec.SubItems(56) = IIf(Len(Trim(txtMedDay1(4))), Trim(txtMedDay1(4)), "")
    lstRec.SubItems(57) = IIf(Len(Trim(txtMedQty1(0))), Trim(txtMedQty1(0)), "")
    lstRec.SubItems(58) = IIf(Len(Trim(txtMedQty1(1))), Trim(txtMedQty1(1)), "")
    lstRec.SubItems(59) = IIf(Len(Trim(txtMedQty1(2))), Trim(txtMedQty1(2)), "")
    lstRec.SubItems(60) = IIf(Len(Trim(txtMedQty1(3))), Trim(txtMedQty1(3)), "")
    lstRec.SubItems(61) = IIf(Len(Trim(txtMedQty1(4))), Trim(txtMedQty1(4)), "")
    lstRec.SubItems(62) = IIf(Len(Trim(txtDiagnosis1)), Trim(txtDiagnosis1), "")
    lstRec.SubItems(63) = IIf(Len(Trim(txtDiagnosis2)), Trim(txtDiagnosis2), "")
    lstRec.SubItems(64) = IIf(Len(Trim(txtDiagnosis3)), Trim(txtDiagnosis3), "")
    lstRec.SubItems(65) = IIf(Len(Trim(txtDiagnosis4)), Trim(txtDiagnosis4), "")
    lstRec.SubItems(66) = TxtPayor
    lstRec.SubItems(67) = txtPlanCode
    lstRec.SubItems(68) = TxtInsType
    lstRec.SubItems(69) = strHLTCode
    lstRec.SubItems(70) = TxtBalLimit
    lstRec.SubItems(71) = TxtVisitLimit
    lstRec.SubItems(72) = TxtCardType
    lstRec.SubItems(73) = strVisitIND
    lstRec.SubItems(74) = strAnnualInd
    lstRec.SubItems(75) = TxtCoPay
'    lstRec.SubItems(76) = TxtPolNo
    lstRec.SubItems(77) = strPatientIC
    lstRec.SubItems(78) = txtEffDate
    lstRec.SubItems(79) = txtExpDate
    lstRec.SubItems(80) = strPayTo
    lstRec.SubItems(81) = strPayee
    lstRec.SubItems(82) = SuppLmtType
    lstRec.SubItems(83) = IIf(IsNumeric(B1), B1, 0)
    lstRec.SubItems(84) = IIf(IsNumeric(B2), B2, 0)
    lstRec.SubItems(85) = IIf(IsNumeric(B3), B3, 0)
    lstRec.SubItems(86) = IIf(IsNumeric(B4), B4, 0)
    lstRec.SubItems(87) = IIf(IsNumeric(B5), B5, 0)
    lstRec.SubItems(88) = IIf(IsNumeric(B6), B6, 0)
    lstRec.SubItems(89) = IIf(IsNumeric(B7), B7, 0)
    lstRec.SubItems(90) = IIf(IsNumeric(B8), B8, 0)
    lstRec.SubItems(91) = IIf(IsNumeric(C1), C1, 0)
    lstRec.SubItems(92) = IIf(IsNumeric(C2), C2, 0)
    lstRec.SubItems(93) = IIf(Len(Trim(TxtDiagnosis5)), Trim(TxtDiagnosis5), "")
    lstRec.SubItems(94) = ReasonOtherIND
    If Trim(ReasonOtherIND) = "Y" Then
        lstRec.SubItems(95) = Trim(TxtDiagnosisOth)
    Else
        lstRec.SubItems(95) = ""
    End If
    lstRec.SubItems(96) = IIf(IsNumeric(A30), A30, 0)
    
    lstRec.SubItems(102) = IIf(Len(Trim(txtMedType1(5))), Trim(txtMedType1(5)), "")
    lstRec.SubItems(103) = IIf(Len(Trim(txtMedType1(6))), Trim(txtMedType1(6)), "")
    lstRec.SubItems(104) = IIf(Len(Trim(txtMedDay1(5))), Trim(txtMedDay1(5)), "")
    lstRec.SubItems(105) = IIf(Len(Trim(txtMedDay1(6))), Trim(txtMedDay1(6)), "")
    lstRec.SubItems(106) = IIf(Len(Trim(txtMedQty1(5))), Trim(txtMedQty1(5)), "")
    lstRec.SubItems(107) = IIf(Len(Trim(txtMedQty1(6))), Trim(txtMedQty1(6)), "")
    lstRec.SubItems(110) = IIf(IsDate(txtInvRecDate), txtInvRecDate, "")
    lstRec.SubItems(111) = Trim(Left(CboInvRecMode, IIf(InStr(1, CboInvRecMode, "-") - 1 < 0, Len(CboInvRecMode), InStr(1, CboInvRecMode, "-") - 1)))
    lstRec.SubItems(112) = IIf(Len(Trim(TxtRemark)), Trim(TxtRemark), "")
    lstRec.SubItems(113) = IIf(IsNumeric(txtMBMAdjAnnLmt), txtMBMAdjAnnLmt, 0)
    lstRec.SubItems(114) = IIf(IsNumeric(A31), A31, 0)
    ' -- --- reserved for invoice until (150) - last use 114
    lstRec.SubItems(469) = IIf(Len(cboFFS), Mid(cboFFS, 1, 1), "")
    
    'If InStr(1, A29, "-") Then
    '     A29 = IIf(IsNumeric(A29), A29, 0)
    '     A29 = A29 * (-1)
    'Else
    '       A29 = IIf(IsNumeric(A29), A29, 0)
    'End If
    
    lstRec.SubItems(471) = IIf(IsNumeric(A29), A29, 0)
    lstRec.SubItems(472) = IIf(IsNumeric(txtCLNBillAmt), txtCLNBillAmt, 0)
    lstRec.SubItems(473) = IIf(IsNumeric(A32), A32, 0)
    lstRec.SubItems(474) = IIf(IsNumeric(A33), A33, 0)
    lstRec.SubItems(475) = IIf(Len(tmpGRPBordxFees), tmpGRPBordxFees, "")
    lstRec.SubItems(480) = IIf(IsNumeric(txtLmtExceeded), txtLmtExceeded, 0)
    'lstRec.SubItems(481) = IIf(IsNumeric(txtRepAction), txtRepAction, "")
    ''Price
    lstRec.SubItems(482) = IIf(Len(Trim(txtMedQty1(7))), Trim(txtMedQty1(7)), 0)
    lstRec.SubItems(483) = IIf(Len(Trim(txtMedQty1(8))), Trim(txtMedQty1(8)), 0)
    lstRec.SubItems(484) = IIf(Len(Trim(txtMedQty1(9))), Trim(txtMedQty1(9)), 0)
    lstRec.SubItems(485) = IIf(Len(Trim(txtMedQty1(10))), Trim(txtMedQty1(10)), 0)
    lstRec.SubItems(486) = IIf(Len(Trim(txtMedQty1(11))), Trim(txtMedQty1(11)), 0)
    lstRec.SubItems(487) = IIf(Len(Trim(txtMedQty1(12))), Trim(txtMedQty1(12)), 0)
    lstRec.SubItems(488) = IIf(Len(Trim(txtMedQty1(13))), Trim(txtMedQty1(13)), 0)
    
End Sub

Private Sub txtMedDay1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtMedQty1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtMedType1_LostFocus(Index As Integer)
    txtMedType1(Index) = ChkStr(txtMedType1(Index))
End Sub

Private Sub txtNPClinic_LostFocus()
    txtNPClinic = ChkStr(txtNPClinic)
End Sub

Private Sub txtPOthDesc_LostFocus()
    txtPOthDesc = ChkStr(txtPOthDesc)
End Sub

Private Sub txtInvDate_LostFocus()
    If txtInvDate <> "__/__/____" Then
        If IsDate(txtInvDate) = False Then
            MsgBox "Invalid date format!", vbCritical
            txtInvDate.SetFocus
        ElseIf CDate(txtInvDate) > CDate(Date) Then
            MsgBox "Invoice Date cannot greater than system date ! ", vbCritical
            txtInvDate.SetFocus
        ElseIf CDate(txtInvDate) < CDate(txtEffDate) Then
            MsgBox "Invoice Date MUST BE GREATER than Clinical Effective date ! ", vbCritical
        ElseIf CDate(txtInvDate) > CDate(txtExpDate) + 14 Then
            MsgBox "Invoice date IS GREATER THAN Clinical Expiry date + 14 days ! "
        ElseIf (CDate(Date) - CDate(txtInvDate)) > 45 Then
            MsgBox "45 Days LATE ! ", vbCritical
        End If
    End If
End Sub

Private Sub txtInvNo_LostFocus()
    txtInvNo = ChkStr(txtInvNo)
End Sub

Private Sub TxtRemark_lostfocus()
    TxtRemark = ChkStr(TxtRemark)
End Sub

Private Sub txtVisitDate_LostFocus()
    If txtVisitDate <> "__/__/____" Then
        If IsDate(txtVisitDate) = False Then
            MsgBox "Invalid date format!", vbCritical
            txtVisitDate.SetFocus
        ElseIf CDate(txtVisitDate) > CDate(Date) Then
            MsgBox "Visit Date cannot greater than system date ! ", vbCritical
            txtVisitDate.SetFocus
        ElseIf CDate(txtVisitDate) < CDate(txtEffDate) Then
            MsgBox "Visit date MUST BE AFTER the Effective date ! ", vbCritical
        ElseIf CDate(txtVisitDate) > CDate(txtExpDate) + 14 Then
            MsgBox "Visit date IS GREATER THAN Expiry date + 14 days ! ", vbCritical
        End If
    End If
End Sub

Private Sub txtVisitTime_LostFocus()
    If txtVisitTime <> "__:__" Then
        If isTime(txtVisitTime) = False Then
            MsgBox "Invalid Time Format! ", vbCritical
            txtVisitTime.SetFocus
        End If
    End If
End Sub

Private Sub VerifyNonPayableDiag(VDiag As String)
Const METHOD_NAME = "VerifyNonPayableDiag"
On Error GoTo debugs
Dim Diag As New ADODB.Recordset
Dim DiagSql As String
        DiagSql = "SELECT * FROM DiagNotPayable WHERE PLNCode='" & Trim(txtPlanCode) & "' AND DiagCode='" & Trim(VDiag) & "'"
        Diag.Open DiagSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        Do While Not Diag.EOF
            MsgBox "Diagnosis '" & VDiag & "' is not payable ! ", vbOKOnly + vbCritical, DialogBoxTitle
        Diag.MoveNext
        Loop
    Diag.Close
    Set Diag = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub SaveRecord()
Const METHOD_NAME = "SaveRecord"
On Error GoTo debugs
On Err GoTo ErrorSave
Dim startTrans As Boolean
    
    MediConnCISDB.beginTrans
    MediConnCISWS.beginTrans
        startTrans = True
        Call SaveCase
        Call SaveInvoice
        Call saveActual
        Call saveActualOthers
        Call SaveINVClaim
        Call SaveINVCLNApp
        Call SaveINVCLNContract
        Call SaveINVActual
        Call SaveINVCLNExcess
        Call SaveINVMBMApp
        Call SaveINVMBMExcess
        Call SaveINVMBMLmt
        Call UpdateMBMBalance
        Call SaveCLMMEDCHARGES
        'Call SaveMgmtNote
    MediConnCISDB.commitTrans
    MediConnCISWS.commitTrans
    startTrans = False
    
    CmdClearAll.Enabled = True
    
    Call EntriesEnabled(False)
    CmdEdit.Visible = True
    CmdEdit.Enabled = True
    EditFlag = False
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorSave:
    Screen.MousePointer = vbDefault
    If startTrans = True Then
        MediConnCISDB.RollbackTrans
        MediConnCISWS.RollbackTrans
    End If
    MsgBox Err.Description
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub SaveCase()
Const METHOD_NAME = "SaveCase"
On Error GoTo debugs
Dim SaveCaseSql As String
Dim NoOfDiag As Integer
Dim NoofMedi As Integer
Dim MCStat As Integer
Dim tmpINVMonth, tmpINVYear As String

    SaveCaseSql = ""
    MCStat = 0
    NoOfDiag = 0
    NoofMedi = 0
    
        tmpCASnumber = Trim(lstInvoice.ListItems(AA).SubItems(1))
        tmpCASStatus = Mid(cboStatus, 1, 1)
        tmpClaimability = Mid(cboClaimability, 1, 1)
        tmpPending = Mid(cboPending, 1, 1)
        tmpLastUpdateDate = Format(Date, "yyyy/mm/dd")
        tmpLastUpdateBy = Logon
        tmpClinicCode = lstInvoice.ListItems(AA).SubItems(12)
        If Trim(lstInvoice.ListItems(AA).SubItems(12)) <> "NP" And Trim(lstInvoice.ListItems(AA).SubItems(12)) <> "NPS" Then
            lstInvoice.ListItems(AA).SubItems(5) = ""
        End If
        tmpNPClinic = lstInvoice.ListItems(AA).SubItems(5)
        tmpInvoiceNo = Trim(lstInvoice.ListItems(AA).SubItems(6))
        If lstInvoice.ListItems(AA).SubItems(8) = "__/__/____" Then
            tmpInvoiceDate = ""
        Else
            tmpInvoiceDate = Format(lstInvoice.ListItems(AA).SubItems(8), "yyyy/mm/dd")
        End If
        
        If tmpInvoiceDate <> "" Then
            tmpINVMonth = Format(Month(tmpInvoiceDate), "00")
        Else
            tmpINVMonth = ""
        End If
        
        If tmpInvoiceDate <> "" Then
            tmpINVYear = Year(tmpInvoiceDate)
        Else
            tmpINVYear = ""
        End If
        
        tmpDateVisit = Format(lstInvoice.ListItems(AA).SubItems(9), "yyyy/mm/dd")
        tmpMCDays = lstInvoice.ListItems(AA).SubItems(11)
        If isTime(lstInvoice.ListItems(AA).SubItems(10)) Then
            tmpTimeVisit = Format(lstInvoice.ListItems(AA).SubItems(10), "hh:mm")
        End If
        
        tmpMedInjType1 = lstInvoice.ListItems(AA).SubItems(47)
        tmpMedInjType2 = lstInvoice.ListItems(AA).SubItems(48)
        tmpMedInjType3 = lstInvoice.ListItems(AA).SubItems(49)
        tmpMedInjType4 = lstInvoice.ListItems(AA).SubItems(50)
        tmpMedInjType5 = lstInvoice.ListItems(AA).SubItems(51)
        tmpMedInjDays1 = lstInvoice.ListItems(AA).SubItems(52)
        tmpMedInjDays2 = lstInvoice.ListItems(AA).SubItems(53)
        tmpMedInjDays3 = lstInvoice.ListItems(AA).SubItems(54)
        tmpMedInjDays4 = lstInvoice.ListItems(AA).SubItems(55)
        tmpMedInjDays5 = lstInvoice.ListItems(AA).SubItems(56)
        tmpMedInjQty1 = lstInvoice.ListItems(AA).SubItems(57)
        tmpMedInjQty2 = lstInvoice.ListItems(AA).SubItems(58)
        tmpMedInjQty3 = lstInvoice.ListItems(AA).SubItems(59)
        tmpMedInjQty4 = lstInvoice.ListItems(AA).SubItems(60)
        tmpMedInjQty5 = lstInvoice.ListItems(AA).SubItems(61)
        
        tmpMedInjType6 = lstInvoice.ListItems(AA).SubItems(102)
        tmpMedInjType7 = lstInvoice.ListItems(AA).SubItems(103)
        tmpMedInjDays6 = lstInvoice.ListItems(AA).SubItems(104)
        tmpMedInjDays7 = lstInvoice.ListItems(AA).SubItems(105)
        tmpMedInjQty6 = lstInvoice.ListItems(AA).SubItems(106)
        tmpMedInjQty7 = lstInvoice.ListItems(AA).SubItems(107)
        'If IsDate(lstInvoice.ListItems(AA).SubItems(110)) Then
        tmpInvoiceReceiptDate = Format(lstInvoice.ListItems(AA).SubItems(110), "yyyy/mm/dd")
        'End If
        tmpInvoiceReceiptMode = lstInvoice.ListItems(AA).SubItems(111)
        tmpCasRemark = lstInvoice.ListItems(AA).SubItems(112)
        
        tmpCASTreatReason1 = lstInvoice.ListItems(AA).SubItems(62)
        tmpCASTreatReason2 = lstInvoice.ListItems(AA).SubItems(63)
        tmpCASTreatReason3 = lstInvoice.ListItems(AA).SubItems(64)
        tmpCASTreatReason4 = lstInvoice.ListItems(AA).SubItems(65)
        tmpCASTreatReason5 = lstInvoice.ListItems(AA).SubItems(93)
        tmpCASDiagnosis5 = lstInvoice.ListItems(AA).SubItems(95)
        
        If CDbl(lstInvoice.ListItems(AA).SubItems(11)) > 0 Then
            tmpMCStatus = 1
        Else
            tmpMCStatus = 0
        End If
' ----- Number of Diagnosis calculation
        If Trim(lstInvoice.ListItems(AA).SubItems(62)) <> "" Then
            NoOfDiag = NoOfDiag + 1
        End If
        If Trim(lstInvoice.ListItems(AA).SubItems(63)) <> "" Then
            NoOfDiag = NoOfDiag + 1
        End If
        If Trim(lstInvoice.ListItems(AA).SubItems(64)) <> "" Then
            NoOfDiag = NoOfDiag + 1
        End If
        If Trim(lstInvoice.ListItems(AA).SubItems(65)) <> "" Then
            NoOfDiag = NoOfDiag + 1
        End If
        If Trim(lstInvoice.ListItems(AA).SubItems(93)) <> "" Then
            NoOfDiag = NoOfDiag + 1
        End If
        tmpNoOfDiag = NoOfDiag
' ----- Number of Medication calculation
        If Trim(lstInvoice.ListItems(AA).SubItems(47)) <> "" Then
            NoofMedi = NoofMedi + 1
        End If
        If Trim(lstInvoice.ListItems(AA).SubItems(48)) <> "" Then
            NoofMedi = NoofMedi + 1
        End If
        If Trim(lstInvoice.ListItems(AA).SubItems(49)) <> "" Then
            NoofMedi = NoofMedi + 1
        End If
        If Trim(lstInvoice.ListItems(AA).SubItems(50)) <> "" Then
            NoofMedi = NoofMedi + 1
        End If
        If Trim(lstInvoice.ListItems(AA).SubItems(51)) <> "" Then
            NoofMedi = NoofMedi + 1
        End If
        If Trim(lstInvoice.ListItems(AA).SubItems(102)) <> "" Then
            NoofMedi = NoofMedi + 1
        End If
        If Trim(lstInvoice.ListItems(AA).SubItems(103)) <> "" Then
            NoofMedi = NoofMedi + 1
        End If
        tmpNoOfMedication = NoofMedi
        
        Dim pharmacy As String
        If Mid(cmbPharmacy, 1, 1) = "N" Then
            pharmacy = "N"
        ElseIf Mid(cmbPharmacy, 1, 1) = "Y" Then
            pharmacy = "Y"
        End If
        txtRepAction = Replace(txtRepAction, "'", "''")
        SaveCaseSql = ""
        SaveCaseSql = "Exec INVAdj_SaveCase_2013V3 '" & tmpCASnumber & "','" & tmpCASStatus & "','" & tmpClaimability & "','" & _
                        tmpPending & "','" & tmpLastUpdateDate & "','" & tmpLastUpdateBy & "','" & _
                        tmpClinicCode & "','" & tmpNPClinic & "','" & tmpInvoiceNo & "','" & tmpInvoiceDate & "','" & _
                        tmpDateVisit & "'," & tmpMCDays & ",'" & tmpTimeVisit & "','" & _
                        tmpMedInjType1 & "','" & tmpMedInjType2 & "','" & tmpMedInjType3 & "','" & _
                        tmpMedInjType4 & "','" & tmpMedInjType5 & "','" & tmpMedInjDays1 & "','" & _
                        tmpMedInjDays2 & "','" & tmpMedInjDays3 & "','" & tmpMedInjDays4 & "','" & _
                        tmpMedInjDays5 & "','" & tmpMedInjQty1 & "','" & tmpMedInjQty2 & "','" & _
                        tmpMedInjQty3 & "','" & tmpMedInjQty4 & "','" & tmpMedInjQty5 & "','" & _
                        tmpMedInjType6 & "','" & tmpMedInjType7 & "','" & tmpMedInjDays6 & "','" & _
                        tmpMedInjDays7 & "','" & tmpMedInjQty6 & "','" & tmpMedInjQty7 & "','" & _
                        tmpInvoiceReceiptDate & "','" & tmpInvoiceReceiptMode & "','" & _
                        tmpCasRemark & "','" & tmpCASTreatReason1 & "','" & tmpCASTreatReason2 & "','" & _
                        tmpCASTreatReason3 & "','" & tmpCASTreatReason4 & "','" & tmpCASTreatReason5 & "','" & _
                        tmpCASDiagnosis5 & "'," & tmpMCStatus & "," & tmpNoOfDiag & "," & tmpNoOfMedication & ", '" & IIf(Len(txtRepReason), txtRepReason, "") & "', '" & IIf(Len(txtCASRemark), txtCASRemark, "") & "', '" & IIf(Len(txtRepAction), txtRepAction, "") & "', '" & IIf(IsNumeric(txtRepAmt), txtRepAmt, "0") & "', '" & Trim(tmpINVMonth) & "', '" & Trim(tmpINVYear) & "','" & pharmacy & "'"
        MediConnCISDB.Execute SaveCaseSql
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub SaveInvoice()
Const METHOD_NAME = "SaveInvoice"
On Error GoTo debugs

Dim StrInvSql As String
        
        StrInvSql = ""
        tmpCASnumber = Trim(lstInvoice.ListItems(AA).SubItems(1))
        tmpINVLastUpdateBy = Logon
        tmpINVLastUpdateDate = Format(Date, "yyyy/mm/dd")
        tmpBillNo = lstInvoice.ListItems(AA).SubItems(7)
        tmpBillDate = Format(lstInvoice.ListItems(AA).SubItems(9), "yyyy/mm/dd")
        tmpINVAmount = Format(lstInvoice.ListItems(AA).SubItems(13), "#,##0.00")
        tmpINVPayTo = lstInvoice.ListItems(AA).SubItems(80)
        tmpINVPayee = lstInvoice.ListItems(AA).SubItems(81)
        tmpINVBillAmt = Format(lstInvoice.ListItems(AA).SubItems(472), "#,##0.00")
        
        StrInvSql = "Exec INVAdj_SaveInvoice '" & tmpCASnumber & "','" & tmpINVLastUpdateBy & "','" & _
                    tmpINVLastUpdateDate & "','" & tmpBillNo & "','" & tmpBillDate & "'," & _
                    tmpINVAmount & ",'" & tmpINVPayTo & "','" & tmpINVPayee & "'," & tmpINVBillAmt & ""
        MediConnCISWS.Execute StrInvSql
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub saveActual()
Const METHOD_NAME = "saveActual"
On Error GoTo debugs
Dim strsql As String
        strsql = ""
        tmpCASnumber = Trim(lstInvoice.ListItems(AA).SubItems(1))
        tmpCLMGPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(15), "#,##0.00")
        tmpCLMGPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(16), "#,##0.00")
        tmpCLMGPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(17), "#,##0.00")
        tmpCLMSPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(18), "#,##0.00")
        tmpCLMSPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(19), "#,##0.00")
        tmpCLMSPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(20), "#,##0.00")
        tmpCLMDental = Format(lstInvoice.ListItems(AA).SubItems(21), "#,##0.00")
        tmpCLMMaternity = Format(lstInvoice.ListItems(AA).SubItems(22), "#,##0.00")
        tmpCLMPRVCheckup = Format(lstInvoice.ListItems(AA).SubItems(23), "#,##0.00")
        tmpCLMXRay = Format(lstInvoice.ListItems(AA).SubItems(24), "#,##0.00")
        tmpCLMUltraSound = Format(lstInvoice.ListItems(AA).SubItems(25), "#,##0.00")
        tmpCLMECG = Format(lstInvoice.ListItems(AA).SubItems(26), "#,##0.00")
        tmpCLMDressing = Format(lstInvoice.ListItems(AA).SubItems(27), "#,##0.00")
        tmpCLMNebulizer = Format(lstInvoice.ListItems(AA).SubItems(28), "#,##0.00")
        tmpCLMID = Format(lstInvoice.ListItems(AA).SubItems(29), "#,##0.00")
        tmpCLMTSSTO = Format(lstInvoice.ListItems(AA).SubItems(30), "#,##0.00")
        tmpCLMMinorSurgery = Format(lstInvoice.ListItems(AA).SubItems(31), "#,##0.00")
        tmpCLMPOthers = Format(lstInvoice.ListItems(AA).SubItems(32), "#,##0.00")
        tmpCLMBloodDrawing = Format(lstInvoice.ListItems(AA).SubItems(33), "#,##0.00")
        tmpCLMBloodGlucose = Format(lstInvoice.ListItems(AA).SubItems(34), "#,##0.00")
        tmpCLMUrine = Format(lstInvoice.ListItems(AA).SubItems(35), "#,##0.00")
        tmpCLMPapSmear = Format(lstInvoice.ListItems(AA).SubItems(36), "#,##0.00")
        tmpCLMLOthers = Format(lstInvoice.ListItems(AA).SubItems(37), "#,##0.00")
        tmpCLMDTPPolio = Format(lstInvoice.ListItems(AA).SubItems(38), "#,##0.00")
        tmpCLMDTPPolioHIB = Format(lstInvoice.ListItems(AA).SubItems(39), "#,##0.00")
        tmpCLMHepatitisB = Format(lstInvoice.ListItems(AA).SubItems(40), "#,##0.00")
        tmpCLMMMR = Format(lstInvoice.ListItems(AA).SubItems(41), "#,##0.00")
        tmpCLMIOthers = Format(lstInvoice.ListItems(AA).SubItems(42), "#,##0.00")
        tmpCLMCoPayment = Format(lstInvoice.ListItems(AA).SubItems(43), "#,##0.00")
        tmpCLMPOthersDesc = lstInvoice.ListItems(AA).SubItems(44)
        tmpCLMLOthersDesc = lstInvoice.ListItems(AA).SubItems(45)
        tmpCLMIOthersDesc = lstInvoice.ListItems(AA).SubItems(46)
        tmpCLMVision = Format(lstInvoice.ListItems(AA).SubItems(96), "#,##0.00")
        tmpCLMTyphiad = Format(lstInvoice.ListItems(AA).SubItems(473), "#,##0.00")
        tmpCLMOtherNoBalDed = Format(lstInvoice.ListItems(AA).SubItems(474), "#,##0.00")

        strsql = "Exec INVAdj_saveActual '" & tmpCASnumber & "'," & tmpCLMGPFeeBasicMed & "," & _
                tmpCLMGPFeeLTMed & "," & tmpCLMGPFeeMedInj & "," & tmpCLMSPFeeBasicMed & "," & _
                tmpCLMSPFeeLTMed & "," & tmpCLMSPFeeMedInj & "," & tmpCLMDental & "," & _
                tmpCLMMaternity & "," & tmpCLMPRVCheckup & "," & tmpCLMXRay & "," & _
                tmpCLMUltraSound & "," & tmpCLMECG & "," & tmpCLMDressing & "," & _
                tmpCLMNebulizer & "," & tmpCLMID & "," & tmpCLMTSSTO & "," & _
                tmpCLMMinorSurgery & "," & tmpCLMPOthers & "," & tmpCLMBloodDrawing & "," & _
                tmpCLMBloodGlucose & "," & tmpCLMUrine & "," & tmpCLMPapSmear & "," & _
                tmpCLMLOthers & "," & tmpCLMDTPPolio & "," & tmpCLMDTPPolioHIB & "," & _
                tmpCLMHepatitisB & "," & tmpCLMMMR & "," & tmpCLMIOthers & "," & _
                tmpCLMCoPayment & ",'" & tmpCLMPOthersDesc & "','" & tmpCLMLOthersDesc & "','" & _
                tmpCLMIOthersDesc & "'," & tmpCLMVision & ", " & tmpCLMTyphiad & ", " & tmpCLMOtherNoBalDed & ", " & _
                "0, 0, 0, " & IIf(IsNumeric(A7_GST), A7_GST, 0) & ", " & IIf(IsNumeric(A8_GST), A8_GST, 0) & ", " & IIf(IsNumeric(A9_GST), A9_GST, 0) & ", " & _
                IIf(IsNumeric(A10_GST), A10_GST, 0) & ", " & IIf(IsNumeric(A11_GST), A11_GST, 0) & ", " & IIf(IsNumeric(A12_GST), A12_GST, 0) & ", " & _
                IIf(IsNumeric(A13_GST), A13_GST, 0) & ", " & IIf(IsNumeric(A14_GST), A14_GST, 0) & ", " & IIf(IsNumeric(A15_GST), A15_GST, 0) & ", " & _
                IIf(IsNumeric(A16_GST), A16_GST, 0) & ", " & IIf(IsNumeric(A17_GST), A17_GST, 0) & ", " & IIf(IsNumeric(A18_GST), A18_GST, 0) & ", " & _
                IIf(IsNumeric(A19_GST), A19_GST, 0) & ", " & IIf(IsNumeric(A20_GST), A20_GST, 0) & ", " & IIf(IsNumeric(A21_GST), A21_GST, 0) & ", " & _
                IIf(IsNumeric(A22_GST), A22_GST, 0) & ", " & IIf(IsNumeric(A23_GST), A23_GST, 0) & ", " & IIf(IsNumeric(A24_GST), A24_GST, 0) & ", " & _
                IIf(IsNumeric(A25_GST), A25_GST, 0) & ", " & IIf(IsNumeric(A26_GST), A26_GST, 0) & ", " & IIf(IsNumeric(A27_GST), A27_GST, 0) & ", " & _
                IIf(IsNumeric(A28_GST), A28_GST, 0) & ", 0, " & IIf(IsNumeric(A30_GST), A30_GST, 0) & ", " & IIf(IsNumeric(A32_GST), A32_GST, 0) & ", " & _
                IIf(IsNumeric(A33_GST), A33_GST, 0) & ""
        MediConnCISWS.Execute strsql
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub saveActualOthers()
Const METHOD_NAME = "saveActualOthers"
On Error GoTo debugs
Dim strsql As String
        strsql = ""
        tmpCASnumber = Trim(lstInvoice.ListItems(AA).SubItems(1))
        tmpGPFee = Format(lstInvoice.ListItems(AA).SubItems(83), "#,##0.00")
        tmpGPLTMed = Format(lstInvoice.ListItems(AA).SubItems(84), "#,##0.00")
        tmpGPInj = Format(lstInvoice.ListItems(AA).SubItems(85), "#,##0.00")
        tmpGPMed = Format(lstInvoice.ListItems(AA).SubItems(86), "#,##0.00")
        tmpSPFee = Format(lstInvoice.ListItems(AA).SubItems(87), "#,##0.00")
        tmpSPLTMed = Format(lstInvoice.ListItems(AA).SubItems(88), "#,##0.00")
        tmpSPInj = Format(lstInvoice.ListItems(AA).SubItems(89), "#,##0.00")
        tmpSPMed = Format(lstInvoice.ListItems(AA).SubItems(90), "#,##0.00")
    strsql = "Exec INVAdj_saveActualOthers '" & tmpCASnumber & "'," & tmpGPFee & "," & _
            tmpGPLTMed & "," & tmpGPInj & "," & tmpGPMed & "," & tmpSPFee & "," & _
            tmpSPLTMed & "," & tmpSPInj & "," & tmpSPMed & ", " & IIf(IsNumeric(B5_GST), B5_GST, 0) & ", " & _
            IIf(IsNumeric(B6_GST), B6_GST, 0) & ", " & IIf(IsNumeric(B7_GST), B7_GST, 0) & ", " & IIf(IsNumeric(B8_GST), B8_GST, 0) & ""
    MediConnCISWS.Execute strsql
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub SaveINVCLNContract()
Const METHOD_NAME = "SaveINVCLNContract"
On Error GoTo debugs
Dim CISSql As String
        CISSql = ""
        tmpCASnumber = Trim(lstInvoice.ListItems(AA).SubItems(1))
        tmpCLMGPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(211), "#,##0.00")
        tmpCLMGPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(212), "#,##0.00")
        tmpCLMGPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(213), "#,##0.00")
        tmpCLMSPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(214), "#,##0.00")
        tmpCLMSPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(215), "#,##0.00")
        tmpCLMSPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(216), "#,##0.00")
        tmpCLMXRay = Format(lstInvoice.ListItems(AA).SubItems(217), "#,##0.00")
        tmpCLMUltraSound = Format(lstInvoice.ListItems(AA).SubItems(218), "#,##0.00")
        tmpCLMECG = Format(lstInvoice.ListItems(AA).SubItems(219), "#,##0.00")
        tmpCLMDressing = Format(lstInvoice.ListItems(AA).SubItems(220), "#,##0.00")
        tmpCLMNebulizer = Format(lstInvoice.ListItems(AA).SubItems(221), "#,##0.00")
        tmpCLMID = Format(lstInvoice.ListItems(AA).SubItems(222), "#,##0.00")
        tmpCLMTSSTO = Format(lstInvoice.ListItems(AA).SubItems(221), "#,##0.00")
        tmpCLMMinorSurgery = Format(lstInvoice.ListItems(AA).SubItems(224), "#,##0.00")
        tmpCLMPOthers = Format(lstInvoice.ListItems(AA).SubItems(225), "#,##0.00")
        tmpCLMBloodDrawing = Format(lstInvoice.ListItems(AA).SubItems(226), "#,##0.00")
        tmpCLMBloodGlucose = Format(lstInvoice.ListItems(AA).SubItems(227), "#,##0.00")
        tmpCLMUrine = Format(lstInvoice.ListItems(AA).SubItems(228), "#,##0.00")
        tmpCLMPapSmear = Format(lstInvoice.ListItems(AA).SubItems(229), "#,##0.00")
        tmpCLMLOthers = Format(lstInvoice.ListItems(AA).SubItems(230), "#,##0.00")
        tmpCLMDTPPolio = Format(lstInvoice.ListItems(AA).SubItems(231), "#,##0.00")
        tmpCLMDTPPolioHIB = Format(lstInvoice.ListItems(AA).SubItems(232), "#,##0.00")
        tmpCLMHepatitisB = Format(lstInvoice.ListItems(AA).SubItems(233), "#,##0.00")
        tmpCLMMMR = Format(lstInvoice.ListItems(AA).SubItems(234), "#,##0.00")
        tmpCLMIOthers = Format(lstInvoice.ListItems(AA).SubItems(235), "#,##0.00")
        tmpCLMDental = Format(lstInvoice.ListItems(AA).SubItems(236), "#,##0.00")
        tmpCLMMaternity = Format(lstInvoice.ListItems(AA).SubItems(237), "#,##0.00")
        tmpCLMPRVCheckup = Format(lstInvoice.ListItems(AA).SubItems(238), "#,##0.00")
        tmpCLMVision = Format(lstInvoice.ListItems(AA).SubItems(239), "#,##0.00")
        tmpCLMTyphiad = Format(lstInvoice.ListItems(AA).SubItems(240), "#,##0.00")
        tmpCLMOtherNoBalDed = Format(lstInvoice.ListItems(AA).SubItems(241), "#,##0.00")
        
        CISSql = " Exec INVAdj_SaveINVCLNContract '" & tmpCASnumber & "'," & tmpCLMGPFeeBasicMed & "," & _
                 tmpCLMGPFeeMedInj & "," & tmpCLMGPFeeLTMed & "," & tmpCLMSPFeeBasicMed & "," & _
                 tmpCLMSPFeeLTMed & "," & tmpCLMSPFeeMedInj & "," & tmpCLMXRay & "," & _
                 tmpCLMUltraSound & "," & tmpCLMECG & "," & tmpCLMDressing & "," & _
                 tmpCLMNebulizer & "," & tmpCLMID & "," & tmpCLMTSSTO & "," & tmpCLMMinorSurgery & "," & _
                 tmpCLMPOthers & "," & tmpCLMBloodDrawing & "," & tmpCLMBloodGlucose & "," & _
                 tmpCLMUrine & "," & tmpCLMPapSmear & "," & tmpCLMLOthers & "," & _
                 tmpCLMDTPPolio & "," & tmpCLMDTPPolioHIB & "," & tmpCLMHepatitisB & "," & _
                 tmpCLMMMR & "," & tmpCLMIOthers & "," & tmpCLMDental & "," & tmpCLMMaternity & "," & _
                 tmpCLMPRVCheckup & "," & tmpCLMVision & ", " & tmpCLMTyphiad & ", " & tmpCLMOtherNoBalDed & ""
        MediConnCISWS.Execute CISSql

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub SaveINVActual()
Const METHOD_NAME = "SaveINVActual"
On Error GoTo debugs

Dim CISSql As String
        
        CISSql = ""
        tmpCLMNo = Trim(lstInvoice.ListItems(AA).SubItems(1))
        tmpCLMGPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(168), "#,##0.00")
        tmpCLMGPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(169), "#,##0.00")
        tmpCLMGPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(170), "#,##0.00")
        tmpCLMSPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(171), "#,##0.00")
        tmpCLMSPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(172), "#,##0.00")
        tmpCLMSPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(173), "#,##0.00")
        tmpCLMXRay = Format(lstInvoice.ListItems(AA).SubItems(174), "#,##0.00")
        tmpCLMUltraSound = Format(lstInvoice.ListItems(AA).SubItems(175), "#,##0.00")
        tmpCLMECG = Format(lstInvoice.ListItems(AA).SubItems(176), "#,##0.00")
        tmpCLMDressing = Format(lstInvoice.ListItems(AA).SubItems(177), "#,##0.00")
        tmpCLMNebulizer = Format(lstInvoice.ListItems(AA).SubItems(178), "#,##0.00")
        tmpCLMID = Format(lstInvoice.ListItems(AA).SubItems(179), "#,##0.00")
        tmpCLMTSSTO = Format(lstInvoice.ListItems(AA).SubItems(180), "#,##0.00")
        tmpCLMMinorSurgery = Format(lstInvoice.ListItems(AA).SubItems(181), "#,##0.00")
        tmpCLMPOthers = Format(lstInvoice.ListItems(AA).SubItems(182), "#,##0.00")
        tmpCLMBloodDrawing = Format(lstInvoice.ListItems(AA).SubItems(183), "#,##0.00")
        tmpCLMBloodGlucose = Format(lstInvoice.ListItems(AA).SubItems(184), "#,##0.00")
        tmpCLMUrine = Format(lstInvoice.ListItems(AA).SubItems(185), "#,##0.00")
        tmpCLMPapSmear = Format(lstInvoice.ListItems(AA).SubItems(186), "#,##0.00")
        tmpCLMLOthers = Format(lstInvoice.ListItems(AA).SubItems(187), "#,##0.00")
        tmpCLMDTPPolio = Format(lstInvoice.ListItems(AA).SubItems(188), "#,##0.00")
        tmpCLMDTPPolioHIB = Format(lstInvoice.ListItems(AA).SubItems(189), "#,##0.00")
        tmpCLMHepatitisB = Format(lstInvoice.ListItems(AA).SubItems(190), "#,##0.00")
        tmpCLMMMR = Format(lstInvoice.ListItems(AA).SubItems(191), "#,##0.00")
        tmpCLMIOthers = Format(lstInvoice.ListItems(AA).SubItems(192), "#,##0.00")
        tmpCLMDental = Format(lstInvoice.ListItems(AA).SubItems(193), "#,##0.00")
        tmpCLMMaternity = Format(lstInvoice.ListItems(AA).SubItems(194), "#,##0.00")
        tmpCLMPRVCheckup = Format(lstInvoice.ListItems(AA).SubItems(195), "#,##0.00")
        tmpCLMVision = Format(lstInvoice.ListItems(AA).SubItems(196), "#,##0.00")
        tmpCLMTyphiad = Format(lstInvoice.ListItems(AA).SubItems(197), "#,##0.00")
        tmpCLMOtherNoBalDed = Format(lstInvoice.ListItems(AA).SubItems(198), "#,##0.00")
        
        CISSql = "Exec INVAdj_SaveINVActual '" & tmpCLMNo & "'," & tmpCLMGPFeeBasicMed & "," & tmpCLMGPFeeMedInj & "," & _
                tmpCLMGPFeeLTMed & "," & tmpCLMSPFeeBasicMed & "," & tmpCLMSPFeeLTMed & "," & tmpCLMSPFeeMedInj & "," & _
                tmpCLMXRay & "," & tmpCLMUltraSound & "," & tmpCLMECG & "," & tmpCLMDressing & "," & _
                tmpCLMNebulizer & "," & tmpCLMID & "," & tmpCLMTSSTO & "," & tmpCLMMinorSurgery & "," & _
                tmpCLMPOthers & "," & tmpCLMBloodDrawing & "," & tmpCLMBloodGlucose & "," & tmpCLMUrine & "," & _
                tmpCLMPapSmear & "," & tmpCLMLOthers & "," & tmpCLMDTPPolio & "," & tmpCLMDTPPolioHIB & "," & _
                tmpCLMHepatitisB & "," & tmpCLMMMR & "," & tmpCLMIOthers & "," & tmpCLMDental & "," & _
                tmpCLMMaternity & "," & tmpCLMPRVCheckup & "," & tmpCLMVision & ", " & tmpCLMTyphiad & ", " & tmpCLMOtherNoBalDed & ""
        MediConnCISWS.Execute CISSql
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub SaveINVMBMApp()
Const METHOD_NAME = "SaveINVMBMApp"
On Error GoTo debugs
Dim CISSql As String
        CISSql = ""
        tmpCASnumber = Trim(lstInvoice.ListItems(AA).SubItems(1))
        tmpCLMGPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(254), "#,##0.00")
        tmpCLMGPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(255), "#,##0.00")
        tmpCLMGPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(256), "#,##0.00")
        tmpCLMSPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(257), "#,##0.00")
        tmpCLMSPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(258), "#,##0.00")
        tmpCLMSPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(259), "#,##0.00")
        tmpCLMXRay = Format(lstInvoice.ListItems(AA).SubItems(260), "#,##0.00")
        tmpCLMUltraSound = Format(lstInvoice.ListItems(AA).SubItems(261), "#,##0.00")
        tmpCLMECG = Format(lstInvoice.ListItems(AA).SubItems(262), "#,##0.00")
        tmpCLMDressing = Format(lstInvoice.ListItems(AA).SubItems(263), "#,##0.00")
        tmpCLMNebulizer = Format(lstInvoice.ListItems(AA).SubItems(264), "#,##0.00")
        tmpCLMID = Format(lstInvoice.ListItems(AA).SubItems(265), "#,##0.00")
        tmpCLMTSSTO = Format(lstInvoice.ListItems(AA).SubItems(266), "#,##0.00")
        tmpCLMMinorSurgery = Format(lstInvoice.ListItems(AA).SubItems(267), "#,##0.00")
        tmpCLMPOthers = Format(lstInvoice.ListItems(AA).SubItems(268), "#,##0.00")
        tmpCLMBloodDrawing = Format(lstInvoice.ListItems(AA).SubItems(269), "#,##0.00")
        tmpCLMBloodGlucose = Format(lstInvoice.ListItems(AA).SubItems(270), "#,##0.00")
        tmpCLMUrine = Format(lstInvoice.ListItems(AA).SubItems(271), "#,##0.00")
        tmpCLMPapSmear = Format(lstInvoice.ListItems(AA).SubItems(272), "#,##0.00")
        tmpCLMLOthers = Format(lstInvoice.ListItems(AA).SubItems(273), "#,##0.00")
        tmpCLMDTPPolio = Format(lstInvoice.ListItems(AA).SubItems(274), "#,##0.00")
        tmpCLMDTPPolioHIB = Format(lstInvoice.ListItems(AA).SubItems(275), "#,##0.00")
        tmpCLMHepatitisB = Format(lstInvoice.ListItems(AA).SubItems(276), "#,##0.00")
        tmpCLMMMR = Format(lstInvoice.ListItems(AA).SubItems(277), "#,##0.00")
        tmpCLMIOthers = Format(lstInvoice.ListItems(AA).SubItems(278), "#,##0.00")
        tmpCLMDental = Format(lstInvoice.ListItems(AA).SubItems(279), "#,##0.00")
        tmpCLMMaternity = Format(lstInvoice.ListItems(AA).SubItems(280), "#,##0.00")
        tmpCLMPRVCheckup = Format(lstInvoice.ListItems(AA).SubItems(281), "#,##0.00")
        tmpCLMVision = Format(lstInvoice.ListItems(AA).SubItems(282), "#,##0.00")
        tmpCLMTyphiad = Format(lstInvoice.ListItems(AA).SubItems(283), "#,##0.00")
        tmpCLMOtherNoBalDed = Format(lstInvoice.ListItems(AA).SubItems(284), "#,##0.00")
        
        CISSql = " Exec INVAdj_SaveINVMBMApp '" & tmpCASnumber & "'," & tmpCLMGPFeeBasicMed & "," & _
                tmpCLMGPFeeMedInj & "," & tmpCLMGPFeeLTMed & "," & tmpCLMSPFeeBasicMed & "," & tmpCLMSPFeeLTMed & "," & _
                tmpCLMSPFeeMedInj & "," & tmpCLMXRay & "," & tmpCLMUltraSound & "," & tmpCLMECG & "," & tmpCLMDressing & "," & _
                tmpCLMNebulizer & "," & tmpCLMID & "," & tmpCLMTSSTO & "," & tmpCLMMinorSurgery & "," & tmpCLMPOthers & "," & _
                tmpCLMBloodDrawing & "," & tmpCLMBloodGlucose & "," & tmpCLMUrine & "," & tmpCLMPapSmear & "," & tmpCLMLOthers & "," & _
                tmpCLMDTPPolio & "," & tmpCLMDTPPolioHIB & "," & tmpCLMHepatitisB & "," & tmpCLMMMR & "," & tmpCLMIOthers & "," & _
                tmpCLMDental & "," & tmpCLMMaternity & "," & tmpCLMPRVCheckup & "," & tmpCLMVision & ", " & tmpCLMTyphiad & ", " & tmpCLMOtherNoBalDed & ""
        MediConnCISWS.Execute CISSql
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub SaveINVMBMLmt()
Const METHOD_NAME = "SaveINVMBMLmt"
On Error GoTo debugs
Dim CISSql As String
        CISSql = ""
        tmpCASnumber = Trim(lstInvoice.ListItems(AA).SubItems(1))
        tmpCLMGPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(297), "#,##0.00")
        tmpCLMGPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(298), "#,##0.00")
        tmpCLMGPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(299), "#,##0.00")
        tmpCLMSPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(300), "#,##0.00")
        tmpCLMSPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(301), "#,##0.00")
        tmpCLMSPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(302), "#,##0.00")
        tmpCLMXRay = Format(lstInvoice.ListItems(AA).SubItems(303), "#,##0.00")
        tmpCLMUltraSound = Format(lstInvoice.ListItems(AA).SubItems(304), "#,##0.00")
        tmpCLMECG = Format(lstInvoice.ListItems(AA).SubItems(305), "#,##0.00")
        tmpCLMDressing = Format(lstInvoice.ListItems(AA).SubItems(306), "#,##0.00")
        tmpCLMNebulizer = Format(lstInvoice.ListItems(AA).SubItems(307), "#,##0.00")
        tmpCLMID = Format(lstInvoice.ListItems(AA).SubItems(308), "#,##0.00")
        tmpCLMTSSTO = Format(lstInvoice.ListItems(AA).SubItems(309), "#,##0.00")
        tmpCLMMinorSurgery = Format(lstInvoice.ListItems(AA).SubItems(310), "#,##0.00")
        tmpCLMPOthers = Format(lstInvoice.ListItems(AA).SubItems(311), "#,##0.00")
        tmpCLMBloodDrawing = Format(lstInvoice.ListItems(AA).SubItems(312), "#,##0.00")
        tmpCLMBloodGlucose = Format(lstInvoice.ListItems(AA).SubItems(313), "#,##0.00")
        tmpCLMUrine = Format(lstInvoice.ListItems(AA).SubItems(314), "#,##0.00")
        tmpCLMPapSmear = Format(lstInvoice.ListItems(AA).SubItems(315), "#,##0.00")
        tmpCLMLOthers = Format(lstInvoice.ListItems(AA).SubItems(316), "#,##0.00")
        tmpCLMDTPPolio = Format(lstInvoice.ListItems(AA).SubItems(317), "#,##0.00")
        tmpCLMDTPPolioHIB = Format(lstInvoice.ListItems(AA).SubItems(318), "#,##0.00")
        tmpCLMHepatitisB = Format(lstInvoice.ListItems(AA).SubItems(319), "#,##0.00")
        tmpCLMMMR = Format(lstInvoice.ListItems(AA).SubItems(320), "#,##0.00")
        tmpCLMIOthers = Format(lstInvoice.ListItems(AA).SubItems(321), "#,##0.00")
        tmpCLMDental = Format(lstInvoice.ListItems(AA).SubItems(322), "#,##0.00")
        tmpCLMMaternity = Format(lstInvoice.ListItems(AA).SubItems(323), "#,##0.00")
        tmpCLMPRVCheckup = Format(lstInvoice.ListItems(AA).SubItems(324), "#,##0.00")
        tmpCLMVision = Format(lstInvoice.ListItems(AA).SubItems(325), "#,##0.00")
        tmpCLMTyphiad = Format(lstInvoice.ListItems(AA).SubItems(326), "#,##0.00")
        tmpCLMOtherNoBalDed = Format(lstInvoice.ListItems(AA).SubItems(327), "#,##0.00")

        CISSql = "Exec INVAdj_SaveINVMBMLmt '" & tmpCASnumber & "'," & tmpCLMGPFeeBasicMed & "," & tmpCLMGPFeeMedInj & "," & _
                tmpCLMGPFeeLTMed & "," & tmpCLMSPFeeBasicMed & "," & tmpCLMSPFeeLTMed & "," & tmpCLMSPFeeMedInj & "," & _
                tmpCLMXRay & "," & tmpCLMUltraSound & "," & tmpCLMECG & "," & tmpCLMDressing & "," & tmpCLMNebulizer & "," & _
                tmpCLMID & "," & tmpCLMTSSTO & "," & tmpCLMMinorSurgery & "," & tmpCLMPOthers & "," & tmpCLMBloodDrawing & "," & _
                tmpCLMBloodGlucose & "," & tmpCLMUrine & "," & tmpCLMPapSmear & "," & tmpCLMLOthers & "," & tmpCLMDTPPolio & "," & _
                tmpCLMDTPPolioHIB & "," & tmpCLMHepatitisB & "," & tmpCLMMMR & "," & tmpCLMIOthers & "," & tmpCLMDental & "," & _
                tmpCLMMaternity & "," & tmpCLMPRVCheckup & "," & tmpCLMVision & ", " & tmpCLMTyphiad & ", " & tmpCLMOtherNoBalDed & ""
        MediConnCISWS.Execute CISSql
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub SaveINVMBMExcess()
Const METHOD_NAME = "SaveINVMBMExcess"
On Error GoTo debugs
Dim CISSql As String
    
        tmpCASnumber = Trim(lstInvoice.ListItems(AA).SubItems(1))
        tmpCLMGPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(339), "#,##0.00")
        tmpCLMGPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(340), "#,##0.00")
        tmpCLMGPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(341), "#,##0.00")
        tmpCLMSPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(342), "#,##0.00")
        tmpCLMSPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(343), "#,##0.00")
        tmpCLMSPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(344), "#,##0.00")
        tmpCLMXRay = Format(lstInvoice.ListItems(AA).SubItems(345), "#,##0.00")
        tmpCLMUltraSound = Format(lstInvoice.ListItems(AA).SubItems(346), "#,##0.00")
        tmpCLMECG = Format(lstInvoice.ListItems(AA).SubItems(347), "#,##0.00")
        tmpCLMDressing = Format(lstInvoice.ListItems(AA).SubItems(348), "#,##0.00")
        tmpCLMNebulizer = Format(lstInvoice.ListItems(AA).SubItems(349), "#,##0.00")
        tmpCLMID = Format(lstInvoice.ListItems(AA).SubItems(350), "#,##0.00")
        tmpCLMTSSTO = Format(lstInvoice.ListItems(AA).SubItems(351), "#,##0.00")
        tmpCLMMinorSurgery = Format(lstInvoice.ListItems(AA).SubItems(352), "#,##0.00")
        tmpCLMPOthers = Format(lstInvoice.ListItems(AA).SubItems(353), "#,##0.00")
        tmpCLMBloodDrawing = Format(lstInvoice.ListItems(AA).SubItems(354), "#,##0.00")
        tmpCLMBloodGlucose = Format(lstInvoice.ListItems(AA).SubItems(355), "#,##0.00")
        tmpCLMUrine = Format(lstInvoice.ListItems(AA).SubItems(356), "#,##0.00")
        tmpCLMPapSmear = Format(lstInvoice.ListItems(AA).SubItems(357), "#,##0.00")
        tmpCLMLOthers = Format(lstInvoice.ListItems(AA).SubItems(358), "#,##0.00")
        tmpCLMDTPPolio = Format(lstInvoice.ListItems(AA).SubItems(359), "#,##0.00")
        tmpCLMDTPPolioHIB = Format(lstInvoice.ListItems(AA).SubItems(360), "#,##0.00")
        tmpCLMHepatitisB = Format(lstInvoice.ListItems(AA).SubItems(361), "#,##0.00")
        tmpCLMMMR = Format(lstInvoice.ListItems(AA).SubItems(362), "#,##0.00")
        tmpCLMIOthers = Format(lstInvoice.ListItems(AA).SubItems(363), "#,##0.00")
        tmpCLMDental = Format(lstInvoice.ListItems(AA).SubItems(364), "#,##0.00")
        tmpCLMMaternity = Format(lstInvoice.ListItems(AA).SubItems(365), "#,##0.00")
        tmpCLMPRVCheckup = Format(lstInvoice.ListItems(AA).SubItems(366), "#,##0.00")
        tmpCLMVision = Format(lstInvoice.ListItems(AA).SubItems(367), "#,##0.00")
        tmpCLMTyphiad = Format(lstInvoice.ListItems(AA).SubItems(368), "#,##0.00")
        tmpCLMOtherNoBalDed = Format(lstInvoice.ListItems(AA).SubItems(369), "#,##0.00")
        CISSql = "Exec INVAdj_SaveINVMBMExcess '" & tmpCASnumber & "'," & tmpCLMGPFeeBasicMed & "," & tmpCLMGPFeeMedInj & "," & _
                tmpCLMGPFeeLTMed & "," & tmpCLMSPFeeBasicMed & "," & tmpCLMSPFeeLTMed & "," & tmpCLMSPFeeMedInj & "," & _
                tmpCLMXRay & "," & tmpCLMUltraSound & "," & tmpCLMECG & "," & tmpCLMDressing & "," & tmpCLMNebulizer & "," & _
                tmpCLMID & "," & tmpCLMTSSTO & "," & tmpCLMMinorSurgery & "," & tmpCLMPOthers & "," & tmpCLMBloodDrawing & "," & _
                tmpCLMBloodGlucose & "," & tmpCLMUrine & "," & tmpCLMPapSmear & "," & tmpCLMLOthers & "," & tmpCLMDTPPolio & "," & _
                tmpCLMDTPPolioHIB & "," & tmpCLMHepatitisB & "," & tmpCLMMMR & "," & tmpCLMIOthers & "," & tmpCLMDental & "," & _
                tmpCLMMaternity & "," & tmpCLMPRVCheckup & "," & tmpCLMVision & ", " & tmpCLMTyphiad & ", " & tmpCLMOtherNoBalDed & ""
        MediConnCISWS.Execute CISSql
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub SaveINVCLNExcess()
Const METHOD_NAME = "SaveINVCLNExcess"
On Error GoTo debugs
Dim CISSql As String
        CISSql = ""
        tmpCASnumber = Trim(lstInvoice.ListItems(AA).SubItems(1))
        tmpCLMGPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(382), "#,##0.00")
        tmpCLMGPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(383), "#,##0.00")
        tmpCLMGPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(384), "#,##0.00")
        tmpCLMSPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(385), "#,##0.00")
        tmpCLMSPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(386), "#,##0.00")
        tmpCLMSPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(387), "#,##0.00")
        tmpCLMXRay = Format(lstInvoice.ListItems(AA).SubItems(388), "#,##0.00")
        tmpCLMUltraSound = Format(lstInvoice.ListItems(AA).SubItems(389), "#,##0.00")
        tmpCLMECG = Format(lstInvoice.ListItems(AA).SubItems(390), "#,##0.00")
        tmpCLMDressing = Format(lstInvoice.ListItems(AA).SubItems(391), "#,##0.00")
        tmpCLMNebulizer = Format(lstInvoice.ListItems(AA).SubItems(392), "#,##0.00")
        tmpCLMID = Format(lstInvoice.ListItems(AA).SubItems(393), "#,##0.00")
        tmpCLMTSSTO = Format(lstInvoice.ListItems(AA).SubItems(394), "#,##0.00")
        tmpCLMMinorSurgery = Format(lstInvoice.ListItems(AA).SubItems(395), "#,##0.00")
        tmpCLMPOthers = Format(lstInvoice.ListItems(AA).SubItems(396), "#,##0.00")
        tmpCLMBloodDrawing = Format(lstInvoice.ListItems(AA).SubItems(397), "#,##0.00")
        tmpCLMBloodGlucose = Format(lstInvoice.ListItems(AA).SubItems(398), "#,##0.00")
        tmpCLMUrine = Format(lstInvoice.ListItems(AA).SubItems(399), "#,##0.00")
        tmpCLMPapSmear = Format(lstInvoice.ListItems(AA).SubItems(400), "#,##0.00")
        tmpCLMLOthers = Format(lstInvoice.ListItems(AA).SubItems(401), "#,##0.00")
        tmpCLMDTPPolio = Format(lstInvoice.ListItems(AA).SubItems(402), "#,##0.00")
        tmpCLMDTPPolioHIB = Format(lstInvoice.ListItems(AA).SubItems(403), "#,##0.00")
        tmpCLMHepatitisB = Format(lstInvoice.ListItems(AA).SubItems(404), "#,##0.00")
        tmpCLMMMR = Format(lstInvoice.ListItems(AA).SubItems(405), "#,##0.00")
        tmpCLMIOthers = Format(lstInvoice.ListItems(AA).SubItems(406), "#,##0.00")
        tmpCLMDental = Format(lstInvoice.ListItems(AA).SubItems(407), "#,##0.00")
        tmpCLMMaternity = Format(lstInvoice.ListItems(AA).SubItems(408), "#,##0.00")
        tmpCLMPRVCheckup = Format(lstInvoice.ListItems(AA).SubItems(409), "#,##0.00")
        tmpCLMVision = Format(lstInvoice.ListItems(AA).SubItems(410), "#,##0.00")
        tmpCLMTyphiad = Format(lstInvoice.ListItems(AA).SubItems(411), "#,##0.00")
        tmpCLMOtherNoBalDed = Format(lstInvoice.ListItems(AA).SubItems(412), "#,##0.00")
        
        CISSql = "Exec INVAdj_SaveINVCLNExcess '" & tmpCASnumber & "'," & tmpCLMGPFeeBasicMed & "," & tmpCLMGPFeeMedInj & "," & _
                tmpCLMGPFeeLTMed & "," & tmpCLMSPFeeBasicMed & "," & tmpCLMSPFeeLTMed & "," & tmpCLMSPFeeMedInj & "," & _
                tmpCLMXRay & "," & tmpCLMUltraSound & "," & tmpCLMECG & "," & tmpCLMDressing & "," & tmpCLMNebulizer & "," & _
                tmpCLMID & "," & tmpCLMTSSTO & "," & tmpCLMMinorSurgery & "," & tmpCLMPOthers & "," & tmpCLMBloodDrawing & "," & _
                tmpCLMBloodGlucose & "," & tmpCLMUrine & "," & tmpCLMPapSmear & "," & tmpCLMLOthers & "," & tmpCLMDTPPolio & "," & _
                tmpCLMDTPPolioHIB & "," & tmpCLMHepatitisB & "," & tmpCLMMMR & "," & tmpCLMIOthers & "," & tmpCLMDental & "," & _
                tmpCLMMaternity & "," & tmpCLMPRVCheckup & "," & tmpCLMVision & ", " & tmpCLMTyphiad & ", " & tmpCLMOtherNoBalDed & ""
        MediConnCISWS.Execute CISSql
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub


Private Sub SaveINVCLNApp()
Const METHOD_NAME = "SaveINVCLNApp"
On Error GoTo debugs
Dim CISSql As String
        CISSql = ""
        tmpCASnumber = Trim(lstInvoice.ListItems(AA).SubItems(1))
        tmpCLMGPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(426), "#,##0.00")
        tmpCLMGPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(427), "#,##0.00")
        tmpCLMGPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(428), "#,##0.00")
        tmpCLMSPFeeBasicMed = Format(lstInvoice.ListItems(AA).SubItems(429), "#,##0.00")
        tmpCLMSPFeeLTMed = Format(lstInvoice.ListItems(AA).SubItems(430), "#,##0.00")
        tmpCLMSPFeeMedInj = Format(lstInvoice.ListItems(AA).SubItems(431), "#,##0.00")
        tmpCLMXRay = Format(lstInvoice.ListItems(AA).SubItems(432), "#,##0.00")
        tmpCLMUltraSound = Format(lstInvoice.ListItems(AA).SubItems(433), "#,##0.00")
        tmpCLMECG = Format(lstInvoice.ListItems(AA).SubItems(434), "#,##0.00")
        tmpCLMDressing = Format(lstInvoice.ListItems(AA).SubItems(435), "#,##0.00")
        tmpCLMNebulizer = Format(lstInvoice.ListItems(AA).SubItems(436), "#,##0.00")
        tmpCLMID = Format(lstInvoice.ListItems(AA).SubItems(437), "#,##0.00")
        tmpCLMTSSTO = Format(lstInvoice.ListItems(AA).SubItems(438), "#,##0.00")
        tmpCLMMinorSurgery = Format(lstInvoice.ListItems(AA).SubItems(439), "#,##0.00")
        tmpCLMPOthers = Format(lstInvoice.ListItems(AA).SubItems(440), "#,##0.00")
        tmpCLMBloodDrawing = Format(lstInvoice.ListItems(AA).SubItems(441), "#,##0.00")
        tmpCLMBloodGlucose = Format(lstInvoice.ListItems(AA).SubItems(442), "#,##0.00")
        tmpCLMUrine = Format(lstInvoice.ListItems(AA).SubItems(443), "#,##0.00")
        tmpCLMPapSmear = Format(lstInvoice.ListItems(AA).SubItems(444), "#,##0.00")
        tmpCLMLOthers = Format(lstInvoice.ListItems(AA).SubItems(445), "#,##0.00")
        tmpCLMDTPPolio = Format(lstInvoice.ListItems(AA).SubItems(446), "#,##0.00")
        tmpCLMDTPPolioHIB = Format(lstInvoice.ListItems(AA).SubItems(447), "#,##0.00")
        tmpCLMHepatitisB = Format(lstInvoice.ListItems(AA).SubItems(448), "#,##0.00")
        tmpCLMMMR = Format(lstInvoice.ListItems(AA).SubItems(449), "#,##0.00")
        tmpCLMIOthers = Format(lstInvoice.ListItems(AA).SubItems(450), "#,##0.00")
        tmpCLMDental = Format(lstInvoice.ListItems(AA).SubItems(451), "#,##0.00")
        tmpCLMMaternity = Format(lstInvoice.ListItems(AA).SubItems(452), "#,##0.00")
        tmpCLMPRVCheckup = Format(lstInvoice.ListItems(AA).SubItems(453), "#,##0.00")
        tmpCLMVision = Format(lstInvoice.ListItems(AA).SubItems(454), "#,##0.00")
        tmpCLMTyphiad = Format(lstInvoice.ListItems(AA).SubItems(455), "#,##0.00")
        tmpCLMOtherNoBalDed = Format(lstInvoice.ListItems(AA).SubItems(456), "#,##0.00")

        CISSql = "Exec INVAdj_SaveINVCLNApp '" & tmpCASnumber & "'," & tmpCLMGPFeeBasicMed & "," & _
                tmpCLMGPFeeMedInj & "," & tmpCLMGPFeeLTMed & "," & tmpCLMSPFeeBasicMed & "," & _
                tmpCLMSPFeeLTMed & "," & tmpCLMSPFeeMedInj & "," & tmpCLMXRay & "," & _
                tmpCLMUltraSound & "," & tmpCLMECG & "," & tmpCLMDressing & "," & _
                tmpCLMNebulizer & "," & tmpCLMID & "," & tmpCLMTSSTO & "," & _
                tmpCLMMinorSurgery & "," & tmpCLMPOthers & "," & tmpCLMBloodDrawing & "," & _
                tmpCLMBloodGlucose & "," & tmpCLMUrine & "," & tmpCLMPapSmear & "," & _
                tmpCLMLOthers & "," & tmpCLMDTPPolio & "," & tmpCLMDTPPolioHIB & "," & _
                tmpCLMHepatitisB & "," & tmpCLMMMR & "," & tmpCLMIOthers & "," & _
                tmpCLMDental & "," & tmpCLMMaternity & "," & tmpCLMPRVCheckup & "," & _
                tmpCLMVision & ", " & tmpCLMTyphiad & ", " & tmpCLMOtherNoBalDed & ""
        MediConnCISWS.Execute CISSql
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
'Price MURALi
Private Sub SaveCLMMEDCHARGES()
Const METHOD_NAME = "SaveCLMMEDCHARGES"
On Error GoTo debugs
Dim SaveCaseSql As String
Dim tmpMEDCHARGES1 As Double
Dim tmpMEDCHARGES2 As Double
Dim tmpMEDCHARGES3 As Double
Dim tmpMEDCHARGES4 As Double
Dim tmpMEDCHARGES5 As Double
Dim tmpMEDCHARGES6 As Double
Dim tmpMEDCHARGES7 As Double

    SaveCaseSql = ""
        tmpMEDCHARGES1 = lstInvoice.ListItems(AA).SubItems(482)
        tmpMEDCHARGES2 = lstInvoice.ListItems(AA).SubItems(483)
        tmpMEDCHARGES3 = lstInvoice.ListItems(AA).SubItems(484)
        tmpMEDCHARGES4 = lstInvoice.ListItems(AA).SubItems(485)
        tmpMEDCHARGES5 = lstInvoice.ListItems(AA).SubItems(486)
        tmpMEDCHARGES6 = lstInvoice.ListItems(AA).SubItems(487)
        tmpMEDCHARGES7 = lstInvoice.ListItems(AA).SubItems(488)
        SaveCaseSql = ""
        SaveCaseSql = "Update CLMMEDCHARGES set MEDCHARGES1=" & tmpMEDCHARGES1 & ", MEDCHARGES2=" & tmpMEDCHARGES2 & ", MEDCHARGES3=" & tmpMEDCHARGES3 & ", MEDCHARGES4=" & tmpMEDCHARGES4 & ", MEDCHARGES5=" & tmpMEDCHARGES5 & ", MEDCHARGES6=" & tmpMEDCHARGES6 & ", MEDCHARGES7=" & tmpMEDCHARGES7 & " where CASNUMBER='" & tmpCASnumber & "'"
        MediConnCISWS.Execute SaveCaseSql
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub SaveINVClaim()
Const METHOD_NAME = "SaveINVClaim"
On Error GoTo debugs
Dim CISSql As String
Dim tmpNonGPFee As Double
    
        CISSql = ""
        tmpCASnumber = Trim(lstInvoice.ListItems(AA).SubItems(1))
        tmpClaimCloseDate = Format(Date, "yyyy/mm/dd")
        tmpClaimClosedBy = Logon
        tmpClaimLastUpdateDate = Format(Date, "yyyy/mm/dd")
        tmpClaimLastUpdateUser = Logon
        If Trim(strPayTo) = "C" Then
            tmpCLMExcess = Format(lstInvoice.ListItems(AA).SubItems(162), "#,##0.00")
            If FFSIndicator Then
                    Dim sqlstr As String
                    Dim FFSRs As New ADODB.Recordset
                    Dim FFSpercentage As Double
                    FFSpercentage = 0
                    sqlstr = "Select CLNMCOProcfeePerc  from CLN where CLNCode ='" + txtClinicCode + "'"
                    FFSRs.Open sqlstr, medixmasterconn, adOpenDynamic, adLockOptimistic
                    Do Until FFSRs.EOF
                        If IsNumeric(FFSRs!CLNMCOProcfeePerc) Then
                            FFSpercentage = FFSRs!CLNMCOProcfeePerc
                        End If
                        FFSRs.MoveNext
                    Loop
                    
                    If IsNumeric(lblCLNBillAmt) Then
                        tmpCLMMCOProcFee = (FFSpercentage * CDbl(lblCLNBillAmt)) / 100
                    Else
                        tmpCLMMCOProcFee = 0
                    End If
            Else
                tmpCLMMCOProcFee = 0
            End If
            
            If (Len(lblCLNBillAmt) > 0) Then
                tmpCLMPayableByMedix2C = CDbl(lblCLNBillAmt) - CDbl(tmpCLMMCOProcFee)
            End If
            
        ElseIf Trim(strPayTo) = "M" Then
            If lstInvoice.ListItems(AA).SubItems(475) = "" Then
                tmpCLMPayableByMedix2M = Format(lstInvoice.ListItems(AA).SubItems(162) + CDbl(tmpNonGPFee), "#,##0.00")
            Else
                If IsNumeric(lstInvoice.ListItems(AA).SubItems(475)) And lstInvoice.ListItems(AA).SubItems(475) > 0 Then
                    tmpCLMPayableByMedix2M = 0
                Else
                    tmpCLMPayableByMedix2M = Format(lstInvoice.ListItems(AA).SubItems(162) + CDbl(tmpNonGPFee), "#,##0.00")
                End If
            End If
        End If
        
        tmpCLMCLNExcess = Format(lstInvoice.ListItems(AA).SubItems(159), "#,##0.00")
        tmpCLMTotalActual = Format(txtCLNBillAmt.Text, "#,##0.00") 'Format(lstInvoice.ListItems(AA).SubItems(160), "#,##0.00")
        tmpCLMFinalApprove = Format(lstInvoice.ListItems(AA).SubItems(161), "#,##0.00")
        'tmpCLMFinalApprove = Format(lstInvoice.ListItems(AA).SubItems(160), "#,##0.00")
        tmpCLMBalanceLimit = Format(lstInvoice.ListItems(AA).SubItems(163), "#,##0.00")
        tmpCLMPreviousBalance = Format(lstInvoice.ListItems(AA).SubItems(164), "#,##0.00")
        tmpCLMPaidBYMedix2M = Format(lstInvoice.ListItems(AA).SubItems(165), "#,##0.00")
        tmpCLMPaidBYMedix2C = Format(lstInvoice.ListItems(AA).SubItems(166), "#,##0.00")
        tmpCLMAdjAnnLmt = Format(lstInvoice.ListItems(AA).SubItems(113), "#,##0.00")
        tmpClmNonGPFee = Format(lstInvoice.ListItems(AA).SubItems(114), "#,##0.00")
        tmpNonGPFee = Format(IIf(IsNumeric(lstInvoice.ListItems(AA).SubItems(114)), lstInvoice.ListItems(AA).SubItems(114), 0), "#,##0.00")
        tmpCLMCoPayColl = Format(IIf(IsNumeric(lstInvoice.ListItems(AA).SubItems(43)), lstInvoice.ListItems(AA).SubItems(43), 0), "#,##0.00")
        tmpCLMMCOProcStatus = IIf(Len(lstInvoice.ListItems(AA).SubItems(469)), lstInvoice.ListItems(AA).SubItems(469), "")
        tmpCLMProcMth = IIf(Len(lstInvoice.ListItems(AA).SubItems(470)), Mid(lstInvoice.ListItems(AA).SubItems(470), 1, 2), "")
        tmpCLMProcYR = IIf(Len(lstInvoice.ListItems(AA).SubItems(470)), Mid(lstInvoice.ListItems(AA).SubItems(470), 4, 4), "")
        tmpCLMCoPayCLNColl = Format(IIf(IsNumeric(lstInvoice.ListItems(AA).SubItems(471)), lstInvoice.ListItems(AA).SubItems(471), 0), "#,##0.00")
        
        
        If InStr(1, tmpCLMCoPayColl, "-") Then
            tmpCLMCoPayColl = ((tmpCLMCoPayColl) * (-1))
        End If
        If InStr(1, tmpCLMCoPayCLNColl, "-") Then
            tmpCLMCoPayCLNColl = ((tmpCLMCoPayCLNColl) * (-1))
        End If
        
        If Mid(CboPayto, 1, 1) = "M" Then
            If lstInvoice.ListItems(AA).SubItems(475) = "" Then
                If lstInvoice.ListItems(AA).SubItems(161) <> 0 Then
                    tmpCLMBordxAmt = Format(lstInvoice.ListItems(AA).SubItems(161) + CDbl(tmpNonGPFee), "#,##0.00")
                Else
                    tmpCLMBordxAmt = Format(CDbl(tmpCLMPayableByMedix2M) + CDbl(tmpNonGPFee), "#,##0.00")
                End If
            Else
                If IsNumeric(lstInvoice.ListItems(AA).SubItems(475)) And lstInvoice.ListItems(AA).SubItems(475) > 0 Then
                    tmpCLMBordxAmt = lstInvoice.ListItems(AA).SubItems(161)
                Else
                    tmpCLMBordxAmt = Format(lstInvoice.ListItems(AA).SubItems(158) + CDbl(tmpNonGPFee), "#,##0.00")
                End If
            End If
        Else
            If Trim(strGRPMCOStatus) = "F" Then
                If lstInvoice.ListItems(AA).SubItems(161) <> 0 Then
                    tmpCLMBordxAmt = Format(lstInvoice.ListItems(AA).SubItems(161) + CDbl(tmpNonGPFee), "#,##0.00")
                Else
                    tmpCLMBordxAmt = Format(CDbl(tmpCLMPayableByMedix2M) + CDbl(tmpNonGPFee), "#,##0.00")
                End If
            Else
                tmpCLMBordxAmt = Format(lstInvoice.ListItems(AA).SubItems(158) + CDbl(tmpNonGPFee), "#,##0.00")
            End If
        End If
        
        tmpLimitExceeded = IIf(IsNumeric(lstInvoice.ListItems(AA).SubItems(480)), lstInvoice.ListItems(AA).SubItems(480), 0)
        txtExcess = IIf(IsNumeric(txtExcess), txtExcess, 0)
        
        If tmpCLMFinalApprove = 0 Then
            If IsNumeric(lblCLNBillAmt) Then
                tmpCLMFinalApprove = CDbl(lblCLNBillAmt)
            End If
        End If
        
        If tmpCLMBordxAmt = 0 Then
            If IsNumeric(lblCLNBillAmt) Then
                tmpCLMBordxAmt = CDbl(lblCLNBillAmt)
            End If
        End If
        
        CISSql = ""
        CISSql = "Exec INVAdj_SaveINVClaim_2015 '" & tmpCASnumber & "', '" & tmpClaimCloseDate & "', '" & _
                tmpClaimClosedBy & "', '" & tmpClaimLastUpdateDate & "', '" & tmpClaimLastUpdateUser & "'," & _
                tmpCLMPayableByMedix2C & "," & tmpCLMExcess & "," & tmpCLMMCOProcFee & "," & _
                tmpCLMPayableByMedix2M & "," & tmpCLMCLNExcess & "," & tmpCLMTotalActual & "," & _
                tmpCLMFinalApprove & "," & tmpCLMBalanceLimit & "," & tmpCLMPreviousBalance & "," & _
                tmpCLMPaidBYMedix2M & "," & tmpCLMPaidBYMedix2C & "," & tmpCLMAdjAnnLmt & "," & _
                tmpClmNonGPFee & "," & tmpNonGPFee & "," & tmpCLMCoPayColl & ",'" & _
                Trim(tmpCLMMCOProcStatus) & "','" & Trim(tmpCLMProcMth) & "','" & Trim(tmpCLMProcYR) & "'," & _
                tmpCLMCoPayCLNColl & "," & tmpCLMBordxAmt & " , '" & IIf(Len(cboMemoClaim), Mid(cboMemoClaim, 1, 1), "") & "', '" & _
                IIf(Len(cboRefLetter), Mid(cboRefLetter, 1, 1), "") & "' , " & tmpLimitExceeded & ", " & txtExcess & "" & ", " & _
                IIf(Trim(txtBTBG.Text) = "", 0, Trim(txtBTBG.Text)) & "" & ", '" & Trim(txtBTBGRemark.Text) & "'"
        MediConnCISWS.Execute CISSql
        
        Dim LGExc As New ADODB.Recordset
        Dim UpSql As String
        UpSql = "Update CISDB.Dbo.LG set CISDB.Dbo.LG.GLAmt= " & tmpCLMTotalActual & " where CISDB.Dbo.LG.LGNo='" & tmpCASnumber & "'"
        LGExc.Open UpSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        Set LGExc = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Function ExportToExcelError(lvw As MSComctlLib.ListView) As Boolean
Const METHOD_NAME = "ExportToExcelError"
On Error GoTo debugs
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet
Dim objRange As Excel.Range
Dim lngResults As Long
Dim i As Integer
Dim intCounter As Integer
Dim intStartRow As Integer
Dim strArray() As String
Dim intVisibleColumns() As Integer
Dim intColumns As Integer
Dim itm As ListItem
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    On Error Resume Next
    Set objExcel = New Excel.Application

    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False

    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Error.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
            
    intCounter = 0
        
    Set objRange = objWorksheet.Rows(4)
    objRange.Font.Size = 10
    objRange.Font.Bold = True

    If lvw Is Nothing = False Then
        For i = 1 To lvw.ColumnHeaders.Count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
            End If
        Next i
        If lvw.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
        End If
        
    End If
    
'Dimension array to number of listitems
    
    intCounter = 0
    intStartRow = 5
    
    If lvw Is Nothing = False Then
        For Each itm In lvw.ListItems
            intCounter = intCounter + 1
            For i = 1 To intColumns
                If i = 1 Then
                    strArray(intCounter, 1) = itm.Text
                Else
                    strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
                End If
            Next i
        Next itm
    End If

    With objWorksheet
        If lvw Is Nothing = False And intCounter <> 0 Then
            .Range(.Cells(5, 1), _
            .Cells(4 + intCounter, intColumns)) = strArray
        End If
    End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "C") = Date
    End If
        
    objWorksheet.Columns.AutoFit
    objExcel.Interactive = True
    
    ExportToExcelError = True
ExitFunction:
    Screen.MousePointer = vbDefault
    Exit Function
HANDLE_ERROR:
    MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
    Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
    Set objRange = Nothing
    Set objWorksheet = Nothing
    Set objWorkbook = Nothing
    objExcel.Quit
    GoTo ExitFunction
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Function

Public Sub ValidateBSNUser(MBMNumber As String)
Const METHOD_NAME = "ValidateBSNUser"
On Error GoTo debugs
    Dim txtGrpCompany As String
    Dim Balcmd As New ADODB.Command
    Dim BalPrm1 As New ADODB.Parameter
    Dim BalRes As New ADODB.Recordset
        Set Balcmd.ActiveConnection = MediConnCISDB
        Balcmd.CommandText = "sp_BSNMemberVerification"
        Balcmd.CommandType = adCmdStoredProc
        Balcmd.CommandTimeout = 300
        Set BalPrm1 = Balcmd.CreateParameter("MemeberNo", adVarChar, adParamInput, 2000, MBMNumber)
        Balcmd.Parameters.Append BalPrm1
        BalRes.CursorLocation = adUseClient
        BalRes.CursorType = adOpenForwardOnly
        BalRes.CacheSize = 100
        Set BalRes = Balcmd.Execute
        Do While Not BalRes.EOF
            txtGrpCompany = BalRes!Number
            BalRes.MoveNext
        Loop
        BalRes.Close
        Set BalRes = Nothing
        
        If txtGrpCompany <> "0" Then
           BSNUser = True
        Else
           BSNUser = False
        End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub UpdateMBMBalance()
Const METHOD_NAME = "UpdateMBMBalance"
On Error GoTo debugs
Dim MBM As New ADODB.Recordset
Dim strsql As String
Dim FinalBalance  As Double
Dim TotalApprove As Double
Dim tmpAccumAmt As Double
Dim tmpMBMAdj As Double
Dim tmpNewMBMAdj As Double
Dim tmpPay2CLN As Double
Dim tmpMaternity As Double


    strsql = "Select Top 1  MBMNUmber , MBMCLNBalLmt, MBMAccumAmt, MBMAdjAnnLmt from MBMCLNOthers " & _
            " where MBMNUmber ='" & Trim(lstInvoice.ListItems(AA).SubItems(2)) & "'"
    
    If Trim(lstInvoice.ListItems(AA).SubItems(3)) <> "00" Then
        If lstInvoice.ListItems(AA).SubItems(82) = "B" Or lstInvoice.ListItems(AA).SubItems(82) = "D" Then 'SuppLimitType
            strsql = "Select MBMCNUmber , MBMCCLNBalLmt as MBMCLNBalLmt, MBMCAccumAmt as MBMAccumAmt, " & _
            "MBMCAdjAnnLmt as MBMAdjAnnLmt from MBMCoveredPersonsCLN " & _
            " where MBMCNUmber ='" & Trim(lstInvoice.ListItems(AA).SubItems(2)) & "' And MBMCCoverID   ='" & "01" & "'"
        ElseIf lstInvoice.ListItems(AA).SubItems(82) = "C" Then
            strsql = "Select MBMCNUmber , MBMCCLNBalLmt as MBMCLNBalLmt, MBMCAccumAmt as MBMAccumAmt, " & _
            "MBMCAdjAnnLmt as MBMAdjAnnLmt from MBMCoveredPersonsCLN " & _
            " where MBMCNUmber ='" & Trim(lstInvoice.ListItems(AA).SubItems(2)) & "' And MBMCCoverID   ='" & Trim(lstInvoice.ListItems(AA).SubItems(3)) & "'"
        End If
    End If
    
    MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    tmpPay2CLN = IIf(IsNumeric(lstInvoice.ListItems(AA).SubItems(158)), lstInvoice.ListItems(AA).SubItems(158), 0)
    If Mid(CboPayto, 1, 1) = "M" Then
        If lstInvoice.ListItems(AA).SubItems(475) = "" Then
            TotalApprove = IIf(IsNumeric(lstInvoice.ListItems(AA).SubItems(162)), lstInvoice.ListItems(AA).SubItems(162), 0)
            'TotalApprove = IIf(IsNumeric(lstInvoice.ListItems(AA).SubItems(160)), lstInvoice.ListItems(AA).SubItems(160), 0)
        Else
            If IsNumeric(lstInvoice.ListItems(AA).SubItems(475)) And lstInvoice.ListItems(AA).SubItems(475) > 0 Then
                TotalApprove = IIf(IsNumeric(lstInvoice.ListItems(AA).SubItems(160)), lstInvoice.ListItems(AA).SubItems(160), 0)
            Else
                TotalApprove = IIf(IsNumeric(lstInvoice.ListItems(AA).SubItems(162)), lstInvoice.ListItems(AA).SubItems(162), 0)
                'TotalApprove = IIf(IsNumeric(lstInvoice.ListItems(AA).SubItems(160)), lstInvoice.ListItems(AA).SubItems(160), 0)
            End If
        End If
    Else
        TotalApprove = IIf(IsNumeric(lstInvoice.ListItems(AA).SubItems(161)), lstInvoice.ListItems(AA).SubItems(161), 0)
        'TotalApprove = IIf(IsNumeric(lstInvoice.ListItems(AA).SubItems(160)), lstInvoice.ListItems(AA).SubItems(160), 0)
    End If
    'TotalApprove = IIf(IsNumeric(lstInvoice.ListItems(AA).SubItems(161)), lstInvoice.ListItems(AA).SubItems(161), 0)
    
    tmpMBMAdj = IIf(IsNumeric(MBM!MBMAdjAnnLmt), MBM!MBMAdjAnnLmt, 0)
    tmpNewMBMAdj = IIf(IsNumeric(lstInvoice.ListItems(AA).SubItems(113)), lstInvoice.ListItems(AA).SubItems(113), 0)
    tmpMBMAdj = tmpMBMAdj - tmpPrevAdjAmt + tmpNewMBMAdj
    MBM!MBMAdjAnnLmt = CDbl(tmpMBMAdj)
    tmpAccumAmt = IIf(IsNumeric(MBM!MBMAccumAmt), MBM!MBMAccumAmt, 0)
    MBM!MBMAccumAmt = tmpAccumAmt - tmpFinalApp + CDbl(TotalApprove)
    tmpMaternity = IIf(IsNumeric(A8), A8, 0)
    
    Dim MBMBalLmt As Long
    MBMBalLmt = IIf(IsNumeric(MBM!MBMCLNBalLmt), MBM!MBMCLNBalLmt, 0)
    
    If lstInvoice.ListItems(AA).SubItems(70) <> "FULL" Then
        If Mid(CboPayto, 1, 1) = "M" Or Trim(strGRPMCOStatus) = "F" Then
            If strPLNGrossStatus = "N" Then
                'FinalBalance = MBM!MBMCLNBalLmt + tmpFinalApp - (TotalApprove - (A32 + A33 + tmpMaternity))
                FinalBalance = MBMBalLmt + tmpFinalApp - (TotalApprove - (A32 + A33 + tmpMaternity))
            ElseIf strPLNGrossStatus = "Y" Then
                'FinalBalance = MBM!MBMCLNBalLmt + (tmpFinalApp - (TotalApprove - (A32 + A33 + tmpMaternity)))
                FinalBalance = MBMBalLmt + ((tmpFinalApp) - ((TotalApprove) - (CDbl(A32) + CDbl(A33) + CDbl(tmpMaternity))))
            End If
        Else
            If strPLNGrossStatus = "N" Then
                'FinalBalance = (CDbl(MBM!MBMCLNBalLmt) + (CDbl(tmpFinalApp)) - (CDbl(A32) + CDbl(A33) + CDbl(tmpMaternity))) - (CDbl(TotalApprove) - (CDbl(A32) + CDbl(A33) + CDbl(tmpMaternity)))
                FinalBalance = (CDbl(MBMBalLmt) + (CDbl(tmpFinalApp)) - (CDbl(A32) + CDbl(A33) + CDbl(tmpMaternity))) - (CDbl(TotalApprove) - (CDbl(A32) + CDbl(A33) + CDbl(tmpMaternity)))
            ElseIf strPLNGrossStatus = "Y" Then
                'FinalBalance = (CDbl(MBM!MBMCLNBalLmt) + CDbl(tmpFinalApp)) - (CDbl(tmpPay2CLN) - (CDbl(A32) + CDbl(A33)))
                'FinalBalance = (CDbl(MBM!MBMCLNBalLmt) + (CDbl(tmpPay2CLN) - (CDbl(A32) + CDbl(A33) + CDbl(tmpMaternity)))) - (CDbl(tmpPay2CLN) - (CDbl(A32) + CDbl(A33) + CDbl(tmpMaternity)))
                FinalBalance = (CDbl(MBMBalLmt) + (CDbl(tmpPay2CLN) - (CDbl(A32) + CDbl(A33) + CDbl(tmpMaternity)))) - (CDbl(tmpPay2CLN) - (CDbl(A32) + CDbl(A33) + CDbl(tmpMaternity)))
            End If
        End If
        FinalBalance = IIf(FinalBalance < 0, 0, FinalBalance)
        MBM!MBMCLNBalLmt = FinalBalance
    End If
    MBM.update
    MBM.Close
    Set MBM = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub WSInfo()
Const METHOD_NAME = "WSInfo"
On Error GoTo debugs
Dim CISCLm As New ADODB.Recordset
Dim CISSql As String

    CISSql = "SELECT Top 1 * FROM Claim WHERE CASNumber='" & Trim(txtCaseNo) & "'"
    CISCLm.Open CISSql, MediConnCISWS, adOpenForwardOnly, adLockOptimistic
    With CISCLm
        lstRec.SubItems(158) = IIf(IsNumeric(!CLMPayableByMedix2c), !CLMPayableByMedix2c, 0)
        lstRec.SubItems(159) = IIf(IsNumeric(!CLMCLNExcess), !CLMCLNExcess, 0)
        lstRec.SubItems(160) = IIf(IsNumeric(!CLMTotalActual), !CLMTotalActual, 0)
        lstRec.SubItems(161) = IIf(IsNumeric(!CLMFinalApprove), !CLMFinalApprove, 0)
        If Mid(CboPayto, 1, 1) = "C" Then
            lstRec.SubItems(162) = IIf(IsNumeric(!CLMExcess), !CLMExcess, 0)
        Else
            lstRec.SubItems(162) = IIf(IsNumeric(!CLMPayableByMedix2M), !CLMPayableByMedix2M, 0)
        End If
        lstRec.SubItems(163) = IIf(IsNumeric(!CLMBalanceLimit), !CLMBalanceLimit, 0)
        lstRec.SubItems(164) = IIf(IsNumeric(!CLMPreviousBalance), !CLMPreviousBalance, 0)
        lstRec.SubItems(165) = IIf(IsNumeric(!CLMPaidBYMedix2M), !CLMPaidBYMedix2M, 0)
        lstRec.SubItems(166) = IIf(IsNumeric(!CLMPaidBYMedix2C), !CLMPaidBYMedix2C, 0)
        txtExcess = IIf(IsNumeric(!CLMTotalExcess), !CLMTotalExcess, 0)
    End With
    CISCLm.Close
    Set CISCLm = Nothing
    
    CISSql = "SELECT Top 1 * FROM CLMActualNew WHERE CLMNo='" & Trim(txtCaseNo) & "'"
    CISCLm.Open CISSql, MediConnCISWS, adOpenForwardOnly, adLockOptimistic

    With CISCLm
        lstRec.SubItems(167) = IIf(IsNumeric(!CLMGPFeeBasicMed), !CLMGPFeeBasicMed, 0)
        lstRec.SubItems(168) = IIf(IsNumeric(!CLMGPFeeMedInj), !CLMGPFeeMedInj, 0)
        lstRec.SubItems(169) = IIf(IsNumeric(!CLMGPFeeLTMed), !CLMGPFeeLTMed, 0)
        lstRec.SubItems(170) = IIf(IsNumeric(!CLMSPFeeBasicMed), !CLMSPFeeBasicMed, 0)
        lstRec.SubItems(171) = IIf(IsNumeric(!CLMSPFeeLTMed), !CLMSPFeeLTMed, 0)
        lstRec.SubItems(172) = IIf(IsNumeric(!CLMSPFeeMedInj), !CLMSPFeeMedInj, 0)
        lstRec.SubItems(173) = IIf(IsNumeric(!CLMXRay), !CLMXRay, 0)
        lstRec.SubItems(174) = IIf(IsNumeric(!CLMUltraSound), !CLMUltraSound, 0)
        lstRec.SubItems(175) = IIf(IsNumeric(!CLMECG), !CLMECG, 0)
        lstRec.SubItems(176) = IIf(IsNumeric(!CLMDressing), !CLMDressing, 0)
        lstRec.SubItems(177) = IIf(IsNumeric(!CLMNebulizer), !CLMNebulizer, 0)
        lstRec.SubItems(178) = IIf(IsNumeric(!CLMID), !CLMID, 0)
        lstRec.SubItems(179) = IIf(IsNumeric(!CLMTSSTO), !CLMTSSTO, 0)
        lstRec.SubItems(180) = IIf(IsNumeric(!CLMMinorSurgery), !CLMMinorSurgery, 0)
        lstRec.SubItems(181) = IIf(IsNumeric(!CLMPOthers), !CLMPOthers, 0)
        lstRec.SubItems(182) = IIf(IsNumeric(!CLMBloodDrawing), !CLMBloodDrawing, 0)
        lstRec.SubItems(183) = IIf(IsNumeric(!CLMBloodGlucose), !CLMBloodGlucose, 0)
        lstRec.SubItems(184) = IIf(IsNumeric(!CLMUrine), !CLMUrine, 0)
        lstRec.SubItems(185) = IIf(IsNumeric(!CLMPapSmear), !CLMPapSmear, 0)
        lstRec.SubItems(186) = IIf(IsNumeric(!CLMLOthers), !CLMLOthers, 0)
        lstRec.SubItems(187) = IIf(IsNumeric(!CLMDTPPolio), !CLMDTPPolio, 0)
        lstRec.SubItems(188) = IIf(IsNumeric(!CLMDTPPolioHIB), !CLMDTPPolioHIB, 0)
        lstRec.SubItems(189) = IIf(IsNumeric(!CLMHepatitisB), !CLMHepatitisB, 0)
        lstRec.SubItems(190) = IIf(IsNumeric(!CLMMMR), !CLMMMR, 0)
        lstRec.SubItems(191) = IIf(IsNumeric(!CLMIOthers), !CLMIOthers, 0)
        lstRec.SubItems(192) = IIf(IsNumeric(!CLMDental), !CLMDental, 0)
        lstRec.SubItems(193) = IIf(IsNumeric(!CLMMaternity), !CLMMaternity, 0)
        lstRec.SubItems(194) = IIf(IsNumeric(!CLMPRVCheckup), !CLMPRVCheckup, 0)
        lstRec.SubItems(195) = IIf(IsNumeric(!CLMVision), !CLMVision, 0)
        lstRec.SubItems(196) = IIf(IsNumeric(!CLMTyphiadJabNoDed), !CLMTyphiadJabNoDed, 0)
        lstRec.SubItems(197) = IIf(IsNumeric(!CLMOthersNoDed), !CLMOthersNoDed, 0)
        
    End With
    CISCLm.Close
    Set CISCLm = Nothing
    CISSql = "SELECT Top 1 * FROM CLMClinicAppNew WHERE CASNumber='" & Trim(txtCaseNo) & "'"
    CISCLm.Open CISSql, MediConnCISWS, adOpenForwardOnly, adLockOptimistic

    With CISCLm
        lstRec.SubItems(425) = IIf(IsNumeric(!CLMGPFeeBasicMed), !CLMGPFeeBasicMed, 0)
        lstRec.SubItems(426) = IIf(IsNumeric(!CLMGPFeeMedInj), !CLMGPFeeMedInj, 0)
        lstRec.SubItems(427) = IIf(IsNumeric(!CLMGPFeeLTMed), !CLMGPFeeLTMed, 0)
        lstRec.SubItems(428) = IIf(IsNumeric(!CLMSPFeeBasicMed), !CLMSPFeeBasicMed, 0)
        lstRec.SubItems(429) = IIf(IsNumeric(!CLMSPFeeLTMed), !CLMSPFeeLTMed, 0)
        lstRec.SubItems(430) = IIf(IsNumeric(!CLMSPFeeMedInj), !CLMSPFeeMedInj, 0)
        lstRec.SubItems(431) = IIf(IsNumeric(!CLMXRay), !CLMXRay, 0)
        lstRec.SubItems(432) = IIf(IsNumeric(!CLMUltraSound), !CLMUltraSound, 0)
        lstRec.SubItems(433) = IIf(IsNumeric(!CLMECG), !CLMECG, 0)
        lstRec.SubItems(434) = IIf(IsNumeric(!CLMDressing), !CLMDressing, 0)
        lstRec.SubItems(435) = IIf(IsNumeric(!CLMNebulizer), !CLMNebulizer, 0)
        lstRec.SubItems(436) = IIf(IsNumeric(!CLMID), !CLMID, 0)
        lstRec.SubItems(437) = IIf(IsNumeric(!CLMTSSTO), !CLMTSSTO, 0)
        lstRec.SubItems(438) = IIf(IsNumeric(!CLMMinorSurgery), !CLMMinorSurgery, 0)
        lstRec.SubItems(439) = IIf(IsNumeric(!CLMPOthers), !CLMPOthers, 0)
        lstRec.SubItems(440) = IIf(IsNumeric(!CLMBloodDrawing), !CLMBloodDrawing, 0)
        lstRec.SubItems(441) = IIf(IsNumeric(!CLMBloodGlucose), !CLMBloodGlucose, 0)
        lstRec.SubItems(442) = IIf(IsNumeric(!CLMUrine), !CLMUrine, 0)
        lstRec.SubItems(443) = IIf(IsNumeric(!CLMPapSmear), !CLMPapSmear, 0)
        lstRec.SubItems(444) = IIf(IsNumeric(!CLMLOthers), !CLMLOthers, 0)
        lstRec.SubItems(445) = IIf(IsNumeric(!CLMDTPPolio), !CLMDTPPolio, 0)
        lstRec.SubItems(446) = IIf(IsNumeric(!CLMDTPPolioHIB), !CLMDTPPolioHIB, 0)
        lstRec.SubItems(447) = IIf(IsNumeric(!CLMHepatitisB), !CLMHepatitisB, 0)
        lstRec.SubItems(448) = IIf(IsNumeric(!CLMMMR), !CLMMMR, 0)
        lstRec.SubItems(449) = IIf(IsNumeric(!CLMIOthers), !CLMIOthers, 0)
        lstRec.SubItems(450) = IIf(IsNumeric(!CLMDental), !CLMDental, 0)
        lstRec.SubItems(451) = IIf(IsNumeric(!CLMMaternity), !CLMMaternity, 0)
        lstRec.SubItems(452) = IIf(IsNumeric(!CLMPRVCheckup), !CLMPRVCheckup, 0)
        lstRec.SubItems(453) = IIf(IsNumeric(!CLMVision), !CLMVision, 0)
        lstRec.SubItems(454) = IIf(IsNumeric(!CLMTyphiadJabNoDed), !CLMTyphiadJabNoDed, 0)
        lstRec.SubItems(455) = IIf(IsNumeric(!CLMOthersNoDed), !CLMOthersNoDed, 0)
        
    End With
    CISCLm.Close
    Set CISCLm = Nothing
    
    ''Price
    
    CISSql = "SELECT  MEDCHARGES1, MEDCHARGES2, MEDCHARGES3, MEDCHARGES4, MEDCHARGES5, MEDCHARGES6, MEDCHARGES7  FROM CISWORKSHEET..CLMMEDCHARGES WHERE CASNumber='" & Trim(txtCaseNo) & "'"
    CISCLm.Open CISSql, MediConnCISWS, adOpenForwardOnly, adLockOptimistic
    With CISCLm
    
        lstRec.SubItems(482) = IIf(IsNumeric(!MEDCHARGES1), !MEDCHARGES1, 0)
        
        lstRec.SubItems(483) = IIf(IsNumeric(!MEDCHARGES2), !MEDCHARGES2, 0)
        lstRec.SubItems(484) = IIf(IsNumeric(!MEDCHARGES3), !MEDCHARGES3, 0)
        lstRec.SubItems(485) = IIf(IsNumeric(!MEDCHARGES4), !MEDCHARGES4, 0)
        lstRec.SubItems(486) = IIf(IsNumeric(!MEDCHARGES5), !MEDCHARGES5, 0)
        lstRec.SubItems(487) = IIf(IsNumeric(!MEDCHARGES6), !MEDCHARGES6, 0)
        lstRec.SubItems(488) = IIf(IsNumeric(!MEDCHARGES7), !MEDCHARGES7, 0)

    End With
    CISCLm.Close
    Set CISCLm = Nothing
    
    CISSql = "SELECT Top 1 * FROM CLMClinicLmtNew WHERE CASNumber='" & Trim(txtCaseNo) & "'"
    CISCLm.Open CISSql, MediConnCISWS, adOpenForwardOnly, adLockOptimistic

    With CISCLm
        lstRec.SubItems(210) = IIf(IsNumeric(!CLMGPFeeBasicMed), !CLMGPFeeBasicMed, 0)
        lstRec.SubItems(211) = IIf(IsNumeric(!CLMGPFeeMedInj), !CLMGPFeeMedInj, 0)
        lstRec.SubItems(212) = IIf(IsNumeric(!CLMGPFeeLTMed), !CLMGPFeeLTMed, 0)
        lstRec.SubItems(213) = IIf(IsNumeric(!CLMSPFeeBasicMed), !CLMSPFeeBasicMed, 0)
        lstRec.SubItems(214) = IIf(IsNumeric(!CLMSPFeeLTMed), !CLMSPFeeLTMed, 0)
        lstRec.SubItems(215) = IIf(IsNumeric(!CLMSPFeeMedInj), !CLMSPFeeMedInj, 0)
        lstRec.SubItems(216) = IIf(IsNumeric(!CLMXRay), !CLMXRay, 0)
        lstRec.SubItems(217) = IIf(IsNumeric(!CLMUltraSound), !CLMUltraSound, 0)
        lstRec.SubItems(218) = IIf(IsNumeric(!CLMECG), !CLMECG, 0)
        lstRec.SubItems(219) = IIf(IsNumeric(!CLMDressing), !CLMDressing, 0)
        lstRec.SubItems(220) = IIf(IsNumeric(!CLMNebulizer), !CLMNebulizer, 0)
        lstRec.SubItems(221) = IIf(IsNumeric(!CLMID), !CLMID, 0)
        lstRec.SubItems(222) = IIf(IsNumeric(!CLMTSSTO), !CLMTSSTO, 0)
        lstRec.SubItems(223) = IIf(IsNumeric(!CLMMinorSurgery), !CLMMinorSurgery, 0)
        lstRec.SubItems(224) = IIf(IsNumeric(!CLMPOthers), !CLMPOthers, 0)
        lstRec.SubItems(225) = IIf(IsNumeric(!CLMBloodDrawing), !CLMBloodDrawing, 0)
        lstRec.SubItems(226) = IIf(IsNumeric(!CLMBloodGlucose), !CLMBloodGlucose, 0)
        lstRec.SubItems(227) = IIf(IsNumeric(!CLMUrine), !CLMUrine, 0)
        lstRec.SubItems(228) = IIf(IsNumeric(!CLMPapSmear), !CLMPapSmear, 0)
        lstRec.SubItems(229) = IIf(IsNumeric(!CLMLOthers), !CLMLOthers, 0)
        lstRec.SubItems(230) = IIf(IsNumeric(!CLMDTPPolio), !CLMDTPPolio, 0)
        lstRec.SubItems(231) = IIf(IsNumeric(!CLMDTPPolioHIB), !CLMDTPPolioHIB, 0)
        lstRec.SubItems(232) = IIf(IsNumeric(!CLMHepatitisB), !CLMHepatitisB, 0)
        lstRec.SubItems(233) = IIf(IsNumeric(!CLMMMR), !CLMMMR, 0)
        lstRec.SubItems(234) = IIf(IsNumeric(!CLMIOthers), !CLMIOthers, 0)
        lstRec.SubItems(235) = IIf(IsNumeric(!CLMDental), !CLMDental, 0)
        lstRec.SubItems(236) = IIf(IsNumeric(!CLMMaternity), !CLMMaternity, 0)
        lstRec.SubItems(237) = IIf(IsNumeric(!CLMPRVCheckup), !CLMPRVCheckup, 0)
        lstRec.SubItems(238) = IIf(IsNumeric(!CLMVision), !CLMVision, 0)
        lstRec.SubItems(239) = IIf(IsNumeric(!CLMTyphiadJabNoDed), !CLMTyphiadJabNoDed, 0)
        lstRec.SubItems(240) = IIf(IsNumeric(!CLMOthersNoDed), !CLMOthersNoDed, 0)
        
    End With
    CISCLm.Close
    Set CISCLm = Nothing
    
    CISSql = "SELECT Top 1 * FROM CLMClinicNew WHERE CASNumber='" & Trim(txtCaseNo) & "'"
    CISCLm.Open CISSql, MediConnCISWS, adOpenForwardOnly, adLockOptimistic
    With CISCLm
        lstRec.SubItems(382) = IIf(IsNumeric(!CLMGPFeeBasicMed), !CLMGPFeeBasicMed, 0)
        lstRec.SubItems(383) = IIf(IsNumeric(!CLMGPFeeMedInj), !CLMGPFeeMedInj, 0)
        lstRec.SubItems(384) = IIf(IsNumeric(!CLMGPFeeLTMed), !CLMGPFeeLTMed, 0)
        lstRec.SubItems(385) = IIf(IsNumeric(!CLMSPFeeBasicMed), !CLMSPFeeBasicMed, 0)
        lstRec.SubItems(386) = IIf(IsNumeric(!CLMSPFeeLTMed), !CLMSPFeeLTMed, 0)
        lstRec.SubItems(387) = IIf(IsNumeric(!CLMSPFeeMedInj), !CLMSPFeeMedInj, 0)
        lstRec.SubItems(388) = IIf(IsNumeric(!CLMXRay), !CLMXRay, 0)
        lstRec.SubItems(389) = IIf(IsNumeric(!CLMUltraSound), !CLMUltraSound, 0)
        lstRec.SubItems(390) = IIf(IsNumeric(!CLMECG), !CLMECG, 0)
        lstRec.SubItems(391) = IIf(IsNumeric(!CLMDressing), !CLMDressing, 0)
        lstRec.SubItems(392) = IIf(IsNumeric(!CLMNebulizer), !CLMNebulizer, 0)
        lstRec.SubItems(393) = IIf(IsNumeric(!CLMID), !CLMID, 0)
        lstRec.SubItems(394) = IIf(IsNumeric(!CLMTSSTO), !CLMTSSTO, 0)
        lstRec.SubItems(395) = IIf(IsNumeric(!CLMMinorSurgery), !CLMMinorSurgery, 0)
        lstRec.SubItems(396) = IIf(IsNumeric(!CLMPOthers), !CLMPOthers, 0)
        lstRec.SubItems(397) = IIf(IsNumeric(!CLMBloodDrawing), !CLMBloodDrawing, 0)
        lstRec.SubItems(398) = IIf(IsNumeric(!CLMBloodGlucose), !CLMBloodGlucose, 0)
        lstRec.SubItems(399) = IIf(IsNumeric(!CLMUrine), !CLMUrine, 0)
        lstRec.SubItems(400) = IIf(IsNumeric(!CLMPapSmear), !CLMPapSmear, 0)
        lstRec.SubItems(401) = IIf(IsNumeric(!CLMLOthers), !CLMLOthers, 0)
        lstRec.SubItems(402) = IIf(IsNumeric(!CLMDTPPolio), !CLMDTPPolio, 0)
        lstRec.SubItems(403) = IIf(IsNumeric(!CLMDTPPolioHIB), !CLMDTPPolioHIB, 0)
        lstRec.SubItems(404) = IIf(IsNumeric(!CLMHepatitisB), !CLMHepatitisB, 0)
        lstRec.SubItems(405) = IIf(IsNumeric(!CLMMMR), !CLMMMR, 0)
        lstRec.SubItems(406) = IIf(IsNumeric(!CLMIOthers), !CLMIOthers, 0)
        lstRec.SubItems(407) = IIf(IsNumeric(!CLMDental), !CLMDental, 0)
        lstRec.SubItems(408) = IIf(IsNumeric(!CLMMaternity), !CLMMaternity, 0)
        lstRec.SubItems(409) = IIf(IsNumeric(!CLMPRVCheckup), !CLMPRVCheckup, 0)
        lstRec.SubItems(410) = IIf(IsNumeric(!CLMVision), !CLMVision, 0)
        lstRec.SubItems(411) = IIf(IsNumeric(!CLMTyphiadJabNoDed), !CLMTyphiadJabNoDed, 0)
        lstRec.SubItems(412) = IIf(IsNumeric(!CLMOthersNoDed), !CLMOthersNoDed, 0)
        
    End With
    CISCLm.Close
    Set CISCLm = Nothing
    
    CISSql = "SELECT Top 1 * FROM CLMAppNew WHERE CASNumber='" & Trim(txtCaseNo) & "'"
    CISCLm.Open CISSql, MediConnCISWS, adOpenForwardOnly, adLockOptimistic
    With CISCLm
        lstRec.SubItems(253) = IIf(IsNumeric(!CLMGPFeeBasicMed), !CLMGPFeeBasicMed, 0)
        lstRec.SubItems(254) = IIf(IsNumeric(!CLMGPFeeMedInj), !CLMGPFeeMedInj, 0)
        lstRec.SubItems(255) = IIf(IsNumeric(!CLMGPFeeLTMed), !CLMGPFeeLTMed, 0)
        lstRec.SubItems(256) = IIf(IsNumeric(!CLMSPFeeBasicMed), !CLMSPFeeBasicMed, 0)
        lstRec.SubItems(257) = IIf(IsNumeric(!CLMSPFeeLTMed), !CLMSPFeeLTMed, 0)
        lstRec.SubItems(258) = IIf(IsNumeric(!CLMSPFeeMedInj), !CLMSPFeeMedInj, 0)
        lstRec.SubItems(259) = IIf(IsNumeric(!CLMXRay), !CLMXRay, 0)
        lstRec.SubItems(260) = IIf(IsNumeric(!CLMUltraSound), !CLMUltraSound, 0)
        lstRec.SubItems(261) = IIf(IsNumeric(!CLMECG), !CLMECG, 0)
        lstRec.SubItems(262) = IIf(IsNumeric(!CLMDressing), !CLMDressing, 0)
        lstRec.SubItems(263) = IIf(IsNumeric(!CLMNebulizer), !CLMNebulizer, 0)
        lstRec.SubItems(264) = IIf(IsNumeric(!CLMID), !CLMID, 0)
        lstRec.SubItems(265) = IIf(IsNumeric(!CLMTSSTO), !CLMTSSTO, 0)
        lstRec.SubItems(266) = IIf(IsNumeric(!CLMMinorSurgery), !CLMMinorSurgery, 0)
        lstRec.SubItems(267) = IIf(IsNumeric(!CLMPOthers), !CLMPOthers, 0)
        lstRec.SubItems(268) = IIf(IsNumeric(!CLMBloodDrawing), !CLMBloodDrawing, 0)
        lstRec.SubItems(269) = IIf(IsNumeric(!CLMBloodGlucose), !CLMBloodGlucose, 0)
        lstRec.SubItems(270) = IIf(IsNumeric(!CLMUrine), !CLMUrine, 0)
        lstRec.SubItems(271) = IIf(IsNumeric(!CLMPapSmear), !CLMPapSmear, 0)
        lstRec.SubItems(272) = IIf(IsNumeric(!CLMLOthers), !CLMLOthers, 0)
        lstRec.SubItems(273) = IIf(IsNumeric(!CLMDTPPolio), !CLMDTPPolio, 0)
        lstRec.SubItems(274) = IIf(IsNumeric(!CLMDTPPolioHIB), !CLMDTPPolioHIB, 0)
        lstRec.SubItems(275) = IIf(IsNumeric(!CLMHepatitisB), !CLMHepatitisB, 0)
        lstRec.SubItems(276) = IIf(IsNumeric(!CLMMMR), !CLMMMR, 0)
        lstRec.SubItems(277) = IIf(IsNumeric(!CLMIOthers), !CLMIOthers, 0)
        lstRec.SubItems(278) = IIf(IsNumeric(!CLMDental), !CLMDental, 0)
        lstRec.SubItems(279) = IIf(IsNumeric(!CLMMaternity), !CLMMaternity, 0)
        lstRec.SubItems(280) = IIf(IsNumeric(!CLMPRVCheckup), !CLMPRVCheckup, 0)
        lstRec.SubItems(281) = IIf(IsNumeric(!CLMVision), !CLMVision, 0)
        lstRec.SubItems(282) = IIf(IsNumeric(!CLMTyphiadJabNoDed), !CLMTyphiadJabNoDed, 0)
        lstRec.SubItems(283) = IIf(IsNumeric(!CLMOthersNoDed), !CLMOthersNoDed, 0)
        
    End With
    CISCLm.Close
    Set CISCLm = Nothing
    
    CISSql = "SELECT Top 1 * FROM CLMMemberNew WHERE CASNumber='" & Trim(txtCaseNo) & "'"
    CISCLm.Open CISSql, MediConnCISWS, adOpenForwardOnly, adLockOptimistic
    With CISCLm
        lstRec.SubItems(339) = IIf(IsNumeric(!CLMGPFeeBasicMed), !CLMGPFeeBasicMed, 0)
        lstRec.SubItems(340) = IIf(IsNumeric(!CLMGPFeeMedInj), !CLMGPFeeMedInj, 0)
        lstRec.SubItems(341) = IIf(IsNumeric(!CLMGPFeeLTMed), !CLMGPFeeLTMed, 0)
        lstRec.SubItems(342) = IIf(IsNumeric(!CLMSPFeeBasicMed), !CLMSPFeeBasicMed, 0)
        lstRec.SubItems(343) = IIf(IsNumeric(!CLMSPFeeLTMed), !CLMSPFeeLTMed, 0)
        lstRec.SubItems(344) = IIf(IsNumeric(!CLMSPFeeMedInj), !CLMSPFeeMedInj, 0)
        lstRec.SubItems(345) = IIf(IsNumeric(!CLMXRay), !CLMXRay, 0)
        lstRec.SubItems(346) = IIf(IsNumeric(!CLMUltraSound), !CLMUltraSound, 0)
        lstRec.SubItems(347) = IIf(IsNumeric(!CLMECG), !CLMECG, 0)
        lstRec.SubItems(348) = IIf(IsNumeric(!CLMDressing), !CLMDressing, 0)
        lstRec.SubItems(349) = IIf(IsNumeric(!CLMNebulizer), !CLMNebulizer, 0)
        lstRec.SubItems(350) = IIf(IsNumeric(!CLMID), !CLMID, 0)
        lstRec.SubItems(351) = IIf(IsNumeric(!CLMTSSTO), !CLMTSSTO, 0)
        lstRec.SubItems(352) = IIf(IsNumeric(!CLMMinorSurgery), !CLMMinorSurgery, 0)
        lstRec.SubItems(353) = IIf(IsNumeric(!CLMPOthers), !CLMPOthers, 0)
        lstRec.SubItems(354) = IIf(IsNumeric(!CLMBloodDrawing), !CLMBloodDrawing, 0)
        lstRec.SubItems(355) = IIf(IsNumeric(!CLMBloodGlucose), !CLMBloodGlucose, 0)
        lstRec.SubItems(356) = IIf(IsNumeric(!CLMUrine), !CLMUrine, 0)
        lstRec.SubItems(357) = IIf(IsNumeric(!CLMPapSmear), !CLMPapSmear, 0)
        lstRec.SubItems(358) = IIf(IsNumeric(!CLMLOthers), !CLMLOthers, 0)
        lstRec.SubItems(359) = IIf(IsNumeric(!CLMDTPPolio), !CLMDTPPolio, 0)
        lstRec.SubItems(360) = IIf(IsNumeric(!CLMDTPPolioHIB), !CLMDTPPolioHIB, 0)
        lstRec.SubItems(361) = IIf(IsNumeric(!CLMHepatitisB), !CLMHepatitisB, 0)
        lstRec.SubItems(362) = IIf(IsNumeric(!CLMMMR), !CLMMMR, 0)
        lstRec.SubItems(363) = IIf(IsNumeric(!CLMIOthers), !CLMIOthers, 0)
        lstRec.SubItems(364) = IIf(IsNumeric(!CLMDental), !CLMDental, 0)
        lstRec.SubItems(365) = IIf(IsNumeric(!CLMMaternity), !CLMMaternity, 0)
        lstRec.SubItems(366) = IIf(IsNumeric(!CLMPRVCheckup), !CLMPRVCheckup, 0)
        lstRec.SubItems(367) = IIf(IsNumeric(!CLMVision), !CLMVision, 0)
        lstRec.SubItems(368) = IIf(IsNumeric(!CLMTyphiadJabNoDed), !CLMTyphiadJabNoDed, 0)
        lstRec.SubItems(369) = IIf(IsNumeric(!CLMOthersNoDed), !CLMOthersNoDed, 0)

    End With
    CISCLm.Close
    Set CISCLm = Nothing
    
    CISSql = "SELECT Top 1 * FROM CLMLimitNew WHERE CASNumber='" & Trim(txtCaseNo) & "'"
    CISCLm.Open CISSql, MediConnCISWS, adOpenForwardOnly, adLockOptimistic
    With CISCLm
        lstRec.SubItems(297) = IIf(IsNumeric(!CLMGPFeeBasicMed), !CLMGPFeeBasicMed, 0)
        lstRec.SubItems(298) = IIf(IsNumeric(!CLMGPFeeMedInj), !CLMGPFeeMedInj, 0)
        lstRec.SubItems(299) = IIf(IsNumeric(!CLMGPFeeLTMed), !CLMGPFeeLTMed, 0)
        lstRec.SubItems(300) = IIf(IsNumeric(!CLMSPFeeBasicMed), !CLMSPFeeBasicMed, 0)
        lstRec.SubItems(301) = IIf(IsNumeric(!CLMSPFeeLTMed), !CLMSPFeeLTMed, 0)
        lstRec.SubItems(302) = IIf(IsNumeric(!CLMSPFeeMedInj), !CLMSPFeeMedInj, 0)
        lstRec.SubItems(303) = IIf(IsNumeric(!CLMXRay), !CLMXRay, 0)
        lstRec.SubItems(304) = IIf(IsNumeric(!CLMUltraSound), !CLMUltraSound, 0)
        lstRec.SubItems(305) = IIf(IsNumeric(!CLMECG), !CLMECG, 0)
        lstRec.SubItems(306) = IIf(IsNumeric(!CLMDressing), !CLMDressing, 0)
        lstRec.SubItems(307) = IIf(IsNumeric(!CLMNebulizer), !CLMNebulizer, 0)
        lstRec.SubItems(308) = IIf(IsNumeric(!CLMID), !CLMID, 0)
        lstRec.SubItems(309) = IIf(IsNumeric(!CLMTSSTO), !CLMTSSTO, 0)
        lstRec.SubItems(310) = IIf(IsNumeric(!CLMMinorSurgery), !CLMMinorSurgery, 0)
        lstRec.SubItems(311) = IIf(IsNumeric(!CLMPOthers), !CLMPOthers, 0)
        lstRec.SubItems(312) = IIf(IsNumeric(!CLMBloodDrawing), !CLMBloodDrawing, 0)
        lstRec.SubItems(313) = IIf(IsNumeric(!CLMBloodGlucose), !CLMBloodGlucose, 0)
        lstRec.SubItems(314) = IIf(IsNumeric(!CLMUrine), !CLMUrine, 0)
        lstRec.SubItems(315) = IIf(IsNumeric(!CLMPapSmear), !CLMPapSmear, 0)
        lstRec.SubItems(316) = IIf(IsNumeric(!CLMLOthers), !CLMLOthers, 0)
        lstRec.SubItems(317) = IIf(IsNumeric(!CLMDTPPolio), !CLMDTPPolio, 0)
        lstRec.SubItems(318) = IIf(IsNumeric(!CLMDTPPolioHIB), !CLMDTPPolioHIB, 0)
        lstRec.SubItems(319) = IIf(IsNumeric(!CLMHepatitisB), !CLMHepatitisB, 0)
        lstRec.SubItems(320) = IIf(IsNumeric(!CLMMMR), !CLMMMR, 0)
        lstRec.SubItems(321) = IIf(IsNumeric(!CLMIOthers), !CLMIOthers, 0)
        lstRec.SubItems(322) = IIf(IsNumeric(!CLMDental), !CLMDental, 0)
        lstRec.SubItems(323) = IIf(IsNumeric(!CLMMaternity), !CLMMaternity, 0)
        lstRec.SubItems(324) = IIf(IsNumeric(!CLMPRVCheckup), !CLMPRVCheckup, 0)
        lstRec.SubItems(325) = IIf(IsNumeric(!CLMVision), !CLMVision, 0)
        lstRec.SubItems(326) = IIf(IsNumeric(!CLMTyphiadJabNoDed), !CLMTyphiadJabNoDed, 0)
        lstRec.SubItems(327) = IIf(IsNumeric(!CLMOthersNoDed), !CLMOthersNoDed, 0)
    End With
    CISCLm.Close
    Set CISCLm = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Public Sub ValidateIbmUser()
Const METHOD_NAME = "ValidateIbmUser"
On Error GoTo debugs
    Dim txtGrpCompany As String
    Dim Balcmd As New ADODB.Command
    Dim BalPrm1 As New ADODB.Parameter
    Dim BalRes As New ADODB.Recordset
    Dim MBMNo As String
    MBMNo = txtMBMNo.Text
    
    If InStr(1, MBMNo, "-") Then
        Dim sArr() As String
        sArr = Split(MBMNo, "-")
        MBMNo = sArr(0)
    End If
    
        Set Balcmd.ActiveConnection = MediConnCISDB
        Balcmd.CommandText = "sp_IBMMemberInformation"
        Balcmd.CommandType = adCmdStoredProc
        Balcmd.CommandTimeout = 300
        Set BalPrm1 = Balcmd.CreateParameter("MemeberNo", adVarChar, adParamInput, 2000, Trim(MBMNo))
        Balcmd.Parameters.Append BalPrm1
        BalRes.CursorLocation = adUseClient
        BalRes.CursorType = adOpenForwardOnly
        BalRes.CacheSize = 100
        Set BalRes = Balcmd.Execute
        Do While Not BalRes.EOF
            txtGrpCompany = BalRes!Number
            BalRes.MoveNext
        Loop
        BalRes.Close
        Set BalRes = Nothing
               
        If txtGrpCompany <> "0" Then
           bIbmUser = True
        Else
           bIbmUser = False
        End If

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub UpdateLstView()
Const METHOD_NAME = "UpdateLstView"
On Error GoTo debugs
Dim strPayee As String
'Dim FFSExc As New ADODB.Recordset
Dim FFSSql As String

    Set lstRec = lstInvoice.SelectedItem
    If strPayTo = "C" Then '
        strPayee = txtClinicCode
    Else
        strPayee = "MEMBER"
    End If
         Dim sDt As String
          Dim SpValue As Double
          
     If Mid(cmbPharmacy, 1, 1) = "Y" Then
          'By 10 is for 10%
           'A29 = ((CDbl(lblDispBillAmt) / 10))
          Call DisplayBillAmt
     ElseIf Mid(cmbPharmacy, 1, 1) = "N" Then
            '  Call ValidateIbmUser
            '      If bIbmUser Then
            '        sDt = Year(txtVisitDate)
            '            If sDt >= "2014" Then
            '               If (Trim(txtClinicCode.Text) = Trim("NPS")) Then
            '                If Mid(cboRefLetter, 1, 1) = "N" Then
            '                    'A29 = ((CDbl(lblDispBillAmt) / 4) * -1)
            '                    A29 = ((CDbl(lblDispBillAmt) / 4) * 1)
            '                End If
            '               ElseIf (Trim(txtClinicCode.Text) = Trim("NPS")) Then
            '                    A29 = "5.00"
            '               Else
            '                    Dim FFSExc As New ADODB.Recordset
            '                    FFSSql = ""
            '                    FFSSql = "select Top 1  CLNType,CLNMCOProcfeePerc from HISDB.dbo.CLN where CLNCode= '" & txtClinicCode & "'"
            '                    FFSExc.Open FFSSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
            '                        With FFSExc
            '                        If !CLNType = "G" Then
            '                             A29 = "5.00"
            '                        ElseIf !CLNType = "S" Then
            '                            If Mid(cboRefLetter, 1, 1) = "N" Then
            '                                 A29 = ((CDbl(lblDispBillAmt) / 4) * 1)
            '                            End If
            '                        End If
            '                        End With
            '               End If
            '            ElseIf sDt = "2013" Then
            '            If (Trim(txtClinicCode.Text) = Trim("NPS")) Then
            '                If Mid(cboRefLetter, 1, 1) = "N" Then
            '                    A29 = ((CDbl(lblDispBillAmt) / 4) * 1)
            '                End If
            '            Else
            '                Dim FFSExc1 As New ADODB.Recordset
            '                FFSSql = ""
            '                FFSSql = "select Top 1  CLNType,CLNMCOProcfeePerc from HISDB.dbo.CLN where CLNCode= '" & txtClinicCode & "'"
            '                FFSExc1.Open FFSSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
            '                        With FFSExc1
            '                        If !CLNType = "S" Then
            '                            If Mid(cboRefLetter, 1, 1) = "N" Then
            '                                 A29 = ((CDbl(lblDispBillAmt) / 4) * 1)
            '                            End If
            '                        End If
            '                        End With
            '
            '               End If
            '            End If
            '      Else
            '        If InStr(1, A29, "-") Then
            '            A29 = (CDbl(A29) * -1)
            '        Else
            '            A29 = A29
            '        End If
            '      End If
     Else
             '   Call ValidateIbmUser
             '     If bIbmUser Then
             '       sDt = Year(txtVisitDate)
             '           If sDt >= "2014" Then
             '           If (Trim(txtClinicCode.Text) = Trim("NPS")) Then
             '             If Mid(cboRefLetter, 1, 1) = "N" Then
             '                A29 = ((CDbl(lblDispBillAmt) / 4) * 1)
             '             End If
             '           ElseIf (Trim(txtClinicCode.Text) = Trim("NP")) Then
             '             A29 = "5.00"
             '           Else
             '                   Dim FFSExc2 As New ADODB.Recordset
             '                   FFSSql = ""
             '                   FFSSql = "select Top 1  CLNType,CLNMCOProcfeePerc from HISDB.dbo.CLN where CLNCode= '" & txtClinicCode & "'"
             '                   FFSExc2.Open FFSSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
             '                       With FFSExc2
             '                       If !CLNType = "G" Then
             '                            A29 = "5.00"
             '                       ElseIf !CLNType = "S" Then
             '                           If Mid(cboRefLetter, 1, 1) = "N" Then
             '                                A29 = ((CDbl(lblDispBillAmt) / 4) * 1)
             '                           End If
             '                       End If
             '                       End With
             '              End If
             '           ElseIf sDt = "2013" Then
             '              If (Trim(txtClinicCode.Text) = Trim("NPS")) Then
             '               If Mid(cboRefLetter, 1, 1) = "N" Then
             '                   A29 = ((CDbl(lblDispBillAmt) / 4) * 1)
             '               End If
             '               Else
             '                   Dim FFSExc3 As New ADODB.Recordset
             '                   FFSSql = ""
             '                   FFSSql = "select Top 1  CLNType,CLNMCOProcfeePerc from HISDB.dbo.CLN where CLNCode= '" & txtClinicCode & "'"
             '                   FFSExc3.Open FFSSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
             '                       With FFSExc3
             '                       If !CLNType = "S" Then
             '                           If Mid(cboRefLetter, 1, 1) = "N" Then
             '                                A29 = ((CDbl(lblDispBillAmt) / 4) * 1)
             '                           End If
             '                       End If
             '                       End With
             '              End If
             '           End If
              '    Else
              '
              '     If InStr(1, A29, "-") Then
              '          A29 = (CDbl(A29) * -1)
              '      Else
              '          A29 = A29
              '     End If
                   
           ' End If
        End If
    Call DisplayBillAmt
    Call getTotal
    lstRec.SubItems(1) = txtCaseNo
    lstRec.SubItems(2) = MembershipNo
    lstRec.SubItems(3) = CovPersonID
    lstRec.SubItems(4) = TxtPatientName
    lstRec.SubItems(5) = IIf(Len(Trim(txtNPClinic)), Trim(txtNPClinic), "")
    lstRec.SubItems(6) = IIf(Len(Trim(txtInvNo)), Trim(txtInvNo), "")
    lstRec.SubItems(7) = IIf(Len(Trim(txtBillNo)), txtBillNo, "")
    lstRec.SubItems(8) = IIf(IsDate(txtInvDate), txtInvDate, "__/__/____")
    lstRec.SubItems(9) = IIf(IsDate(txtVisitDate), txtVisitDate, "__/__/____")
    lstRec.SubItems(10) = IIf(isTime(txtVisitTime), txtVisitTime, "__:__")
    lstRec.SubItems(11) = IIf(IsNumeric(txtMC), txtMC, 0)
    lstRec.SubItems(12) = Trim(txtClinicCode)
    lstRec.SubItems(13) = CDbl(lblDispBillAmt)
    lstRec.SubItems(14) = Logon
    lstRec.SubItems(15) = IIf(IsNumeric(A1), A1, 0)
    lstRec.SubItems(16) = IIf(IsNumeric(A2), A2, 0)
    lstRec.SubItems(17) = IIf(IsNumeric(A3), A3, 0)
    lstRec.SubItems(18) = IIf(IsNumeric(A4), A4, 0)
    lstRec.SubItems(19) = IIf(IsNumeric(A5), A5, 0)
    lstRec.SubItems(20) = IIf(IsNumeric(A6), A6, 0)
    lstRec.SubItems(21) = IIf(IsNumeric(A7), A7, 0)
    lstRec.SubItems(22) = IIf(IsNumeric(A8), A8, 0)
    lstRec.SubItems(23) = IIf(IsNumeric(A9), A9, 0)
    lstRec.SubItems(24) = IIf(IsNumeric(A10), A10, 0)
    lstRec.SubItems(25) = IIf(IsNumeric(A11), A11, 0)
    lstRec.SubItems(26) = IIf(IsNumeric(A12), A12, 0)
    lstRec.SubItems(27) = IIf(IsNumeric(A13), A13, 0)
    lstRec.SubItems(28) = IIf(IsNumeric(A14), A14, 0)
    lstRec.SubItems(29) = IIf(IsNumeric(A15), A15, 0)
    lstRec.SubItems(30) = IIf(IsNumeric(A16), A16, 0)
    lstRec.SubItems(31) = IIf(IsNumeric(A17), A17, 0)
    lstRec.SubItems(32) = IIf(IsNumeric(A18), A18, 0)
    lstRec.SubItems(33) = IIf(IsNumeric(A19), A19, 0)
    lstRec.SubItems(34) = IIf(IsNumeric(A20), A20, 0)
    lstRec.SubItems(35) = IIf(IsNumeric(A21), A21, 0)
    lstRec.SubItems(36) = IIf(IsNumeric(A22), A22, 0)
    lstRec.SubItems(37) = IIf(IsNumeric(A23), A23, 0)
    lstRec.SubItems(38) = IIf(IsNumeric(A24), A24, 0)
    lstRec.SubItems(39) = IIf(IsNumeric(A25), A25, 0)
    lstRec.SubItems(40) = IIf(IsNumeric(A26), A26, 0)
    lstRec.SubItems(41) = IIf(IsNumeric(A27), A27, 0)
    lstRec.SubItems(42) = IIf(IsNumeric(A28), A28, 0)
    lstRec.SubItems(43) = IIf(IsNumeric(lblActCoPay), lblActCoPay, 0)
    
    lstRec.SubItems(44) = IIf(Len(Trim(txtPOthDesc)), Trim(txtPOthDesc), "")
    lstRec.SubItems(45) = IIf(Len(Trim(txtLOthDesc)), Trim(txtLOthDesc), "")
    lstRec.SubItems(46) = IIf(Len(Trim(txtIOthDesc)), Trim(txtIOthDesc), "")
    lstRec.SubItems(47) = IIf(Len(Trim(txtMedType1(0))), Trim(txtMedType1(0)), "")
    lstRec.SubItems(48) = IIf(Len(Trim(txtMedType1(1))), Trim(txtMedType1(1)), "")
    lstRec.SubItems(49) = IIf(Len(Trim(txtMedType1(2))), Trim(txtMedType1(2)), "")
    lstRec.SubItems(50) = IIf(Len(Trim(txtMedType1(3))), Trim(txtMedType1(3)), "")
    lstRec.SubItems(51) = IIf(Len(Trim(txtMedType1(4))), Trim(txtMedType1(4)), "")
    lstRec.SubItems(52) = IIf(Len(Trim(txtMedDay1(0))), Trim(txtMedDay1(0)), "")
    lstRec.SubItems(53) = IIf(Len(Trim(txtMedDay1(1))), Trim(txtMedDay1(1)), "")
    lstRec.SubItems(54) = IIf(Len(Trim(txtMedDay1(2))), Trim(txtMedDay1(2)), "")
    lstRec.SubItems(55) = IIf(Len(Trim(txtMedDay1(3))), Trim(txtMedDay1(3)), "")
    lstRec.SubItems(56) = IIf(Len(Trim(txtMedDay1(4))), Trim(txtMedDay1(4)), "")
    lstRec.SubItems(57) = IIf(Len(Trim(txtMedQty1(0))), Trim(txtMedQty1(0)), "")
    lstRec.SubItems(58) = IIf(Len(Trim(txtMedQty1(1))), Trim(txtMedQty1(1)), "")
    lstRec.SubItems(59) = IIf(Len(Trim(txtMedQty1(2))), Trim(txtMedQty1(2)), "")
    lstRec.SubItems(60) = IIf(Len(Trim(txtMedQty1(3))), Trim(txtMedQty1(3)), "")
    lstRec.SubItems(61) = IIf(Len(Trim(txtMedQty1(4))), Trim(txtMedQty1(4)), "")
    lstRec.SubItems(62) = IIf(Len(Trim(txtDiagnosis1)), Trim(txtDiagnosis1), "")
    lstRec.SubItems(63) = IIf(Len(Trim(txtDiagnosis2)), Trim(txtDiagnosis2), "")
    lstRec.SubItems(64) = IIf(Len(Trim(txtDiagnosis3)), Trim(txtDiagnosis3), "")
    lstRec.SubItems(65) = IIf(Len(Trim(txtDiagnosis4)), Trim(txtDiagnosis4), "")
    lstRec.SubItems(66) = TxtPayor
    lstRec.SubItems(67) = txtPlanCode
    lstRec.SubItems(68) = TxtInsType
    lstRec.SubItems(69) = strHLTCode
    lstRec.SubItems(70) = TxtBalLimit
    lstRec.SubItems(71) = TxtVisitLimit
    lstRec.SubItems(72) = TxtCardType
    lstRec.SubItems(73) = strVisitIND
    lstRec.SubItems(74) = strAnnualInd
    lstRec.SubItems(75) = TxtCoPay
'    lstRec.SubItems(76) = TxtPolNo
    lstRec.SubItems(77) = strPatientIC
    lstRec.SubItems(78) = txtEffDate
    lstRec.SubItems(79) = txtExpDate
    lstRec.SubItems(80) = IIf(Len(strPayTo), strPayTo, "")
    lstRec.SubItems(81) = strPayee
    lstRec.SubItems(82) = SuppLmtType
    lstRec.SubItems(83) = IIf(IsNumeric(B1), B1, 0)
    lstRec.SubItems(84) = IIf(IsNumeric(B2), B2, 0)
    lstRec.SubItems(85) = IIf(IsNumeric(B3), B3, 0)
    lstRec.SubItems(86) = IIf(IsNumeric(B4), B4, 0)
    lstRec.SubItems(87) = IIf(IsNumeric(B5), B5, 0)
    lstRec.SubItems(88) = IIf(IsNumeric(B6), B6, 0)
    lstRec.SubItems(89) = IIf(IsNumeric(B7), B7, 0)
    lstRec.SubItems(90) = IIf(IsNumeric(B8), B8, 0)
    lstRec.SubItems(91) = IIf(IsNumeric(C1), C1, 0)
    lstRec.SubItems(92) = IIf(IsNumeric(C2), C2, 0)
    lstRec.SubItems(93) = IIf(Len(Trim(TxtDiagnosis5)), Trim(TxtDiagnosis5), "")
    lstRec.SubItems(94) = ReasonOtherIND
    If Trim(ReasonOtherIND) = "Y" Then
        lstRec.SubItems(95) = Trim(TxtDiagnosisOth)
    Else
        lstRec.SubItems(95) = ""
    End If
    lstRec.SubItems(96) = IIf(IsNumeric(A30), A30, 0)
    
    lstRec.SubItems(102) = IIf(Len(Trim(txtMedType1(5))), Trim(txtMedType1(5)), "")
    lstRec.SubItems(103) = IIf(Len(Trim(txtMedType1(6))), Trim(txtMedType1(6)), "")
    lstRec.SubItems(104) = IIf(Len(Trim(txtMedDay1(5))), Trim(txtMedDay1(5)), "")
    lstRec.SubItems(105) = IIf(Len(Trim(txtMedDay1(6))), Trim(txtMedDay1(6)), "")
    lstRec.SubItems(106) = IIf(Len(Trim(txtMedQty1(5))), Trim(txtMedQty1(5)), "")
    lstRec.SubItems(107) = IIf(Len(Trim(txtMedQty1(6))), Trim(txtMedQty1(6)), "")
    lstRec.SubItems(110) = IIf(IsDate(txtInvRecDate), txtInvRecDate, "")
    lstRec.SubItems(111) = Trim(Left(CboInvRecMode, IIf(InStr(1, CboInvRecMode, "-") - 1 < 0, Len(CboInvRecMode), InStr(1, CboInvRecMode, "-") - 1)))
    lstRec.SubItems(112) = IIf(Len(Trim(TxtRemark)), Trim(TxtRemark), "")
    lstRec.SubItems(113) = IIf(IsNumeric(txtMBMAdjAnnLmt), txtMBMAdjAnnLmt, 0)
    lstRec.SubItems(114) = IIf(IsNumeric(A31), A31, 0)
    lstRec.SubItems(469) = IIf(Len(cboFFS), Mid(cboFFS, 1, 1), "")
    lstRec.SubItems(470) = IIf(Len(cboProcMth), cboProcMth, "")
    
    If InStr(1, A29, "-") Then
         A29 = IIf(IsNumeric(A29), A29, 0)
         A29 = A29 * (-1)
    Else
           A29 = IIf(IsNumeric(A29), A29, 0)
    End If
    
    lstRec.SubItems(471) = IIf(IsNumeric(A29), A29, 0)
    lstRec.SubItems(472) = IIf(IsNumeric(txtCLNBillAmt), txtCLNBillAmt, 0)
    lstRec.SubItems(473) = IIf(IsNumeric(A32), A32, 0)
    lstRec.SubItems(474) = IIf(IsNumeric(A33), A33, 0)
    lstRec.SubItems(475) = IIf(Len(tmpGRPBordxFees), tmpGRPBordxFees, "")
    lstRec.SubItems(479) = IIf(Len(txtCASRemark), txtCASRemark, "")
    lstRec.SubItems(480) = IIf(IsNumeric(txtLmtExceeded), txtLmtExceeded, 0)
    
    lstRec.SubItems(482) = IIf(Len(Trim(txtMedQty1(7))), Trim(txtMedQty1(7)), 0)
    lstRec.SubItems(483) = IIf(Len(Trim(txtMedQty1(8))), Trim(txtMedQty1(8)), 0)
    lstRec.SubItems(484) = IIf(Len(Trim(txtMedQty1(9))), Trim(txtMedQty1(9)), 0)
    lstRec.SubItems(485) = IIf(Len(Trim(txtMedQty1(10))), Trim(txtMedQty1(10)), 0)
    lstRec.SubItems(486) = IIf(Len(Trim(txtMedQty1(11))), Trim(txtMedQty1(11)), 0)
    lstRec.SubItems(487) = IIf(Len(Trim(txtMedQty1(12))), Trim(txtMedQty1(12)), 0)
    lstRec.SubItems(488) = IIf(Len(Trim(txtMedQty1(13))), Trim(txtMedQty1(13)), 0)
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub DiagDesc(DIagCode As String, tmpDiagDesc As String)
Const METHOD_NAME = "DiagDesc"
On Error GoTo debugs
Dim DiagRec As New ADODB.Recordset
Dim StrDiagSql As String

    If Trim(DIagCode) <> "" Then
        StrDiagSql = " SELECT ReasonCode, ReasonDesc " & _
                 " from TreatmentReason where 0=0 " & _
                 " and ReasonCode='" & Trim(DIagCode) & "'"
        DiagRec.Open StrDiagSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        
        Do While Not DiagRec.EOF
            tmpDiagDesc = DiagRec!ReasonDesc
            DiagRec.MoveNext
        Loop
        
        DiagRec.Close
        Set DiagRec = Nothing
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
Private Sub DisplayBillAmt()
Dim tmpBillAmt As Double
    tmpBillAmt = getTotal
    lblDispBillAmt = tmpBillAmt
    lblDispCLNBillAmt = tmpBillAmt
    
      If tmpGrpCompany = "C014" Then   ' Group company for Prince Court
        If Len(Trim(C1)) Then
            lblActCoPay = 5
            A29 = 5
        Else
            lblActCoPay = 0
            A29 = 0
        End If
    End If
    'Call A29_Change
    End Sub

Private Sub cboMemoClaim_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboMemoClaim.Text, cboMemoClaim.SelStart) + Chr(KeyAscii) + Mid(cboMemoClaim.Text, cboMemoClaim.SelStart + cboMemoClaim.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboMemoClaim.ListCount - 1
       
        If NewText = LCase(Left(cboMemoClaim.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = cboMemoClaim.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            cboMemoClaim.Text = ValidValue
            cboMemoClaim.SelStart = 0
            cboMemoClaim.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub PaymentInfo()
Const METHOD_NAME = "PaymentInfo"
On Error GoTo debugs
Dim strsql As String
Dim PayRec As New ADODB.Recordset
Dim PayList As ListItem
    
    
    strsql = "SELECT PVCLNNo, PVCLNDate,  CLNPaidAmt, PVCLNChequeAmt, " & _
             "CLNRemark From Payment2CLN WHERE CASNumber='" & Trim(txtCaseNo) & "'"
    PayRec.Open strsql, MediConnCISPayment, adOpenForwardOnly, adLockOptimistic
    Do While Not PayRec.EOF
        tmpPymtStatus = True
        Label22 = "True"
        PayRec.MoveNext
    Loop
    PayRec.Close
    Set PayRec = Nothing
    
    strsql = "SELECT PVNo, PVDate, MBMPaidAmt, PVPayto, PVCLNChequeAmt, " & _
             "Remark From Payment2MBM WHERE CASNumber='" & Trim(txtCaseNo) & "'"
    PayRec.Open strsql, MediConnCISPayment, adOpenForwardOnly, adLockOptimistic
    Do While Not PayRec.EOF
        tmpPymtStatus = True
        Label22 = "True"
    PayRec.MoveNext
    Loop
    PayRec.Close
    Set PayRec = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CheckDoubleEntry()
Const METHOD_NAME = "CheckDoubleEntry"
On Error GoTo debugs

    Dim cmdChk As New ADODB.Command
    Dim PrmMBMNo As New ADODB.Parameter
    Dim PrmCoverID As New ADODB.Parameter
    Dim PrmClnCode As New ADODB.Parameter
    Dim PrmInvNo As New ADODB.Parameter
    Dim PrmVstTime As New ADODB.Parameter
    Dim PrmVstDate As New ADODB.Parameter
    Dim PrmTotalActAmt As New ADODB.Parameter
    Dim PrmCaseNo As New ADODB.Parameter
    Dim rsChk As New ADODB.Recordset
    
    Set cmdChk.ActiveConnection = MediConnCISDB
    cmdChk.CommandText = "sp_chkDoubleCaseEntry"
    cmdChk.CommandType = adCmdStoredProc
    cmdChk.CommandTimeout = 300
    
    Set PrmMBMNo = cmdChk.CreateParameter("MBMNumber", adVarChar, adParamInput, 50, Trim(txtTabMBMNo))
    cmdChk.Parameters.Append PrmMBMNo
    Set PrmCoverID = cmdChk.CreateParameter("CoverID", adVarChar, adParamInput, 2, Trim(txtTabCoverID))
    cmdChk.Parameters.Append PrmCoverID
    Set PrmClnCode = cmdChk.CreateParameter("ClnCode", adVarChar, adParamInput, 10, Trim(txtClinicCode))
    cmdChk.Parameters.Append PrmClnCode
    Set PrmInvNo = cmdChk.CreateParameter("InvoiceNo", adVarChar, adParamInput, 200, Trim(txtInvNo))
    cmdChk.Parameters.Append PrmInvNo
    Set PrmVstTime = cmdChk.CreateParameter("VisitTime", adVarChar, adParamInput, 20, Trim(txtVisitTime))
    cmdChk.Parameters.Append PrmVstTime
    Set PrmVstDate = cmdChk.CreateParameter("DateVisit", adDate, adParamInput, , Trim(txtVisitDate))
    cmdChk.Parameters.Append PrmVstDate
    Set PrmTotalActAmt = cmdChk.CreateParameter("TotalActualAmt", adCurrency, adParamInput, , Trim(txtCLNBillAmt))
    cmdChk.Parameters.Append PrmTotalActAmt
    Set PrmCaseNo = cmdChk.CreateParameter("CasNumber", adVarChar, adParamInput, 100, Trim(txtCaseNo))
    cmdChk.Parameters.Append PrmCaseNo
    
    rsChk.CursorLocation = adUseClient
    rsChk.CursorType = adOpenForwardOnly
    rsChk.CacheSize = 100
    Set rsChk = cmdChk.Execute
    Do While Not rsChk.EOF
        strChkDoubleEntry = rsChk!isExist
        rsChk.MoveNext
    Loop
    rsChk.Close
    Set rsChk = Nothing
    
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Function CheckExceedLimit() As Boolean
Dim FinalApprovAmt As Double
Dim isPrinDepGPShared As Boolean

    FinalApprovAmt = lstInvoice.ListItems(1).SubItems(161)
    
    Dim strnetCLNBillAmt As String
    
    If IsNumeric(txtAnnLmt) = True Then
        If ASOInd <> "Y" Then
            Dim cmdChk As New ADODB.Command
            Dim PrmAppend As New ADODB.Parameter
            Dim rsChk As New ADODB.Recordset
            Dim rsInd As New ADODB.Recordset
            Dim strAppend, strMaternityInd, strDentalInd, strMBMno, strCoverID As String
            Dim strDBAmt, strDBSpecAmt As String
            Dim dbMaternity, dbDental, dbOtherDed, dbThyroid As Double
            Dim sArr() As String
            
            strDBAmt = 0
            strDBSpecAmt = 0
            
            rsInd.Open "SELECT Top 1 PLNMaternityInd, PDentalStatus, COALESCE(IsPrinDepGPShared,0) AS IsPrinDepGPShared FROM PLN WHERE PLNCode='" & Trim(strClnPln) & "'", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            If rsInd.EOF = False Then
                strMaternityInd = IIf(Len(rsInd!PLNMaternityInd), rsInd!PLNMaternityInd, "")
                strDentalInd = IIf(Len(rsInd!PDentalStatus), rsInd!PDentalStatus, "")
                isPrinDepGPShared = rsInd!isPrinDepGPShared
            End If
            rsInd.Close
            Set rsInd = Nothing
                    
            If Len(txtMBMNo) Then
               If InStr(1, txtMBMNo, "-") Then
                  sArr = Split(txtMBMNo, "-")
                  strMBMno = sArr(0)
                  strCoverID = sArr(1)
               Else
                  strMBMno = txtMBMNo
                  strCoverID = "00"
               End If
            End If
        
            strAppend = ""
            If isPrinDepGPShared Then
                strAppend = strAppend & "  CISCas.MBMnumber= '" & Trim(strMBMno) & "' AND CISCas.CASNumber <> '" & Trim(txtCaseNo) & "' "
            Else
                If strPlanCategory = "A" Then
                    strAppend = strAppend & "  CISCas.MBMnumber= '" & Trim(strMBMno) & "' AND CISCas.CASNumber <> '" & Trim(txtCaseNo) & "' "
                ElseIf strPlanCategory = "B" Then
                    If strCoverID = "00" Then
                        strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(strMBMno) & "' AND CoverID = '00' AND CISCas.CASNumber <> '" & Trim(txtCaseNo) & "' "
                    Else
                        strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(strMBMno) & "' AND CoverID <> '00' AND CISCas.CASNumber <> '" & Trim(txtCaseNo) & "' "
                    End If
                ElseIf strPlanCategory = "C" Or strPlanCategory = "" Then
                    strAppend = strAppend & "  CISCas.MBMnumber= '" & Trim(strMBMno) & "' AND CoverID ='" + strCoverID + "' AND CISCas.CASNumber <> '" & Trim(txtCaseNo) & "' "
                ElseIf strPlanCategory = "D" Then
                    If strCoverID = "00" Then
                        strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(strMBMno) & "' AND CoverID = '00' AND CISCas.CASNumber <> '" & Trim(txtCaseNo) & "' "
                    Else
                        strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(strMBMno) & "' AND CoverID <> '00' AND CISCas.CASNumber <> '" & Trim(txtCaseNo) & "' "
                    End If
                End If
            End If
                
            Set cmdChk.ActiveConnection = MediConnCISDB
            cmdChk.CommandText = "CIS_FFS2013"
            cmdChk.CommandType = adCmdStoredProc
            cmdChk.CommandTimeout = 300
            Set PrmAppend = cmdChk.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
            cmdChk.Parameters.Append PrmAppend
            rsChk.CursorLocation = adUseClient
            rsChk.CursorType = adOpenForwardOnly
            rsChk.CacheSize = 100
            Set rsChk = cmdChk.Execute
            
            If Trim(txtSpecBal) = "N/A" Then
                Do While Not rsChk.EOF
                    If IsNumeric(rsChk!CLMMaternity) Then
                        dbMaternity = CDbl(rsChk!CLMMaternity)
                    Else
                        dbMaternity = 0
                    End If
                           
                    If IsNumeric(rsChk!CLMDental) Then
                        dbDental = CDbl(rsChk!CLMDental)
                    Else
                        dbDental = 0
                    End If
                          
                    If IsNumeric(rsChk!CLMOthersNoDed) Then
                        dbOtherDed = CDbl(rsChk!CLMOthersNoDed)
                    Else
                        dbOtherDed = 0
                    End If
                           
                    If IsNumeric(rsChk!CLMTyphiadJabNoDed) Then
                        dbThyroid = CDbl(rsChk!CLMTyphiadJabNoDed)
                    Else
                        dbThyroid = 0
                    End If
                         
                    If strMaternityInd = "Y" And strDentalInd = "Y" Then
                        strDBAmt = CDbl(strDBAmt) + IIf(IsNumeric(rsChk!CLMBordxAmt), CDbl(rsChk!CLMBordxAmt) - (dbMaternity + dbOtherDed + dbThyroid + dbDental), 0)
                    ElseIf strMaternityInd = "Y" And strDentalInd <> "Y" Then
                        strDBAmt = CDbl(strDBAmt) + IIf(IsNumeric(rsChk!CLMBordxAmt), CDbl(rsChk!CLMBordxAmt) - (dbMaternity + dbOtherDed + dbThyroid), 0)
                    ElseIf strMaternityInd <> "Y" And strDentalInd = "Y" Then
                        strDBAmt = CDbl(strDBAmt) + IIf(IsNumeric(rsChk!CLMBordxAmt), CDbl(rsChk!CLMBordxAmt) - (dbOtherDed + dbThyroid + dbDental), 0)
                    ElseIf strMaternityInd <> "Y" And strDentalInd <> "Y" Then
                        strDBAmt = CDbl(strDBAmt) + IIf(IsNumeric(rsChk!CLMBordxAmt), CDbl(rsChk!CLMBordxAmt) - (dbOtherDed + dbThyroid), 0)
                    End If
                rsChk.MoveNext
                Loop
            Else
                Do While Not rsChk.EOF
                    If IsNumeric(rsChk!CLMMaternity) Then
                            dbMaternity = CDbl(rsChk!CLMMaternity)
                    Else
                        dbMaternity = 0
                    End If
                      
                    If IsNumeric(rsChk!CLMDental) Then
                        dbDental = CDbl(rsChk!CLMDental)
                    Else
                        dbDental = 0
                    End If
                                    
                    If IsNumeric(rsChk!CLMOthersNoDed) Then
                        dbOtherDed = CDbl(rsChk!CLMOthersNoDed)
                    Else
                        dbOtherDed = 0
                    End If
                    
                    If IsNumeric(rsChk!CLMTyphiadJabNoDed) Then
                        dbThyroid = CDbl(rsChk!CLMTyphiadJabNoDed)
                    Else
                        dbThyroid = 0
                    End If
                                
                    If rsChk!PLNGPSPShared = "Y" Then
                        If strMaternityInd = "Y" And strDentalInd = "Y" Then
                            strDBAmt = CDbl(strDBAmt) + IIf(IsNumeric(rsChk!CLMBordxAmt), CDbl(rsChk!CLMBordxAmt) - (dbMaternity + dbOtherDed + dbThyroid + dbDental), 0)
                        ElseIf strMaternityInd = "Y" And strDentalInd <> "Y" Then
                            strDBAmt = CDbl(strDBAmt) + IIf(IsNumeric(rsChk!CLMBordxAmt), CDbl(rsChk!CLMBordxAmt) - (dbMaternity + dbOtherDed + dbThyroid), 0)
                        ElseIf strMaternityInd <> "Y" And strDentalInd = "Y" Then
                            strDBAmt = CDbl(strDBAmt) + IIf(IsNumeric(rsChk!CLMBordxAmt), CDbl(rsChk!CLMBordxAmt) - (dbOtherDed + dbThyroid + dbDental), 0)
                        ElseIf strMaternityInd <> "Y" And strDentalInd <> "Y" Then
                            strDBAmt = CDbl(strDBAmt) + IIf(IsNumeric(rsChk!CLMBordxAmt), CDbl(rsChk!CLMBordxAmt) - (dbOtherDed + dbThyroid), 0)
                        End If
                    Else
                        If rsChk!CLNType <> "S" Then
                            If strMaternityInd = "Y" And strDentalInd = "Y" Then
                                strDBAmt = CDbl(strDBAmt) + IIf(IsNumeric(rsChk!CLMBordxAmt), CDbl(rsChk!CLMBordxAmt) - (dbMaternity + dbOtherDed + dbThyroid + dbDental), 0)
                            ElseIf strMaternityInd = "Y" And strDentalInd <> "Y" Then
                                strDBAmt = CDbl(strDBAmt) + IIf(IsNumeric(rsChk!CLMBordxAmt), CDbl(rsChk!CLMBordxAmt) - (dbMaternity + dbOtherDed + dbThyroid), 0)
                            ElseIf strMaternityInd <> "Y" And strDentalInd = "Y" Then
                                strDBAmt = CDbl(strDBAmt) + IIf(IsNumeric(rsChk!CLMBordxAmt), CDbl(rsChk!CLMBordxAmt) - (dbOtherDed + dbThyroid + dbDental), 0)
                            ElseIf strMaternityInd <> "Y" And strDentalInd <> "Y" Then
                                strDBAmt = CDbl(strDBAmt) + IIf(IsNumeric(rsChk!CLMBordxAmt), CDbl(rsChk!CLMBordxAmt) - (dbOtherDed + dbThyroid), 0)
                            End If
                        ElseIf rsChk!CLNType = "S" Then
                            Dim bAccLmt As Boolean
                            bAccLmt = False
                            If strPlanCategory = "A" Then
                                bAccLmt = True
                            ElseIf (strPlanCategory = "B" And strCoverID = "00" And rsChk!coverid = "00") Or (strPlanCategory = "B" And strCoverID <> "00" And rsChk!coverid <> "00") Then
                                bAccLmt = True
                            ElseIf (strPlanCategory = "C" Or strPlanCategory = "") And rsChk!coverid = strCoverID Then
                                bAccLmt = True
                            ElseIf (strPlanCategory = "D" And strCoverID = "00" And rsChk!coverid = "00") Or (strPlanCategory = "D" And strCoverID <> "00" And rsChk!coverid <> "00") Then
                                bAccLmt = True
                            End If
                            
                            If bAccLmt = True Then
                                If strMaternityInd = "Y" And strDentalInd = "Y" Then
                                    strDBSpecAmt = CDbl(strDBSpecAmt) + IIf(IsNumeric(rsChk!CLMBordxAmt), CDbl(rsChk!CLMBordxAmt) - (dbMaternity + dbOtherDed + dbThyroid + dbDental), 0)
                                ElseIf strMaternityInd = "Y" And strDentalInd <> "Y" Then
                                    strDBSpecAmt = CDbl(strDBSpecAmt) + IIf(IsNumeric(rsChk!CLMBordxAmt), CDbl(rsChk!CLMBordxAmt) - (dbMaternity + dbOtherDed + dbThyroid), 0)
                                ElseIf strMaternityInd <> "Y" And strDentalInd = "Y" Then
                                    strDBSpecAmt = CDbl(strDBSpecAmt) + IIf(IsNumeric(rsChk!CLMBordxAmt), CDbl(rsChk!CLMBordxAmt) - (dbOtherDed + dbThyroid + dbDental), 0)
                                ElseIf strMaternityInd <> "Y" And strDentalInd <> "Y" Then
                                    strDBSpecAmt = CDbl(strDBSpecAmt) + IIf(IsNumeric(rsChk!CLMBordxAmt), CDbl(rsChk!CLMBordxAmt) - (dbOtherDed + dbThyroid), 0)
                                End If
                            End If
                        End If
                    End If
                rsChk.MoveNext
                Loop
            End If
        
            rsChk.Close
            Set rsChk = Nothing
            Set cmdChk = Nothing
            Set PrmAppend = Nothing
            
            If strMaternityInd = "Y" And strDentalInd = "Y" Then
                strnetCLNBillAmt = CDbl(txtCLNBillAmt) - (CDbl(A7) + CDbl(A8) + CDbl(A32) + CDbl(A33))
            ElseIf strMaternityInd = "Y" And strDentalInd <> "Y" Then
                strnetCLNBillAmt = CDbl(txtCLNBillAmt) - (CDbl(A8) + CDbl(A32) + CDbl(A33))
            ElseIf strMaternityInd <> "Y" And strDentalInd = "Y" Then
                strnetCLNBillAmt = CDbl(txtCLNBillAmt) - (CDbl(A7) + CDbl(A32) + CDbl(A33))
            ElseIf strMaternityInd <> "Y" And strDentalInd <> "Y" Then
                strnetCLNBillAmt = CDbl(txtCLNBillAmt) - (CDbl(A32) + CDbl(A33))
            End If
            
            If Trim(lblClnType) = "S" And IsNumeric(txtSpecBal) Then
                'check against SP Annual Limit
                If (CDbl(strDBSpecAmt) + CDbl(strnetCLNBillAmt)) <= txtSpecAccLmt Then
                        CheckExceedLimit = False
                    Else
                        CheckExceedLimit = True
                End If
            Else
                'check against GP Annual Limit
                If ((CDbl(strDBAmt) + CDbl(strnetCLNBillAmt)) - CDbl(tmpTopupAmt)) <= txtAnnLmt Then
                    CheckExceedLimit = False
                Else
                    CheckExceedLimit = True
                End If
            End If
        Else
            strnetCLNBillAmt = CDbl(txtCLNBillAmt) - (CDbl(A32) + CDbl(A33))
            
            If ((CDbl(TxtBalLimit) + FinalApprovAmt) - CDbl(strnetCLNBillAmt)) <= 0 Then
                CheckExceedLimit = True
            Else
                CheckExceedLimit = False
            End If
        End If
    Else
        CheckExceedLimit = False
    End If

End Function

