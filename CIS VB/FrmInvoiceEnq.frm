VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmInvoiceEnq 
   Caption         =   "Invoice Enquiry"
   ClientHeight    =   7992
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7992
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   5115
      Left            =   0
      TabIndex        =   83
      Top             =   2280
      Width           =   11745
      _ExtentX        =   20722
      _ExtentY        =   9017
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Invoice Info"
      TabPicture(0)   =   "FrmInvoiceEnq.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(48)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameMed"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrameDiag"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "TxtRemark"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FrameInv"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Charges"
      TabPicture(1)   =   "FrmInvoiceEnq.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameDet"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Mgmt Note"
      TabPicture(2)   =   "FrmInvoiceEnq.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtCASRemark"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Others"
      TabPicture(3)   =   "FrmInvoiceEnq.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lstReceipt"
      Tab(3).Control(1)=   "lstPayment"
      Tab(3).Control(2)=   "lstMedixPay"
      Tab(3).Control(3)=   "Label2(4)"
      Tab(3).Control(4)=   "Label2(3)"
      Tab(3).Control(5)=   "Label2(2)"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Claim Documents "
      TabPicture(4)   =   "FrmInvoiceEnq.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lstOtherScanDocs"
      Tab(4).Control(1)=   "List3"
      Tab(4).Control(2)=   "List2"
      Tab(4).Control(3)=   "List1"
      Tab(4).Control(4)=   "Label2(10)"
      Tab(4).Control(5)=   "Label2(9)"
      Tab(4).Control(6)=   "Label2(8)"
      Tab(4).Control(7)=   "Label2(6)"
      Tab(4).ControlCount=   8
      TabCaption(5)   =   "Pharmacy"
      TabPicture(5)   =   "FrmInvoiceEnq.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lvwMedicineDet"
      Tab(5).Control(1)=   "cmdPrint"
      Tab(5).Control(2)=   "lvwMedRequest"
      Tab(5).Control(3)=   "Label4"
      Tab(5).Control(4)=   "Label3"
      Tab(5).ControlCount=   5
      Begin MSComctlLib.ListView lvwMedicineDet 
         Height          =   2415
         Left            =   -74880
         TabIndex        =   295
         Top             =   2280
         Visible         =   0   'False
         Width           =   11535
         _ExtentX        =   20341
         _ExtentY        =   4255
         View            =   3
         LabelEdit       =   1
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Medicine"
            Object.Width           =   14641
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Quantity"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Price"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Export to Excel"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   294
         Top             =   4680
         Width           =   1335
      End
      Begin MSComctlLib.ListView lvwMedRequest 
         Height          =   1215
         Left            =   -74880
         TabIndex        =   292
         Top             =   720
         Visible         =   0   'False
         Width           =   11535
         _ExtentX        =   20341
         _ExtentY        =   2138
         View            =   3
         LabelEdit       =   1
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
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.ListBox lstOtherScanDocs 
         Height          =   1584
         ItemData        =   "FrmInvoiceEnq.frx":00A8
         Left            =   -70200
         List            =   "FrmInvoiceEnq.frx":00AA
         TabIndex        =   291
         ToolTipText     =   "right click to randomize contents"
         Top             =   2820
         Width           =   4455
      End
      Begin VB.ListBox List3 
         Height          =   1584
         ItemData        =   "FrmInvoiceEnq.frx":00AC
         Left            =   -74880
         List            =   "FrmInvoiceEnq.frx":00AE
         TabIndex        =   289
         ToolTipText     =   "right click to randomize contents"
         Top             =   2820
         Width           =   4455
      End
      Begin VB.ListBox List2 
         Height          =   1584
         ItemData        =   "FrmInvoiceEnq.frx":00B0
         Left            =   -70200
         List            =   "FrmInvoiceEnq.frx":00B2
         TabIndex        =   285
         ToolTipText     =   "right click to randomize contents"
         Top             =   900
         Width           =   4455
      End
      Begin VB.ListBox List1 
         Height          =   1584
         ItemData        =   "FrmInvoiceEnq.frx":00B4
         Left            =   -74880
         List            =   "FrmInvoiceEnq.frx":00B6
         TabIndex        =   284
         ToolTipText     =   "right click to randomize contents"
         Top             =   900
         Width           =   4455
      End
      Begin VB.TextBox txtCASRemark 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4545
         Left            =   -74880
         MaxLength       =   3800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   274
         Top             =   480
         Width           =   11415
      End
      Begin VB.Frame FrameInv 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   261
         Top             =   620
         Width           =   11445
         Begin VB.TextBox txtMC 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5400
            MaxLength       =   3
            TabIndex        =   265
            Top             =   170
            Width           =   1095
         End
         Begin VB.TextBox txtBillNo 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   795
            MaxLength       =   20
            TabIndex        =   264
            Top             =   170
            Width           =   1815
         End
         Begin VB.ComboBox cboFFS 
            Height          =   315
            Left            =   5400
            TabIndex        =   263
            Top             =   450
            Width           =   1125
         End
         Begin VB.TextBox txtCLNBillAmt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   795
            MaxLength       =   20
            TabIndex        =   262
            Top             =   450
            Width           =   1815
         End
         Begin MSMask.MaskEdBox txtVisitDate 
            Height          =   315
            Left            =   3480
            TabIndex        =   266
            Top             =   165
            Width           =   1215
            _ExtentX        =   2138
            _ExtentY        =   550
            _Version        =   393216
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
         Begin MSMask.MaskEdBox txtVisitTime 
            Height          =   315
            Left            =   3480
            TabIndex        =   267
            Top             =   450
            Width           =   1215
            _ExtentX        =   2138
            _ExtentY        =   550
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Caption         =   "Bill No"
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
            Index           =   77
            Left            =   120
            TabIndex        =   273
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Vst Date"
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
            Index           =   76
            Left            =   2760
            TabIndex        =   272
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Vst Time"
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
            Index           =   75
            Left            =   2700
            TabIndex        =   271
            Top             =   540
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "MC"
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
            Index           =   74
            Left            =   5040
            TabIndex        =   270
            Top             =   225
            Width           =   375
         End
         Begin VB.Label Label101 
            Caption         =   "FFS"
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
            Index           =   9
            Left            =   5040
            TabIndex        =   269
            Top             =   555
            Width           =   735
         End
         Begin VB.Label Label101 
            Caption         =   "Bill Amt"
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
            Index           =   15
            Left            =   120
            TabIndex        =   268
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.TextBox TxtRemark 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   120
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   259
         Top             =   3780
         Width           =   5085
      End
      Begin VB.Frame FrameDiag 
         Height          =   1995
         Left            =   120
         TabIndex        =   241
         Top             =   1500
         Width           =   5055
         Begin VB.TextBox txtDiagnosis1 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   252
            Top             =   400
            Width           =   855
         End
         Begin VB.TextBox txtDiagnosis2 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   251
            Top             =   680
            Width           =   855
         End
         Begin VB.TextBox txtDiagnosis3 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   250
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtDiagnosis4 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   249
            Top             =   1240
            Width           =   855
         End
         Begin VB.TextBox TxtDiagnosis5 
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   248
            Top             =   1520
            Width           =   855
         End
         Begin VB.TextBox TxtDiagDesc1 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1125
            MaxLength       =   100
            TabIndex        =   247
            Top             =   400
            Width           =   3855
         End
         Begin VB.TextBox TxtDiagDesc2 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1125
            MaxLength       =   100
            TabIndex        =   246
            Top             =   680
            Width           =   3855
         End
         Begin VB.TextBox TxtDiagDesc3 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1125
            MaxLength       =   100
            TabIndex        =   245
            Top             =   960
            Width           =   3855
         End
         Begin VB.TextBox TxtDiagDesc4 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1125
            MaxLength       =   100
            TabIndex        =   244
            Top             =   1240
            Width           =   3855
         End
         Begin VB.TextBox TxtDiagnosisOth 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1125
            MaxLength       =   100
            TabIndex        =   243
            Top             =   1520
            Width           =   3855
         End
         Begin VB.TextBox TxtDiagDesc5 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1000
            MaxLength       =   10
            TabIndex        =   242
            Top             =   3250
            Width           =   2950
         End
         Begin VB.Label Label1 
            Caption         =   "Diagnosis"
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
            Index           =   49
            Left            =   240
            TabIndex        =   258
            Top             =   165
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   69
            Left            =   120
            TabIndex        =   257
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   70
            Left            =   120
            TabIndex        =   256
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   71
            Left            =   120
            TabIndex        =   255
            Top             =   1020
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   72
            Left            =   120
            TabIndex        =   254
            Top             =   1300
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   73
            Left            =   120
            TabIndex        =   253
            Top             =   1575
            Width           =   255
         End
      End
      Begin VB.Frame FrameMed 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   5280
         TabIndex        =   209
         Top             =   1500
         Width           =   6255
         Begin VB.TextBox txtMedDay1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   4680
            MaxLength       =   3
            TabIndex        =   230
            Top             =   405
            Width           =   735
         End
         Begin VB.TextBox txtMedQty1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
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
            TabIndex        =   229
            Top             =   405
            Width           =   735
         End
         Begin VB.TextBox txtMedType1 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
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
            TabIndex        =   228
            Top             =   405
            Width           =   4455
         End
         Begin VB.TextBox txtMedType1 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   227
            Top             =   680
            Width           =   4455
         End
         Begin VB.TextBox txtMedType1 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   240
            TabIndex        =   226
            Top             =   960
            Width           =   4455
         End
         Begin VB.TextBox txtMedType1 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   240
            TabIndex        =   225
            Top             =   1215
            Width           =   4455
         End
         Begin VB.TextBox txtMedType1 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   240
            TabIndex        =   224
            Top             =   1455
            Width           =   4455
         End
         Begin VB.TextBox txtMedDay1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   4680
            MaxLength       =   3
            TabIndex        =   223
            Top             =   690
            Width           =   735
         End
         Begin VB.TextBox txtMedQty1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
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
            TabIndex        =   222
            Top             =   690
            Width           =   735
         End
         Begin VB.TextBox txtMedDay1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   4680
            MaxLength       =   3
            TabIndex        =   221
            Top             =   945
            Width           =   735
         End
         Begin VB.TextBox txtMedQty1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
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
            TabIndex        =   220
            Top             =   945
            Width           =   735
         End
         Begin VB.TextBox txtMedDay1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   4680
            MaxLength       =   3
            TabIndex        =   219
            Top             =   1230
            Width           =   735
         End
         Begin VB.TextBox txtMedQty1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
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
            TabIndex        =   218
            Top             =   1230
            Width           =   735
         End
         Begin VB.TextBox txtMedDay1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   4680
            MaxLength       =   3
            TabIndex        =   217
            Top             =   1485
            Width           =   735
         End
         Begin VB.TextBox txtMedQty1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
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
            TabIndex        =   216
            Top             =   1485
            Width           =   735
         End
         Begin VB.TextBox txtMedType1 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   240
            TabIndex        =   215
            Top             =   1740
            Width           =   4455
         End
         Begin VB.TextBox txtMedDay1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   4680
            MaxLength       =   3
            TabIndex        =   214
            Top             =   1740
            Width           =   735
         End
         Begin VB.TextBox txtMedQty1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
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
            TabIndex        =   213
            Top             =   1740
            Width           =   735
         End
         Begin VB.TextBox txtMedType1 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   240
            TabIndex        =   212
            Top             =   2030
            Width           =   4455
         End
         Begin VB.TextBox txtMedDay1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   4680
            MaxLength       =   3
            TabIndex        =   211
            Top             =   2030
            Width           =   735
         End
         Begin VB.TextBox txtMedQty1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
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
            TabIndex        =   210
            Top             =   2030
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Medication && Injection"
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
            Index           =   107
            Left            =   240
            TabIndex        =   240
            Top             =   135
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   65
            Left            =   120
            TabIndex        =   239
            Top             =   1230
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   64
            Left            =   120
            TabIndex        =   238
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   63
            Left            =   120
            TabIndex        =   237
            Top             =   705
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   62
            Left            =   120
            TabIndex        =   236
            Top             =   405
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   66
            Left            =   120
            TabIndex        =   235
            Top             =   1515
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Days"
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
            Index           =   50
            Left            =   4680
            TabIndex        =   234
            Top             =   135
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Qty"
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
            Index           =   51
            Left            =   5505
            TabIndex        =   233
            Top             =   135
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   67
            Left            =   120
            TabIndex        =   232
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "7"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   68
            Left            =   120
            TabIndex        =   231
            Top             =   2080
            Width           =   255
         End
      End
      Begin VB.Frame FrameDet 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   -74880
         TabIndex        =   101
         Top             =   360
         Width           =   11490
         Begin VB.TextBox txtLmtExceeded 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9705
            MaxLength       =   8
            TabIndex        =   275
            Top             =   2280
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox A32 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7005
            MaxLength       =   8
            TabIndex        =   164
            Top             =   3825
            Width           =   1095
         End
         Begin VB.Frame Frame1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1975
            Left            =   8280
            TabIndex        =   148
            Top             =   240
            Width           =   3135
            Begin VB.TextBox A29 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   880
               MaxLength       =   8
               TabIndex        =   149
               Top             =   690
               Width           =   975
            End
            Begin VB.Label Label101 
               Caption         =   "Actual"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   14
               Left            =   1680
               TabIndex        =   163
               Top             =   120
               Width           =   615
            End
            Begin VB.Label Label101 
               Caption         =   "Clinic"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   13
               Left            =   890
               TabIndex        =   162
               Top             =   120
               Width           =   615
            End
            Begin VB.Label lblActCoPay 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   1920
               TabIndex        =   161
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label lblDispCLNBillAmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   885
               TabIndex        =   160
               Top             =   420
               Width           =   975
            End
            Begin VB.Label lblCLNBillAmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   840
               TabIndex        =   159
               Top             =   1005
               Width           =   975
            End
            Begin VB.Line Line1 
               Index           =   0
               X1              =   1920
               X2              =   3000
               Y1              =   1560
               Y2              =   1560
            End
            Begin VB.Label lblPay2CLN 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   1920
               TabIndex        =   158
               Top             =   1605
               Width           =   1095
            End
            Begin VB.Label Label76 
               Caption         =   "Amt Pay"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
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
               TabIndex        =   157
               Top             =   1620
               Width           =   1095
            End
            Begin VB.Label lblFFS 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   1920
               TabIndex        =   156
               Top             =   1275
               Width           =   1095
            End
            Begin VB.Label Label76 
               Caption         =   "Less:FFS"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
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
               TabIndex        =   155
               Top             =   1260
               Width           =   855
            End
            Begin VB.Label Label76 
               Caption         =   "Bill Amt"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
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
               TabIndex        =   154
               Top             =   975
               Width           =   735
            End
            Begin VB.Label lblBillAmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   1920
               TabIndex        =   153
               Top             =   1005
               Width           =   1095
            End
            Begin VB.Label Label101 
               Caption         =   "Co-Pymt"
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
               Index           =   0
               Left            =   75
               TabIndex        =   152
               Top             =   690
               Width           =   735
            End
            Begin VB.Label lblDispBillAmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   1920
               TabIndex        =   151
               Top             =   420
               Width           =   1095
            End
            Begin VB.Label Label76 
               Caption         =   "Total Bill "
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.4
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
               TabIndex        =   150
               Top             =   420
               Width           =   735
            End
         End
         Begin VB.TextBox txtPOthDesc 
            Height          =   285
            Left            =   2760
            MaxLength       =   50
            TabIndex        =   147
            Top             =   2460
            Width           =   975
         End
         Begin VB.TextBox B1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   146
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox B5 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   145
            Top             =   2160
            Width           =   1215
         End
         Begin VB.TextBox A10 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3765
            MaxLength       =   8
            TabIndex        =   144
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox A11 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3765
            MaxLength       =   8
            TabIndex        =   143
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox A12 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3765
            MaxLength       =   8
            TabIndex        =   142
            Top             =   855
            Width           =   1095
         End
         Begin VB.TextBox A13 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3765
            MaxLength       =   8
            TabIndex        =   141
            Top             =   1095
            Width           =   1095
         End
         Begin VB.TextBox A14 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3765
            MaxLength       =   8
            TabIndex        =   140
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox A15 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3765
            MaxLength       =   8
            TabIndex        =   139
            Top             =   1605
            Width           =   1095
         End
         Begin VB.TextBox A16 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3765
            MaxLength       =   8
            TabIndex        =   138
            Top             =   1890
            Width           =   1095
         End
         Begin VB.TextBox A17 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3765
            MaxLength       =   8
            TabIndex        =   137
            Top             =   2160
            Width           =   1095
         End
         Begin VB.TextBox A18 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
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
            Top             =   2400
            Width           =   1095
         End
         Begin VB.TextBox A19 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6960
            MaxLength       =   8
            TabIndex        =   135
            Top             =   420
            Width           =   1095
         End
         Begin VB.TextBox A20 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6960
            MaxLength       =   8
            TabIndex        =   134
            Top             =   645
            Width           =   1095
         End
         Begin VB.TextBox A21 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6960
            MaxLength       =   8
            TabIndex        =   133
            Top             =   870
            Width           =   1095
         End
         Begin VB.TextBox A22 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6960
            MaxLength       =   8
            TabIndex        =   132
            Top             =   1110
            Width           =   1095
         End
         Begin VB.TextBox A23 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6960
            MaxLength       =   8
            TabIndex        =   131
            Top             =   1380
            Width           =   1095
         End
         Begin VB.TextBox txtLOthDesc 
            Height          =   285
            Left            =   5880
            MaxLength       =   50
            TabIndex        =   130
            Top             =   1420
            Width           =   1095
         End
         Begin VB.TextBox A7 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3765
            MaxLength       =   8
            TabIndex        =   129
            Top             =   3120
            Width           =   1095
         End
         Begin VB.TextBox A8 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3765
            MaxLength       =   8
            TabIndex        =   128
            Top             =   3360
            Width           =   1095
         End
         Begin VB.TextBox A9 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3765
            MaxLength       =   8
            TabIndex        =   127
            Top             =   3600
            Width           =   1095
         End
         Begin VB.TextBox A30 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3765
            MaxLength       =   8
            TabIndex        =   126
            Top             =   3855
            Width           =   1095
         End
         Begin VB.TextBox A24 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6960
            MaxLength       =   8
            TabIndex        =   125
            Top             =   2115
            Width           =   1095
         End
         Begin VB.TextBox A25 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6960
            MaxLength       =   8
            TabIndex        =   124
            Top             =   2355
            Width           =   1095
         End
         Begin VB.TextBox A26 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6960
            MaxLength       =   8
            TabIndex        =   123
            Top             =   2625
            Width           =   1095
         End
         Begin VB.TextBox A27 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6960
            MaxLength       =   8
            TabIndex        =   122
            Top             =   2865
            Width           =   1095
         End
         Begin VB.TextBox A28 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6960
            MaxLength       =   8
            TabIndex        =   121
            Top             =   3120
            Width           =   1095
         End
         Begin VB.TextBox txtIOthDesc 
            Height          =   285
            Left            =   5880
            MaxLength       =   50
            TabIndex        =   120
            Top             =   3165
            Width           =   1095
         End
         Begin VB.TextBox A31 
            Alignment       =   1  'Right Justify
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
            Left            =   3765
            MaxLength       =   12
            TabIndex        =   119
            ToolTipText     =   "54"
            Top             =   4080
            Width           =   1095
         End
         Begin VB.TextBox txtMBMAdjAnnLmt 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
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
            Left            =   3765
            MaxLength       =   12
            TabIndex        =   118
            ToolTipText     =   "55"
            Top             =   4320
            Width           =   1095
         End
         Begin VB.TextBox B8 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   117
            Top             =   2400
            Width           =   1215
         End
         Begin VB.TextBox B7 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   116
            Top             =   2640
            Width           =   1215
         End
         Begin VB.TextBox B6 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   115
            Top             =   2880
            Width           =   1215
         End
         Begin VB.TextBox C2 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   114
            Top             =   3120
            Width           =   1215
         End
         Begin VB.TextBox A6 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   113
            Top             =   3120
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox A5 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   112
            Top             =   3120
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox A4 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   111
            Top             =   3120
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox B4 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   110
            Top             =   645
            Width           =   1215
         End
         Begin VB.TextBox B3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   109
            Top             =   920
            Width           =   1215
         End
         Begin VB.TextBox B2 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   108
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox C1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   107
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox A2 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   106
            Top             =   1440
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox A1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   105
            Top             =   1440
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox A3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   104
            Top             =   1440
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtInVIndex 
            Height          =   285
            Left            =   3960
            TabIndex        =   103
            Text            =   "txtInvIndex"
            Top             =   3720
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox A33 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7005
            MaxLength       =   8
            TabIndex        =   102
            Top             =   4065
            Width           =   1095
         End
         Begin VB.Label Label85 
            Caption         =   "Limit Exceeded"
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
            Index           =   3
            Left            =   8280
            TabIndex        =   276
            Top             =   2310
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label Label90 
            Caption         =   "Typhoid Jab"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   5400
            TabIndex        =   208
            Top             =   3885
            Width           =   915
         End
         Begin VB.Label txtOther 
            Caption         =   "Others"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   5400
            TabIndex        =   207
            Top             =   4125
            Width           =   1095
         End
         Begin VB.Label Label73 
            Caption         =   "No Limit Deduction"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   5400
            TabIndex        =   206
            ToolTipText     =   "Geberal Pratitioner"
            Top             =   3600
            Width           =   1875
         End
         Begin VB.Label Label1 
            Caption         =   "Oth"
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
            Index           =   18
            Left            =   2400
            TabIndex        =   205
            Top             =   2475
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   110
            Left            =   120
            TabIndex        =   204
            ToolTipText     =   "Geberal Pratitioner"
            Top             =   3160
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   109
            Left            =   120
            TabIndex        =   203
            ToolTipText     =   "Geberal Pratitioner"
            Top             =   1480
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Specialist"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   108
            Left            =   120
            TabIndex        =   202
            ToolTipText     =   "Geberal Pratitioner"
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Med."
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
            Index           =   21
            Left            =   120
            TabIndex        =   201
            Top             =   2400
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Inj."
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
            Index           =   23
            Left            =   120
            TabIndex        =   200
            Top             =   2640
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Med."
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
            Index           =   20
            Left            =   120
            TabIndex        =   199
            Top             =   690
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Inj."
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
            Index           =   22
            Left            =   120
            TabIndex        =   198
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Fee + LT Med."
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
            Index           =   8
            Left            =   120
            TabIndex        =   197
            Top             =   2160
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "LT Med."
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
            Index           =   25
            Left            =   120
            TabIndex        =   196
            Top             =   2920
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Fee + LT Med."
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
            Index           =   7
            Left            =   120
            TabIndex        =   195
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "LT Med."
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
            Index           =   24
            Left            =   120
            TabIndex        =   194
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "DTP && Polio"
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
            Index           =   39
            Left            =   5400
            TabIndex        =   193
            Top             =   2145
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Others"
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
            Index           =   32
            Left            =   5400
            TabIndex        =   192
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Pap Smear"
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
            Index           =   31
            Left            =   5400
            TabIndex        =   191
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Urine"
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
            Index           =   30
            Left            =   5400
            TabIndex        =   190
            Top             =   825
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "B. Glucose"
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
            Index           =   29
            Left            =   5400
            TabIndex        =   189
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Laboratory"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   27
            Left            =   5400
            TabIndex        =   188
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "B. Drawing"
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
            Index           =   28
            Left            =   5400
            TabIndex        =   187
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "M. Surg."
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   17
            Left            =   2400
            TabIndex        =   186
            Top             =   2115
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "I&&D"
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
            Index           =   15
            Left            =   2400
            TabIndex        =   185
            Top             =   1635
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "T&&S/Sto"
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
            Index           =   16
            Left            =   2400
            TabIndex        =   184
            Top             =   1875
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Nebulizer"
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
            Index           =   14
            Left            =   2400
            TabIndex        =   183
            Top             =   1395
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Dressing"
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
            Index           =   13
            Left            =   2400
            TabIndex        =   182
            Top             =   1155
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "ECG"
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
            Index           =   12
            Left            =   2400
            TabIndex        =   181
            Top             =   915
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Utr Sound"
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
            Index           =   11
            Left            =   2400
            TabIndex        =   180
            Top             =   645
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "MMR"
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
            Index           =   41
            Left            =   5400
            TabIndex        =   179
            Top             =   2910
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Hepatitis B"
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
            Index           =   42
            Left            =   5400
            TabIndex        =   178
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "DTP&&Polio&&HIB"
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
            Index           =   40
            Left            =   5400
            TabIndex        =   177
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Immunization"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   106
            Left            =   5400
            TabIndex        =   176
            Top             =   1890
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Others"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   43
            Left            =   5400
            TabIndex        =   175
            Top             =   3135
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Dental"
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
            Index           =   35
            Left            =   2400
            TabIndex        =   174
            Top             =   3105
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Maternity"
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
            Index           =   36
            Left            =   2400
            TabIndex        =   173
            Top             =   3360
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "PRV Chk."
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
            Index           =   37
            Left            =   2400
            TabIndex        =   172
            Top             =   3600
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "GP"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   105
            Left            =   120
            TabIndex        =   171
            ToolTipText     =   "Geberal Pratitioner"
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "X-Ray "
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
            Index           =   10
            Left            =   2400
            TabIndex        =   170
            Top             =   375
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Procedures"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   2400
            TabIndex        =   169
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Vision"
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
            Index           =   38
            Left            =   2400
            TabIndex        =   168
            Top             =   3840
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Non GP Amt"
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
            Index           =   46
            Left            =   2400
            TabIndex        =   167
            Top             =   4080
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Adj Ann Lmt"
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
            Index           =   47
            Left            =   2400
            TabIndex        =   166
            Top             =   4320
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Others"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   34
            Left            =   2400
            TabIndex        =   165
            Top             =   2880
            Width           =   1455
         End
      End
      Begin MSComctlLib.ListView lstReceipt 
         Height          =   1305
         Left            =   -68760
         TabIndex        =   95
         Top             =   1020
         Width           =   5055
         _ExtentX        =   8911
         _ExtentY        =   2307
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
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Index"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Receipt No"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Receipt Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Receipt Type"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView lstPayment 
         Height          =   1305
         Left            =   -74880
         TabIndex        =   96
         Top             =   1020
         Width           =   5655
         _ExtentX        =   9970
         _ExtentY        =   2307
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
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Index"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PV No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "PV Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Pay To"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Amt Paid"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Remark"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Bank-In Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstMedixPay 
         Height          =   1275
         Left            =   -74880
         TabIndex        =   97
         Top             =   3060
         Width           =   5655
         _ExtentX        =   9970
         _ExtentY        =   2244
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
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Index"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Medix CaseNo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Medix PV No"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Medix PV Date"
            Object.Width           =   2293
         EndProperty
      End
      Begin VB.Label Label4 
         Caption         =   "Medicine Details"
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
         Left            =   -74880
         TabIndex        =   296
         Top             =   2040
         Width           =   4335
      End
      Begin VB.Label Label3 
         Caption         =   "Medication Request "
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
         Left            =   -74880
         TabIndex        =   293
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "OTHER SCAN Documents"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   -70200
         TabIndex        =   290
         Top             =   2580
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "REIMBURSEMENT Documents"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   -74880
         TabIndex        =   288
         Top             =   2580
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "INVOICE Documents"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   -70200
         TabIndex        =   287
         Top             =   660
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "GL Documents"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   -74880
         TabIndex        =   283
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Remark"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   48
         Left            =   120
         TabIndex        =   260
         Top             =   3540
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "UHC Payment to Clinic / Member"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -74880
         TabIndex        =   100
         Top             =   780
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Receipt from Client"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -68760
         TabIndex        =   99
         Top             =   780
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "Insurance Claims Paid By Medix"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   98
         Top             =   2820
         Width           =   3135
      End
   End
   Begin VB.Frame FrameCAS 
      Height          =   615
      Left            =   5790
      TabIndex        =   73
      Top             =   280
      Width           =   6080
      Begin VB.ComboBox cboStatus 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   600
         TabIndex        =   5
         Text            =   "O - Open"
         Top             =   207
         Width           =   1095
      End
      Begin VB.ComboBox cboClaimability 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "FrmInvoiceEnq.frx":00B8
         Left            =   2760
         List            =   "FrmInvoiceEnq.frx":00BA
         TabIndex        =   6
         Text            =   "A - Accepted"
         Top             =   200
         Width           =   1455
      End
      Begin VB.ComboBox cboPending 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
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
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Status"
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
         Index           =   60
         Left            =   120
         TabIndex        =   76
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "Claimability"
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
         Index           =   53
         Left            =   1800
         TabIndex        =   75
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Pending"
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
         Index           =   52
         Left            =   4320
         TabIndex        =   74
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   30
      TabIndex        =   71
      Top             =   280
      Width           =   5745
      Begin VB.TextBox txtCaseNo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
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
      Begin VB.CommandButton cmdGo 
         Caption         =   "&Go"
         Default         =   -1  'True
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
         Left            =   3360
         TabIndex        =   1
         Top             =   200
         Width           =   540
      End
      Begin VB.CommandButton CmdPrev 
         Caption         =   "Pre&v"
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
         Left            =   3840
         TabIndex        =   2
         Top             =   200
         Width           =   660
      End
      Begin VB.CommandButton CmdNext 
         Caption         =   "&Next"
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
         Left            =   4440
         TabIndex        =   3
         Top             =   200
         Width           =   540
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
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
         Left            =   4980
         TabIndex        =   4
         Top             =   200
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "Case No."
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
         Index           =   1
         Left            =   120
         TabIndex        =   72
         Top             =   255
         Width           =   735
      End
   End
   Begin VB.Frame FrameMBM 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   30
      TabIndex        =   42
      Top             =   840
      Width           =   8085
      Begin VB.TextBox txtClinicName 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1360
         TabIndex        =   9
         Top             =   180
         Width           =   2610
      End
      Begin VB.TextBox txtTabCoverID 
         Height          =   285
         Left            =   3120
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox TxtPayor 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4680
         TabIndex        =   48
         Top             =   1560
         Visible         =   0   'False
         Width           =   975
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
         Left            =   1680
         TabIndex        =   47
         Top             =   1560
         Width           =   495
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
         Left            =   1680
         TabIndex        =   46
         Top             =   1800
         Width           =   495
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
         Left            =   4680
         TabIndex        =   45
         Top             =   1800
         Width           =   975
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
         Left            =   4680
         TabIndex        =   44
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtMBMStatus 
         Height          =   285
         Left            =   2520
         TabIndex        =   43
         Top             =   1560
         Width           =   375
      End
      Begin MSMask.MaskEdBox txtExpDate 
         Height          =   315
         Left            =   1680
         TabIndex        =   50
         Top             =   2280
         Width           =   975
         _ExtentX        =   1715
         _ExtentY        =   550
         _Version        =   393216
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
      Begin MSMask.MaskEdBox txtEffDate 
         Height          =   285
         Left            =   1680
         TabIndex        =   51
         Top             =   2040
         Width           =   975
         _ExtentX        =   1715
         _ExtentY        =   508
         _Version        =   393216
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
      Begin VB.TextBox txtClinicCode 
         Height          =   330
         Left            =   795
         TabIndex        =   8
         Top             =   180
         Width           =   580
      End
      Begin VB.TextBox txtNPClinic 
         Height          =   330
         Left            =   1360
         TabIndex        =   52
         Top             =   180
         Visible         =   0   'False
         Width           =   2610
      End
      Begin VB.ComboBox CboPayto 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   795
         TabIndex        =   10
         Text            =   "C - Clinic"
         Top             =   480
         Width           =   1380
      End
      Begin VB.TextBox txtInvNo 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   795
         MaxLength       =   30
         TabIndex        =   12
         Top             =   760
         Width           =   1380
      End
      Begin VB.TextBox txtMBMNo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4800
         TabIndex        =   16
         Top             =   180
         Width           =   3195
      End
      Begin VB.TextBox txtPatientIC 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4800
         TabIndex        =   17
         Top             =   435
         Width           =   3195
      End
      Begin VB.TextBox txtMBMName 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4800
         TabIndex        =   18
         Top             =   705
         Width           =   3195
      End
      Begin MSMask.MaskEdBox txtInvDate 
         Height          =   315
         Left            =   795
         TabIndex        =   14
         Top             =   1005
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   550
         _Version        =   393216
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
      Begin MSMask.MaskEdBox txtInvRecDate 
         Height          =   315
         Left            =   2955
         TabIndex        =   11
         Top             =   480
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   550
         _Version        =   393216
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
      Begin VB.ComboBox CboInvRecMode 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2955
         TabIndex        =   13
         Top             =   720
         Width           =   1020
      End
      Begin VB.ComboBox cboProcMth 
         Height          =   315
         Left            =   2955
         TabIndex        =   15
         Top             =   1005
         Width           =   1020
      End
      Begin VB.TextBox txtDept 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4800
         TabIndex        =   91
         Top             =   990
         Width           =   3195
      End
      Begin VB.Label Label2 
         Caption         =   "Dept"
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
         Index           =   5
         Left            =   4080
         TabIndex        =   92
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label101 
         Caption         =   "Mth"
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
         Index           =   5
         Left            =   2220
         TabIndex        =   90
         ToolTipText     =   "Processing Month"
         Top             =   1110
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "InvNo"
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
         Index           =   4
         Left            =   120
         TabIndex        =   70
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "InvDate"
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
         Index           =   104
         Left            =   120
         TabIndex        =   69
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
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
         Index           =   61
         Left            =   4080
         TabIndex        =   68
         Top             =   795
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Clinic"
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
         Index           =   2
         Left            =   120
         TabIndex        =   67
         Top             =   260
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "MBM No"
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
         Index           =   6
         Left            =   4080
         TabIndex        =   66
         Top             =   255
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Pay To"
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
         Index           =   3
         Left            =   120
         TabIndex        =   65
         Top             =   560
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "RecDate"
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
         Index           =   103
         Left            =   2220
         TabIndex        =   64
         Top             =   560
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "RecMode"
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
         Index           =   5
         Left            =   2220
         TabIndex        =   63
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblPatientIC 
         Caption         =   "IC No."
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
         Left            =   4080
         TabIndex        =   62
         Top             =   525
         Width           =   615
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
         Left            =   3720
         TabIndex        =   61
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label62 
         Caption         =   "InsType"
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
         Left            =   3720
         TabIndex        =   60
         Top             =   2040
         Width           =   615
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
         Left            =   3720
         TabIndex        =   59
         Top             =   1560
         Visible         =   0   'False
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
         Left            =   720
         TabIndex        =   58
         Top             =   1560
         Width           =   975
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
         Left            =   720
         TabIndex        =   57
         Top             =   2280
         Width           =   855
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
         Left            =   720
         TabIndex        =   56
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   720
         TabIndex        =   55
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label102 
         Caption         =   "Status"
         Height          =   255
         Left            =   2160
         TabIndex        =   54
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label lblClnType 
         Caption         =   "lblClnType"
         Height          =   255
         Left            =   3240
         TabIndex        =   53
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame FrameDis 
      Enabled         =   0   'False
      Height          =   1395
      Left            =   8160
      TabIndex        =   28
      Top             =   840
      Width           =   3645
      Begin VB.TextBox txtSpAnnLmt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   281
         Top             =   120
         Width           =   780
      End
      Begin VB.TextBox txtSpAccAmt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   279
         Top             =   360
         Width           =   780
      End
      Begin VB.TextBox txtspLmt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   277
         Top             =   600
         Width           =   780
      End
      Begin VB.PictureBox PSignature 
         Height          =   285
         Left            =   2160
         ScaleHeight     =   240
         ScaleWidth      =   384
         TabIndex        =   29
         Top             =   1680
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtPolStatus 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   19
         Top             =   160
         Width           =   780
      End
      Begin VB.TextBox TxtCardType 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   20
         Top             =   1860
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.TextBox txtAnnLmt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   21
         Top             =   405
         Width           =   780
      End
      Begin VB.TextBox txtAccumAmt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   22
         Top             =   660
         Width           =   780
      End
      Begin VB.TextBox TxtBalLimit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   23
         Top             =   885
         Width           =   780
      End
      Begin VB.TextBox TxtVisitLimit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   24
         Top             =   885
         Width           =   780
      End
      Begin VB.TextBox txtDentStatus 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   30
         Top             =   680
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.TextBox txtLTM 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   31
         Top             =   1710
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.TextBox TxtCoPay 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   32
         Top             =   1770
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "Sp Ann Lmt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   33
         Left            =   1920
         TabIndex        =   282
         Top             =   195
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "SP Acc Amt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   26
         Left            =   1920
         TabIndex        =   280
         Top             =   435
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "SP Bal. Lmt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   19
         Left            =   1920
         TabIndex        =   278
         Top             =   675
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Ann Lmt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   56
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Acc. Amt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   57
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label34 
         BackColor       =   &H80000013&
         Caption         =   "Ann. Lmt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Vst Lmt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   59
         Left            =   1920
         TabIndex        =   36
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Card Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   55
         Left            =   2280
         TabIndex        =   34
         Top             =   1920
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
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
         Index           =   54
         Left            =   120
         TabIndex        =   33
         Top             =   200
         Width           =   735
      End
      Begin VB.Label Label69 
         BackColor       =   &H80000013&
         Caption         =   "LTM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   135
         Left            =   1920
         TabIndex        =   39
         Top             =   1800
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Bal. Lmt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   58
         Left            =   120
         TabIndex        =   35
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label112 
         Caption         =   "Co-Pay"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1080
         TabIndex        =   40
         Top             =   1845
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame Frame7 
      Height          =   555
      Left            =   0
      TabIndex        =   77
      Top             =   7440
      Width           =   11775
      Begin VB.CommandButton cmdExclusions 
         Caption         =   "&Exclusions"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8400
         TabIndex        =   93
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   11040
         TabIndex        =   27
         Top             =   160
         Width           =   615
      End
      Begin VB.CommandButton CmdClearAll 
         Caption         =   "&Clear All"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10200
         TabIndex        =   26
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdMBMInfo 
         Caption         =   "&MBM Info"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9240
         TabIndex        =   25
         Top             =   160
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Bordx ?"
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
         Index           =   78
         Left            =   2040
         TabIndex        =   89
         Top             =   195
         Width           =   615
      End
      Begin VB.Label lblBordx 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   88
         Top             =   195
         Width           =   480
      End
      Begin VB.Label lblBordxDate 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
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
         Left            =   6840
         TabIndex        =   87
         Top             =   195
         Width           =   1080
      End
      Begin VB.Label lblBordxNo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
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
         Left            =   4560
         TabIndex        =   86
         Top             =   195
         Width           =   1080
      End
      Begin VB.Label Label2 
         Caption         =   "Bordx. Date"
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
         Index           =   1
         Left            =   5880
         TabIndex        =   85
         Top             =   195
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Bordx. Ref. No."
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
         Index           =   0
         Left            =   3360
         TabIndex        =   84
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Batch No."
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
         Index           =   79
         Left            =   120
         TabIndex        =   79
         Top             =   195
         Width           =   735
      End
      Begin VB.Label LblBatchNo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
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
         Left            =   840
         TabIndex        =   78
         Top             =   195
         Width           =   960
      End
      Begin VB.Label lblGrpCompany 
         Caption         =   "Label10"
         Height          =   255
         Left            =   6960
         TabIndex        =   94
         Top             =   195
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lstInvoice 
      Height          =   105
      Left            =   0
      TabIndex        =   80
      Top             =   7680
      Visible         =   0   'False
      Width           =   11895
      _ExtentX        =   20976
      _ExtentY        =   191
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
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   473
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
         Text            =   "CLMActOthers"
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
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(241) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   240
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
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(284) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   283
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
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(327) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   326
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
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(370) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   369
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
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(413) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   412
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
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(456) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   455
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
   End
   Begin MSComctlLib.ListView lstErr 
      Height          =   255
      Left            =   0
      TabIndex        =   81
      Top             =   7560
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   445
      _ExtentY        =   445
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
   Begin VB.Label Label2 
      Caption         =   "GL Documents"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   4800
      TabIndex        =   286
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Label lblInvoice 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Invoice Enquiry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   82
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "FrmInvoiceEnq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bDataStatus As Boolean
Dim bExit As Boolean
Dim IntIndex As String
Dim ScanCode As String
Dim AA, CC As Integer
Dim strPayTo, strPayee As String
Dim strMBMno As String
Dim IntInvAmt As Double
Dim CASENumber, NewCaseNo As String
Dim MembershipNo As String
Dim CovPersonID As String
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
Dim strAccumAmt As String
'Dim StrInvIndex As String
Dim MBM As New ADODB.Recordset
Dim lstErrRec As ListItem
Dim VPlan, VExpDate, VEffDate, VDiagCode, VDiagIND As String
Dim SuppLmtType As String
Dim Z1, Z2, Z3, Z4, Z5, Z6, Z7, Z8 As Double
Dim CISCAS1 As New ADODB.Recordset
Dim ReasonOtherIND As String
Dim StrDentStatus As String
Dim tmpFinalApp As Double
Public lstRec As ListItem
Dim strGRPMCOStatus As String
Dim strPLNGrossStatus As String
Dim strMBMDepartment As String
Dim strMCStatus As String
Dim strMBMMCDaysBal As String
Dim strGRPPostageFees As String
Dim strGRPBordxFees As String
Dim strPLNMMPerc As String
Dim PictureFileName As String
Dim ASOInd As String
Dim UtilisedAmt As Double
Dim strMBMCancelDate  As String
Dim strCMBMCancelDate As String
Dim strCMBMCCLNExpDate As String
Dim BSNUser As Boolean
Private Const SW_RESTORE        As Long = &H9&
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias _
"SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long

Private Const LVM_FIRST = &H1000

Private Sub A1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub cmdExclusions_Click()
    If Trim(lblGrpCompany) = "" Then
        MsgBox "Please Check Group Company! ", vbOKOnly
    Else
        Load frmGeneralBenefitExclusion
        frmGeneralBenefitExclusion.cboGrpCompany = Trim(lblGrpCompany)
        frmGeneralBenefitExclusion.cmdSearch_Click
        'frmGeneralBenefitExclusion.Show vbModal
        
    End If
End Sub

Private Sub cmdPrint_Click()
    Call ExportToExcel(lvwMedRequest)
End Sub

Public Function ExportToExcel(lvw As MSComctlLib.ListView) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objWorksheet2 As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim objRange2 As Excel.Range
    Dim lngResults As Long
    Dim i As Integer
    Dim intCounter As Integer
    Dim intStartRow As Integer
    Dim strArray() As String
    Dim intVisibleColumns() As Integer
    Dim intColumns As Integer
    Dim itm As ListItem
    
    'If there are no selected items in the listview control
    
    If lvw.SelectedItem Is Nothing Then
        MsgBox "There aren't any items In the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    End If
    'Ask the user if they want to export
    lngResults = MsgBox("Do you want to export to Excel?", vbYesNo + vbQuestion, "Export Report")

    If lngResults = vbNo Then
        GoTo ExitFunction
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
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
    
    Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\Medication\MedicationRequest.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
            
    intCounter = 0
        
    ' Title 1
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
    
        ' A response of vbYes meant to export all the items
    
        If lngResults = vbYes Or itm.Selected Then
            ' increment the number of selected rows
            intCounter = intCounter + 1
        
            For i = 1 To intColumns
                If i = 1 Then
                    strArray(intCounter, 1) = itm.Text
                Else
                    strArray(intCounter, i) = Replace(itm.SubItems(intVisibleColumns(i) - 1), "~", vbCrLf)
                End If
            Next i
        End If
    Next itm
    End If

    ' Send entire array to Excel range
    With objWorksheet
        If lvw Is Nothing = False And intCounter <> 0 Then
            .Range(.Cells(5, 1), _
            .Cells(4 + intCounter, intColumns)) = strArray
        End If
    End With

    If lvw Is Nothing = False Then
        objWorksheet.Cells(2, "D") = Date
    End If
        
    objWorksheet.Columns.AutoFit
    
    intColumns = 0
    Set objWorksheet2 = objWorkbook.Sheets(2)
    Set objRange2 = objWorksheet2.Rows(1)
    objRange2.Font.Size = 10
    objRange2.Font.Bold = True

    If lvwMedicineDet Is Nothing = False Then
        For i = 1 To lvwMedicineDet.ColumnHeaders.Count
            If lvwMedicineDet.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange2.Cells(1, intColumns) = lvwMedicineDet.ColumnHeaders(i).Text
            End If
        Next i
        If lvwMedicineDet.ListItems.Count <> 0 Then
            ReDim strArray(1 To lvwMedicineDet.ListItems.Count, 1 To intColumns)
        End If
    End If
    intCounter = 0
    If lvwMedicineDet Is Nothing = False Then
    For Each itm In lvwMedicineDet.ListItems
        ' A response of vbYes meant to export all the items
        If lngResults = vbYes Or itm.Selected Then
            ' increment the number of selected rows
            intCounter = intCounter + 1
        
            For i = 1 To intColumns
                If i = 1 Then
                    strArray(intCounter, 1) = itm.Text
                Else
                    strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
                End If
            Next i
        End If
    Next itm
    End If
    With objWorksheet2
        If lvwMedicineDet Is Nothing = False And intCounter <> 0 Then
            .Range(.Cells(2, 1), _
            .Cells(1 + intCounter, intColumns)) = strArray
        End If
    End With
    objExcel.Interactive = True

ExportToExcel = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Set objRange = Nothing
Set objWorksheet = Nothing
Set objRange2 = Nothing
Set objWorksheet2 = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

Private Sub lblActCoPay_Change()
Dim a As Double
Dim B As Double

    a = IIf(IsNumeric(lblDispBillAmt), lblDispBillAmt, 0)
    B = IIf(IsNumeric(lblActCoPay), lblActCoPay, 0)
    If CDbl(a) <> "0" Then
        If a < 0 And B < 0 Then
            lblBillAmt = Format(Abs(CDbl(a)) - Abs(CDbl(B)), "-#,##0.00")
        
        Else
            lblBillAmt = Format(CDbl(a) - CDbl(B), "#,##0.00")
        End If
    End If
End Sub
Private Sub A29_Change()
Dim a As Double
Dim B As Double

    a = IIf(IsNumeric(lblDispCLNBillAmt), lblDispCLNBillAmt, 0)
    B = IIf(IsNumeric(A29), A29, 0)
    If CDbl(a) <> "0" Then
    If a < 0 And B < 0 Then
        lblCLNBillAmt = Format(Abs(CDbl(a)) - Abs(CDbl(B)), "-#,##0.00")
    Else
        lblCLNBillAmt = Format(CDbl(a) - CDbl(B), "#,##0.00")
    End If
    End If
    
End Sub

Private Sub A3_KeyPress(KeyAscii As Integer)
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
    CmdMBMInfo.Enabled = False
    txtNPClinic.Visible = False
    lstInvoice.ListItems.Clear
    lstErr.ListItems.Clear
End Sub

Private Sub CallLimit(tmpMBMTYpe As String, tmpAnnLmt As String, tmpVisitLmt As String, tmpSpecLmt As String, tmpLimitStatus As String)
    Const METHOD_NAME = "CallLimit"
    On Error GoTo debugs
    Dim ASOqry As String
    Dim ASORS As New ADODB.Recordset
    Dim ASOlimit As Double
    Dim IPLimit As Double
    Dim OPLimit As Double
    Dim IPUtilization As Double
    Dim OPUtilization As Double
    Dim IsCombined As String
    Dim sArr() As String
    Dim MBMNo As String
    Dim CovID As String
    
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
Private Sub CmdExit_Click()
    Unload Me
End Sub

Public Sub CmdGo_Click()
Const METHOD_NAME = "CmdGo_Click"
On Error GoTo debugs
Dim ZZ As String
Dim TotalSpecLimit As Double
Dim tmpMBMTYpe As String
Dim PictureFileName As String
Dim Strrelationtype As String
Dim MBMNo As String
Dim CovID As String
Dim sArr() As String
Dim dblMeternity As Double
Dim dblDental As Double
Dim OtherDed As Double
Dim thyroid As Double

Dim SpecLmtRec As New ADODB.Recordset
Dim tmpLimitStatus As String
Dim tmpAnnLmt As String
Dim tmpVisitLmt As String
Dim tmpSpecLmt As String
Dim tmpSpecAnnLmt As String
Dim strAppend As String
Dim tmpMaternityInd As String
Dim tmpSpecLimitStatus As String
tmpLimitStatus = SuppLmtType
Dim tmpSpecAccLmt As String
Dim tmpDentalInd As String
Dim isPrinDepGPShared As Boolean
   
    Call ClearEntries
    Call InvoiceINFO
    Call ClaimBordxInfo
    SSTab1.Tab = 0
    txtMBMNo = ChkStr(txtMBMNo)
    
    If Trim(txtMBMNo) <> "" Then
        ZZ = InStr(1, txtMBMNo, "-")
        MembershipNo = Trim(Left(txtMBMNo, IIf(ZZ - 1 <= 0, Len(Trim(txtMBMNo)), ZZ - 1)))
        If ZZ = 0 Then
            CovPersonID = "00"
        Else
            CovPersonID = Trim(Right(txtMBMNo, IIf(ZZ - 1 <= 0, 0, Len(Trim(txtMBMNo)) - ZZ)))
        End If
        Call MemberSuppInfo(MembershipNo, CovPersonID, strBalLimit, strVisitLimit, strAccumAmt, _
                            StrDentStatus, strCoPay, strPayor, strHLTCode, strClnPln, _
                            strClnIns, strEffdate, strExpDate, strPolNo, strMBMDataStatus, _
                            strPOlStatus, strPatientName, strPatientDOB, strPatientIC, _
                            strMBMCLNLTM, strCardType, strVisitIND, strAnnualInd, strSpecStatus, _
                            strMEPStatus, strDentLmt, strSpecLmt, strMEPLmt, SuppLmtType, _
                            StrDentLmtStatus, StrSpecLmtStatus, strBNFStatus, strDentExtLmt, str1stSpecLmt, _
                            strGRPMCOStatus, strPLNGrossStatus, strMBMDepartment, strMCStatus, strMBMMCDaysBal, strGRPPostageFees, _
                            strGRPBordxFees, strPLNMMPerc, Strrelationtype, strMBMCancelDate, strCMBMCancelDate, strCMBMCCLNExpDate)
        tmpMBMTYpe = "P"
        If CovPersonID <> "00" Then
            tmpMBMTYpe = "S"
        End If
        If Trim(MembershipNo) <> "" Then
            CmdMBMInfo.Enabled = True
            txtTabMBMNo = MembershipNo
            txtTabCoverID = CovPersonID
            txtMBMName = strPatientName
            TxtPayor = strPayor
            txtPlanCode = strClnPln
            TxtInsType = strClnIns
            txtPatientName = strPatientName
            txtEffDate = strEffdate
            txtExpDate = strExpDate
            txtPolStatus = strPOlStatus
            txtMBMStatus = strMBMDataStatus
            txtPatientIC = strPatientIC
            txtDept = strMBMDepartment
            'PictureFileName = LocSignPath & Trim(txtPatientIC) & ".jpg"
            'PSignature.Picture = LoadPicture()
            On Error Resume Next

            'PSignature.Picture = LoadPicture(PictureFileName)
            'If Err Then
            '    PSignature.Picture = LoadPicture()
            'End If
            
            txtPatientName = strPatientName
            TxtVisitLimit = Format(GetVisitLimit(strClnIns, strClnPln, strHLTCode, CDate(strEffdate), tmpMBMTYpe, strPayor), "#,##0.00")
            TxtBalLimit = txtAnnLmt
            If IsNumeric(txtAnnLmt) Then
                TxtBalLimit = Format(strBalLimit, "#,##0.00")
            End If
            If IsNumeric(TxtVisitLimit) Then
                TxtVisitLimit = Format(strVisitLimit, "#,##0.00")
            End If
            
            
              '------Calculate Limt Start-------
            
            txtAccumAmt = 0
            
            tmpSpecLimitStatus = ""
            SpecLmtRec.Open "SELECT Top 1 PLNMaternityInd,PDentalStatus,ASOInd, COALESCE(IsPrinDepGPShared,0) AS IsPrinDepGPShared FROM PLN WHERE PLNCode='" & Trim(strClnPln) & "'", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            If SpecLmtRec.EOF = False Then
                tmpMaternityInd = IIf(Len(SpecLmtRec!PLNMaternityInd), SpecLmtRec!PLNMaternityInd, "")
                tmpDentalInd = IIf(Len(SpecLmtRec!PDentalStatus), SpecLmtRec!PDentalStatus, "")
                ASOInd = IIf(Len(SpecLmtRec!ASOInd), SpecLmtRec!ASOInd, "N")
                isPrinDepGPShared = SpecLmtRec!isPrinDepGPShared
            End If
            SpecLmtRec.Close
            Set SpecLmtRec = Nothing
                
            
            Call CallLimit(tmpMBMTYpe, tmpAnnLmt, tmpVisitLmt, tmpSpecLmt, tmpLimitStatus)
            
            txtAnnLmt = tmpAnnLmt
            
            tmpSpecAnnLmt = tmpSpecLmt
            txtSpAnnLmt = tmpSpecLmt
            tmpSpecLimitStatus = ""
            SpecLmtRec.Open "SELECT SuppLimitStatus FROM SpecialistLimit WHERE PLNCode='" & Trim(strClnPln) & "'", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
            If SpecLmtRec.EOF = False Then
                tmpSpecLimitStatus = IIf(Len(SpecLmtRec!SuppLimitStatus), SpecLmtRec!SuppLimitStatus, "A")
            End If
            SpecLmtRec.Close
            Set SpecLmtRec = Nothing
            
            If Trim(tmpSpecAnnLmt) = "N/A" Then
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
                        strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "'"
                    ElseIf tmpLimitStatus = "B" Then
                        If CovID = "00" Then
                            strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID = '00' "
                        Else
                            strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID <> '00' "
                        End If
                    ElseIf tmpLimitStatus = "C" Or tmpLimitStatus = "" Then
                        strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID ='" + CovID + "' "
                    ElseIf tmpLimitStatus = "D" Then
                        If CovID = "00" Then
                            strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID = '00' "
                        Else
                            strAppend = strAppend & "  CISCas.MBMNumber = '" & Trim(MBMNo) & "' AND CoverID <> '00' "
                        End If
                    End If
                End If
             End If
                
               
                
                
                Dim cmd As New ADODB.Command
                Dim Prm1 As New ADODB.Parameter
                Dim rst2 As New ADODB.Recordset
                
                If Trim(strAppend) = "" Then
                  
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
                            If ASOInd = "Y" Then
                                txtAccumAmt = UtilisedAmt
                            End If
                        End If
                    End If
                End If
              rst2.MoveNext
                    Loop
                    rst2.Close
                    Set rst2 = Nothing
                    Set cmd = Nothing
                    Set Prm1 = Nothing
                End If
                
                tmpSpecAccLmt = Format(TotalSpecLimit, "#,###.00")
                If IsNumeric(tmpSpecLmt) Then
                    tmpSpecAnnLmt = Format(tmpSpecAnnLmt, "#,##0.00")
                     If IsNumeric(tmpSpecAccLmt) Then
                        txtspLmt = CDbl(tmpSpecAnnLmt) - CDbl(tmpSpecAccLmt)
                        txtSpAccAmt = CDbl(tmpSpecAccLmt)
                    End If
                Else
                    txtSpAccAmt = tmpSpecLmt
                    txtspLmt = tmpSpecLmt
                End If
            End If
            If IsNumeric(tmpAnnLmt) Then
                TxtBalLimit = Format(CDbl(tmpAnnLmt) - CDbl(txtAccumAmt), "#,##0.00")
            Else
            End If
            '------Calculate Limt End-------
            TxtCardType = strCardType
            txtLTM = strMBMCLNLTM
            txtDentStatus = StrDentStatus
        End If
        txtAccumAmt.Text = Format(txtAccumAmt.Text, "#,##0.00")
        txtSpAccAmt.Text = Format(txtSpAccAmt.Text, "#,##0.00")
        txtspLmt.Text = Format(txtspLmt.Text, "#,##0.00")
        Call AssignData2ListView
        Call WSInfo
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub CmdMBMInfo_Click()
    FrmINVMBMinfo.Source = 4
    FrmINVMBMinfo.Show vbModal
End Sub

Private Sub CmdNext_Click()
Const METHOD_NAME = "CmdNext_Click"
On Error GoTo debugs
Dim CAS As New ADODB.Recordset
Dim sql As String
    sql = " select Top 1 CASNumber from VIEW_Cas_MBM Where MBMCLNDataStatus <> 'D' " & _
          " AND MBMCLNStatus <> 'D' AND (HSCLNInd = 'C' OR HSCLNInd = 'M' OR HSCLNInd = 'B')" & _
          " AND CaseType='C' AND CASNumber > '" & RemStr(txtCaseNo) & "' Order by CASNumber "
          
    CAS.MaxRecords = 1
    CAS.Open sql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    If CAS.EOF = False Then
        txtCaseNo.Text = Trim(CAS!CasNumber)
        CAS.Close
        Set CAS = Nothing
        CmdGo_Click
    Else
        CAS.Close
        Set CAS = Nothing
        MsgBox " Last Record Reached! ", vbCritical + vbOKOnly
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
    sql = " select Top 1 CASNumber from VIEW_Cas_MBM Where MBMCLNDataStatus <> 'D' " & _
          " AND MBMCLNStatus <> 'D' AND (HSCLNInd = 'C' OR HSCLNInd = 'M' OR HSCLNInd = 'B')" & _
          " AND CaseType='C' AND CASNumber < '" & RemStr(txtCaseNo) & "' Order by CASNumber DESC"
    CAS.MaxRecords = 1
    CAS.Open sql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    
    If CAS.EOF = False Then
        txtCaseNo.Text = Trim(CAS!CasNumber)
        CAS.Close
        Set CAS = Nothing
        CmdGo_Click
    Else
        CAS.Close
        Set CAS = Nothing
        MsgBox "First Record Reached! " & vbNewLine & "Not a valid Case Number! ", vbCritical + vbOKOnly
    End If
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub cmdSearch_Click()
    frmCasList.Source = 1
    frmCasList.Show vbModal
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
    Call ActionLogUpdate(12, "FrmInvoiceEnq")
    bDataStatus = True
    bExit = False
    Call ClearEntries
    Call ComboListing
    CmdClearAll.Enabled = True
    IntIndex = 0
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
        
        ElseIf TypeOf ctl Is ListView Then
            ctl.ListItems.Clear
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
Dim tempCase As String
Dim tmpDiagDesc As String
Dim CISCAS1 As New ADODB.Recordset
Dim cmd1 As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
    lblGrpCompany = ""
    CASENumber = txtCaseNo
    Set cmd1.ActiveConnection = MediConnCISWS
    cmd1.CommandText = "SP_INV_CLAIM"
    cmd1.CommandType = adCmdStoredProc
    cmd1.CommandTimeout = 300
    Set Prm1 = cmd1.CreateParameter("StrSql", adVarChar, adParamInput, 20, CASENumber)
    cmd1.Parameters.Append Prm1
    CISCAS1.CursorLocation = adUseClient
    CISCAS1.CursorType = adOpenForwardOnly
    CISCAS1.CacheSize = 100
    Set CISCAS1 = cmd1.Execute
    
    If CISCAS1.EOF Then
        MsgBox "Not Record Found ! ", vbOKOnly + vbCritical, DialogBoxTitle
        tempCase = txtCaseNo
        lstInvoice.ListItems.Clear
        MembershipNo = ""
        CASENumber = ""
        txtCaseNo = tempCase
        Call ClearMBMInfo
        Call ClearEntries
        txtCaseNo.SetFocus
    Else
        'Do While Not CISCAS1.EOF
        With CISCAS1
            lblGrpCompany = IIf(Trim(Len(!MBMGroupCompany)), !MBMGroupCompany, "")
            txtMBMNo = !MBMNumber & "-" & !coverid
            MembershipNo = !MBMNumber
            CovPersonID = !coverid
            CASENumber = !CasNumber
            strPatientID = !coverid
            txtInvDate = IIf(Len(!InvoiceDate), Format(!InvoiceDate, "dd/mm/yyyy"), "__/__/____")
            txtInvNo = IIf(Len(!InvoiceNo), !InvoiceNo, "")
            txtInvRecDate = IIf(Len(!InvoiceReceiptDate), !InvoiceReceiptDate, "__/__/____")
            CboInvRecMode = IIf(Len(!InvoiceReceiptMode), !InvoiceReceiptMode, "")
            txtVisitDate = !DateVisit
            txtMC = IIf(Len(!MCDays), !MCDays, 0)
            lblDispBillAmt = IIf(IsNumeric(!INVAmount), Format(!INVAmount, "#,##0.00"), 0)
            lblDispCLNBillAmt = IIf(IsNumeric(!INVAmount), Format(!INVAmount, "#,##0.00"), 0)
            If Len(Trim(!TimeVisit)) Then
                txtVisitTime = Format(!TimeVisit, "hh:mm")
            End If
            txtBillNo = IIf(Len(Trim(!BillNo)), Trim(!BillNo), "")
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
            LblBatchNo = IIf(Len(!OPBatchNo), !OPBatchNo, "")
            strPayTo = !INVPayTo
            CboPayto.ListIndex = 0
            If Trim(!INVPayTo) = "M" Then
                CboPayto.ListIndex = 1
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
            txtCLNBillAmt = IIf(Len(!INVBillAmt), !INVBillAmt, 0)
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
            If Trim(!CASStatus) = "O" Then
                cboStatus.ListIndex = 0
            ElseIf Trim(!CASStatus) = "C" Then
                cboStatus.ListIndex = 1
            ElseIf Trim(!CASStatus) = "D" Then
                cboStatus.ListIndex = 2
            End If
            If Trim(!Claimability) = "A" Then
                cboClaimability.ListIndex = 0
            ElseIf Trim(!Claimability) = "R" Then
                cboClaimability.ListIndex = 1
            ElseIf Trim(!Claimability) = "C" Then
                cboClaimability.ListIndex = 2
            ElseIf Trim(!Claimability) = "S" Then
                cboClaimability.ListIndex = 3
            End If
            If Trim(!Pending) = "Y" Then
                cboPending.Text = "Y - Yes"
            ElseIf Trim(!Pending) = "N" Then
                cboPending.Text = "N - No"
            End If
            TxtRemark = IIf(Len(Trim(!CasRemark)), Trim(!CasRemark), "")
            txtCASRemark = IIf(Len(Trim(!CASMgmtNotes)), Trim(!CASMgmtNotes), "")
        End With
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
    txtPatientName = ""
    txtEffDate.mask = "##/##/####"
    txtExpDate.mask = "##/##/####"
    txtPolStatus = ""
    txtMBMStatus = ""
'    txtdob = ""
    txtPatientIC = ""
    txtPatientName = ""
    TxtBalLimit = ""
    TxtVisitLimit = ""
    TxtCardType = ""
    txtLTM = ""
End Sub
Private Sub ComboListing()
Const METHOD_NAME = "ComboListing"
On Error GoTo debugs
Dim InvRec As New ADODB.Recordset
        CboInvRecMode.Clear
        InvRec.Open "Select RecModeCode, RecModeDesc From InvRecMode ", MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        Do While Not InvRec.EOF
            CboInvRecMode.AddItem InvRec!RecModeCode & " - " & Trim(InvRec!RecModeDesc)
            InvRec.MoveNext
        Loop
        InvRec.Close
        Set InvRec = Nothing
        cboStatus.AddItem "O - Open"
        cboStatus.AddItem "C - Close"
        cboStatus.AddItem "D - Delete"
        cboClaimability.AddItem "A - Accepted"
        cboClaimability.AddItem "R - Repudiated"
        cboClaimability.AddItem "C - Cancelled"
        cboClaimability.AddItem "S - Special Approve"
        cboPending.AddItem "Y - Yes"
        cboPending.AddItem "N - No"
        CboPayto.AddItem "C - CLINIC"
        CboPayto.AddItem "M - MEMBER"
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
    If a < 0 And B < 0 Then
            lblBillAmt = Format(Abs(CDbl(a)) - Abs(CDbl(B)), "-#,##0.00")
    Else
            lblBillAmt = Format(CDbl(a) - CDbl(B), "#,##0.00")
    End If
End Sub

Private Sub lblDispCLNBillAmt_Change()
Dim a As Double
Dim B As Double

    a = IIf(IsNumeric(lblDispCLNBillAmt), lblDispCLNBillAmt, 0)
    B = IIf(IsNumeric(A29), A29, 0)
    If CDbl(a) <> "0" Then
        lblCLNBillAmt = Format(CDbl(a) - CDbl(B), "#,##0.00")
    End If

End Sub
Private Sub List1_Click()
On Error Resume Next
Dim lRet    As Long
Dim sFile   As String
    PictureFileName = DocPath & "Outpatient\" & "Guarantee Letter\" & txtCaseNo & ""
    crt PictureFileName & "\" & List1.List(List1.ListIndex)
    sFile = PictureFileName & "\" & List1.List(List1.ListIndex)
    lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
End Sub
Private Sub List2_Click()
On Error Resume Next
Dim lRet    As Long
Dim sFile   As String
    PictureFileName = DocPath & "Outpatient\" & "SP_ScannedDoc\" & txtCaseNo & ".pdf"
    sFile = PictureFileName
    lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
End Sub
Sub crt(pth As String)
Dim f
f = FreeFile
Open App.path + "\temp.html" For Output As #f
    Print #f, "<html><body><center><img src=""" & pth & """></center></body></html>"
Close #f
DoEvents
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

Sub addTheFiles()
Dim B As String
List1.Clear
PictureFileName = DocPath & "Outpatient\" & "Guarantee Letter\" & txtCaseNo & ""
B = Dir(PictureFileName & "\*.*")

If B <> "" Then
    If Right(B, 4) = ".jpg" Then
        List1.AddItem B
    End If
    If Right(B, 4) = ".tif" Then
        List1.AddItem B
    End If
    If Right(B, 4) = ".gif" Then
        List1.AddItem B
    End If
    If Right(B, 4) = ".JPG" Then
        List1.AddItem B
    End If
    If Right(B, 4) = ".TIF" Then
        List1.AddItem B
    End If
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
        If Right(B, 4) = ".gif" Then
            List1.AddItem B
        End If
        If Right(B, 4) = ".JPG" Then
            List1.AddItem B
        End If
        If Right(B, 4) = ".TIF" Then
            List1.AddItem B
        End If
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
End Sub
Sub addTheIntFiles()
Dim B As String
List2.Clear

'1. Create file name.
'2. Open folder
'3. check file exist in folder if True then add to list
Dim fname As String
Dim folderPath As String
Dim myFSO As Object
Dim file As String

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

folderPath = DocPath & "Outpatient\" & "OtherScannedDoc\" & txtCaseNo
Set myFSO = CreateObject("Scripting.FileSystemObject")

lstOtherScanDocs.Clear
If myFSO.FolderExists(folderPath) Then
    file = Dir(folderPath & "\*.*")
    Do While file > ""
      lstOtherScanDocs.AddItem file
      file = Dir()
    Loop
End If

'List1.Height = Me.Height - 2520 - 350
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

Private Sub lstOtherScanDocs_Click()
On Error Resume Next
Dim lRet    As Long
Dim sFile   As String
    PictureFileName = DocPath & "Outpatient\" & "OtherScannedDoc\" & txtCaseNo & ""
    sFile = PictureFileName & "\" & lstOtherScanDocs.List(lstOtherScanDocs.ListIndex)
    lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
End Sub

Private Sub lvwMedRequest_Click()
    If lvwMedRequest.ListItems.Count <> 0 Then
        If lvwMedRequest.SelectedItem.Text <> "" Then
            Call MedicineDetails
        End If
    End If
End Sub

Private Sub MedicineDetails()
Const METHOD_NAME = "MedicineDetails"
On Error GoTo debugs
Dim n As Integer
Dim i As Integer
Dim MedReqList As ListItem
Dim bHeaderCreated As Boolean
'Dim iFieldsCount As Integer

    bHeaderCreated = False
    lvwMedicineDet.ListItems.Clear
'    lvwMedicineDet.ColumnHeaders.Clear
    Dim MRcmd As New ADODB.Command
    Dim MRPrm1 As New ADODB.Parameter
    Dim MRRes As New ADODB.Recordset
        Set MRcmd.ActiveConnection = MediConnCISDB
        MRcmd.CommandText = "sp_CISMedReqDOMedicine"
        MRcmd.CommandType = adCmdStoredProc
        MRcmd.CommandTimeout = 300
        Set MRPrm1 = MRcmd.CreateParameter("CaseNo", adVarChar, adParamInput, 100, Trim(txtCaseNo.Text))
        MRcmd.Parameters.Append MRPrm1
        MRRes.CursorLocation = adUseClient
        MRRes.CursorType = adOpenForwardOnly
        MRRes.CacheSize = 100
        Set MRRes = MRcmd.Execute
        Do While Not MRRes.EOF
'            iFieldsCount = MRRes.Fields.count
'            If bHeaderCreated = False Then
'                For n = 0 To MRRes.Fields.count - 1
'                    lvwMedicineDet.ColumnHeaders.Add , , MRRes.Fields(n).Name
'                    bHeaderCreated = True
'                Next n
'            End If
            Set MedReqList = lvwMedicineDet.ListItems.Add(, , MRRes(0) & vbNullString)
            For i = 1 To MRRes.Fields.Count - 1
                If i = 2 Then
                    MedReqList.SubItems(i) = Format(MRRes(i), "#,##0.00")
                Else
                    MedReqList.SubItems(i) = MRRes(i) & vbNullString
                End If
            Next i
            MRRes.MoveNext
        Loop
        MRRes.Close
        Set MRRes = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Public Sub LV_AutoSizeColumn(LV As ListView, Optional Column _
 As ColumnHeader = Nothing)
 
 Dim C As ColumnHeader
 If Column Is Nothing Then
  For Each C In LV.ColumnHeaders
   SendMessage LV.hwnd, LVM_FIRST + 60, C.Index - 1, -1
  Next
 Else
  SendMessage LV.hwnd, LVM_FIRST + 60, Column.Index - 1, -1
 End If
 LV.Refresh
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Screen.MousePointer = vbHourglass
    If SSTab1.Tab = 3 Then
            Call PaymentInfo
            Call ReceiptInfo
    End If
    If SSTab1.Tab = 4 Then
        Call LoadDoc
        Call LoadIntDoc
        Call LoadWebDoc
    End If
    If SSTab1.Tab = 5 Then
        lvwMedRequest.Visible = False
        lvwMedicineDet.Visible = False
        Call MedicationRequestInfo
        Call MedicineDetails
        lvwMedRequest.Visible = True
        lvwMedicineDet.Visible = True
    End If
    Screen.MousePointer = vbDefault
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
    If Trim(lblClnType) = "S" Then
        B1.Enabled = False
        B2.Enabled = False
        B3.Enabled = False
        B4.Enabled = False
        B5.Enabled = True
        B6.Enabled = True
        B7.Enabled = True
        B8.Enabled = True
    End If
End Sub

Private Sub txtMBMNo_GotFocus()
    cmdGo.Default = True
End Sub

Private Sub txtMBMNo_LostFocus()
    cmdGo.Default = False
End Sub

Private Function getTotal() As Double
Dim ctl  As Control
Dim tmpBillAmt As Double
Dim tmpBalLmt As Double
Dim tmpAdjAmt As Double
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
        If Mid(ctl.Name, 1, 1) = "A" And Len(ctl.Name) <= 7 Then
            If IsNumeric(ctl.Text) = True Then
                IntInvAmt = IntInvAmt + CDbl(ctl.Text)
            End If
        End If
    Next
    IntInvAmt = Format(IntInvAmt, "#,##0.00")
    getTotal = IntInvAmt
End Function

Private Sub AssignData2ListView()
Dim strPayee As String
    Set lstRec = lstInvoice.ListItems.Add(, , "")
    If strPayTo = "C" Then
        'strPayee = Trim(Left(txtclinicname, IIf(InStr(1, txtclinicname, "-") - 1 < 0, Len(txtclinicname), InStr(1, txtclinicname, "-") - 1)))
        strPayee = txtClinicCode
    Else
        strPayee = "MEMBER"
    End If
       
    Call getTotal
    lstRec.SubItems(1) = txtCaseNo
    lstRec.SubItems(2) = MembershipNo
    lstRec.SubItems(3) = CovPersonID
    lstRec.SubItems(4) = txtPatientName
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
    lstRec.SubItems(110) = IIf(IsDate(txtInvRecDate), txtInvRecDate, "__/__/____")
    lstRec.SubItems(111) = Trim(Left(CboInvRecMode, IIf(InStr(1, CboInvRecMode, "-") - 1 < 0, Len(CboInvRecMode), InStr(1, CboInvRecMode, "-") - 1)))
    lstRec.SubItems(112) = IIf(Len(Trim(TxtRemark)), Trim(TxtRemark), "")
    lstRec.SubItems(113) = IIf(IsNumeric(txtMBMAdjAnnLmt), txtMBMAdjAnnLmt, 0)
    lstRec.SubItems(114) = IIf(IsNumeric(A31), A31, 0)
    ' -- --- reserved for invoice until (150) - last use 114
End Sub

Private Sub WSInfo()
Const METHOD_NAME = "WSInfo"
On Error GoTo debugs
Dim CISCLm As New ADODB.Recordset
Dim CISSql As String

    CISSql = "SELECT Top 1 * FROM Claim WHERE CASNumber='" & Trim(txtCaseNo) & "'"
    CISCLm.Open CISSql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    With CISCLm
        lblActCoPay = IIf(Len(!CLMCoPayCLNColl), Format(!CLMCoPayCLNColl, "#,##0.00"), 0)
        A29 = IIf(Len(!CLMCoPayCLNColl), Format(!CLMCoPayCLNColl, "#,##0.00"), 0)
        lstRec.SubItems(43) = IIf(IsNumeric(lblActCoPay), lblActCoPay, 0)
        lstRec.SubItems(471) = A29
        If Mid(CboPayto, 1, 1) = "M" Then
            lblPay2CLN = IIf(IsNumeric(!CLMPayableByMedix2M), Format(!CLMPayableByMedix2M, "#,##0.00"), 0)
            lstRec.SubItems(158) = IIf(IsNumeric(!CLMPayableByMedix2M), !CLMPayableByMedix2M, 0)
        Else
            lblPay2CLN = IIf(IsNumeric(!CLMPayableByMedix2c), Format(!CLMPayableByMedix2c, "#,##0.00"), 0)
            lstRec.SubItems(158) = IIf(IsNumeric(!CLMPayableByMedix2c), !CLMPayableByMedix2c, 0)
        End If
        lstRec.SubItems(159) = IIf(IsNumeric(!CLMCLNExcess), !CLMCLNExcess, 0)
        lstRec.SubItems(160) = IIf(IsNumeric(!CLMTotalActual), !CLMTotalActual, 0)
        lstRec.SubItems(161) = IIf(IsNumeric(!CLMFinalApprove), !CLMFinalApprove, 0)
        lstRec.SubItems(162) = IIf(IsNumeric(!CLMExcess), !CLMExcess, 0)
        lstRec.SubItems(163) = IIf(IsNumeric(!CLMBalanceLimit), !CLMBalanceLimit, 0)
        lstRec.SubItems(164) = IIf(IsNumeric(!CLMPreviousBalance), !CLMPreviousBalance, 0)
        lstRec.SubItems(165) = IIf(IsNumeric(!CLMPaidBYMedix2M), !CLMPaidBYMedix2M, 0)
        lstRec.SubItems(166) = IIf(IsNumeric(!CLMPaidBYMedix2C), !CLMPaidBYMedix2C, 0)
        lstRec.SubItems(166) = IIf(IsNumeric(!CLMPaidBYMedix2C), !CLMPaidBYMedix2C, 0)
        lblFFS = IIf(IsNumeric(!CLMMCOProcFee), Format(!CLMMCOProcFee, "#,##0.00"), 0)
        lstRec.SubItems(468) = IIf(IsNumeric(!CLMMCOProcFee), !CLMMCOProcFee, 0)
        cboFFS = IIf(Len(!CLMMCOProcStatus), !CLMMCOProcStatus, "N")
        lstRec.SubItems(469) = IIf(Len(!CLMMCOProcStatus), !CLMMCOProcStatus, "N")
        cboProcMth = IIf(Len(!CLMProcMth), !CLMProcMth & "/" & !CLMProcYR, "")
        lstRec.SubItems(470) = IIf(Len(!CLMProcMth), !CLMProcMth & "/" & !CLMProcYR, "")
    End With
    CISCLm.Close
    Set CISCLm = Nothing
    CISSql = "SELECT Top 1 * FROM CLMActualNew WHERE CLMNo='" & Trim(txtCaseNo) & "'"
    CISCLm.Open CISSql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    With CISCLm
        lstRec.SubItems(168) = IIf(IsNumeric(!CLMGPFeeBasicMed), !CLMGPFeeBasicMed, 0)
        lstRec.SubItems(169) = IIf(IsNumeric(!CLMGPFeeMedInj), !CLMGPFeeMedInj, 0)
        lstRec.SubItems(170) = IIf(IsNumeric(!CLMGPFeeLTMed), !CLMGPFeeLTMed, 0)
        lstRec.SubItems(171) = IIf(IsNumeric(!CLMSPFeeBasicMed), !CLMSPFeeBasicMed, 0)
        lstRec.SubItems(172) = IIf(IsNumeric(!CLMSPFeeLTMed), !CLMSPFeeLTMed, 0)
        lstRec.SubItems(173) = IIf(IsNumeric(!CLMSPFeeMedInj), !CLMSPFeeMedInj, 0)
        lstRec.SubItems(174) = IIf(IsNumeric(!CLMXRay), !CLMXRay, 0)
        lstRec.SubItems(175) = IIf(IsNumeric(!CLMUltraSound), !CLMUltraSound, 0)
        lstRec.SubItems(176) = IIf(IsNumeric(!CLMECG), !CLMECG, 0)
        lstRec.SubItems(177) = IIf(IsNumeric(!CLMDressing), !CLMDressing, 0)
        lstRec.SubItems(178) = IIf(IsNumeric(!CLMNebulizer), !CLMNebulizer, 0)
        lstRec.SubItems(179) = IIf(IsNumeric(!CLMID), !CLMID, 0)
        lstRec.SubItems(180) = IIf(IsNumeric(!CLMTSSTO), !CLMTSSTO, 0)
        lstRec.SubItems(181) = IIf(IsNumeric(!CLMMinorSurgery), !CLMMinorSurgery, 0)
        lstRec.SubItems(182) = IIf(IsNumeric(!CLMPOthers), !CLMPOthers, 0)
        lstRec.SubItems(183) = IIf(IsNumeric(!CLMBloodDrawing), !CLMBloodDrawing, 0)
        lstRec.SubItems(184) = IIf(IsNumeric(!CLMBloodGlucose), !CLMBloodGlucose, 0)
        lstRec.SubItems(185) = IIf(IsNumeric(!CLMUrine), !CLMUrine, 0)
        lstRec.SubItems(186) = IIf(IsNumeric(!CLMPapSmear), !CLMPapSmear, 0)
        lstRec.SubItems(187) = IIf(IsNumeric(!CLMLOthers), !CLMLOthers, 0)
        lstRec.SubItems(188) = IIf(IsNumeric(!CLMDTPPolio), !CLMDTPPolio, 0)
        lstRec.SubItems(189) = IIf(IsNumeric(!CLMDTPPolioHIB), !CLMDTPPolioHIB, 0)
        lstRec.SubItems(190) = IIf(IsNumeric(!CLMHepatitisB), !CLMHepatitisB, 0)
        lstRec.SubItems(191) = IIf(IsNumeric(!CLMMMR), !CLMMMR, 0)
        lstRec.SubItems(192) = IIf(IsNumeric(!CLMIOthers), !CLMIOthers, 0)
        lstRec.SubItems(193) = IIf(IsNumeric(!CLMDental), !CLMDental, 0)
        lstRec.SubItems(194) = IIf(IsNumeric(!CLMMaternity), !CLMMaternity, 0)
        lstRec.SubItems(195) = IIf(IsNumeric(!CLMPRVCheckup), !CLMPRVCheckup, 0)
        lstRec.SubItems(196) = IIf(IsNumeric(!CLMVision), !CLMVision, 0)
        lstRec.SubItems(197) = IIf(IsNumeric(!CLMTyphiadJabNoDed), !CLMTyphiadJabNoDed, 0)
        lstRec.SubItems(198) = IIf(IsNumeric(!CLMOthersNoDed), !CLMOthersNoDed, 0)
    End With
    CISCLm.Close
    Set CISCLm = Nothing
    CISSql = "SELECT Top 1 * FROM CLMClinicAppNew WHERE CASNumber='" & Trim(txtCaseNo) & "'"
    CISCLm.Open CISSql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly

    With CISCLm
        lstRec.SubItems(426) = IIf(IsNumeric(!CLMGPFeeBasicMed), !CLMGPFeeBasicMed, 0)
        lstRec.SubItems(427) = IIf(IsNumeric(!CLMGPFeeMedInj), !CLMGPFeeMedInj, 0)
        lstRec.SubItems(428) = IIf(IsNumeric(!CLMGPFeeLTMed), !CLMGPFeeLTMed, 0)
        lstRec.SubItems(429) = IIf(IsNumeric(!CLMSPFeeBasicMed), !CLMSPFeeBasicMed, 0)
        lstRec.SubItems(430) = IIf(IsNumeric(!CLMSPFeeLTMed), !CLMSPFeeLTMed, 0)
        lstRec.SubItems(431) = IIf(IsNumeric(!CLMSPFeeMedInj), !CLMSPFeeMedInj, 0)
        lstRec.SubItems(432) = IIf(IsNumeric(!CLMXRay), !CLMXRay, 0)
        lstRec.SubItems(433) = IIf(IsNumeric(!CLMUltraSound), !CLMUltraSound, 0)
        lstRec.SubItems(434) = IIf(IsNumeric(!CLMECG), !CLMECG, 0)
        lstRec.SubItems(435) = IIf(IsNumeric(!CLMDressing), !CLMDressing, 0)
        lstRec.SubItems(436) = IIf(IsNumeric(!CLMNebulizer), !CLMNebulizer, 0)
        lstRec.SubItems(437) = IIf(IsNumeric(!CLMID), !CLMID, 0)
        lstRec.SubItems(438) = IIf(IsNumeric(!CLMTSSTO), !CLMTSSTO, 0)
        lstRec.SubItems(439) = IIf(IsNumeric(!CLMMinorSurgery), !CLMMinorSurgery, 0)
        lstRec.SubItems(440) = IIf(IsNumeric(!CLMPOthers), !CLMPOthers, 0)
        lstRec.SubItems(441) = IIf(IsNumeric(!CLMBloodDrawing), !CLMBloodDrawing, 0)
        lstRec.SubItems(442) = IIf(IsNumeric(!CLMBloodGlucose), !CLMBloodGlucose, 0)
        lstRec.SubItems(443) = IIf(IsNumeric(!CLMUrine), !CLMUrine, 0)
        lstRec.SubItems(444) = IIf(IsNumeric(!CLMPapSmear), !CLMPapSmear, 0)
        lstRec.SubItems(445) = IIf(IsNumeric(!CLMLOthers), !CLMLOthers, 0)
        lstRec.SubItems(446) = IIf(IsNumeric(!CLMDTPPolio), !CLMDTPPolio, 0)
        lstRec.SubItems(447) = IIf(IsNumeric(!CLMDTPPolioHIB), !CLMDTPPolioHIB, 0)
        lstRec.SubItems(448) = IIf(IsNumeric(!CLMHepatitisB), !CLMHepatitisB, 0)
        lstRec.SubItems(449) = IIf(IsNumeric(!CLMMMR), !CLMMMR, 0)
        lstRec.SubItems(450) = IIf(IsNumeric(!CLMIOthers), !CLMIOthers, 0)
        lstRec.SubItems(451) = IIf(IsNumeric(!CLMDental), !CLMDental, 0)
        lstRec.SubItems(452) = IIf(IsNumeric(!CLMMaternity), !CLMMaternity, 0)
        lstRec.SubItems(453) = IIf(IsNumeric(!CLMPRVCheckup), !CLMPRVCheckup, 0)
        lstRec.SubItems(454) = IIf(IsNumeric(!CLMVision), !CLMVision, 0)
        lstRec.SubItems(455) = IIf(IsNumeric(!CLMTyphiadJabNoDed), !CLMTyphiadJabNoDed, 0)
        lstRec.SubItems(456) = IIf(IsNumeric(!CLMOthersNoDed), !CLMOthersNoDed, 0)
        
    End With
    CISCLm.Close
    Set CISCLm = Nothing
    
    CISSql = "SELECT Top 1 * FROM CLMClinicLmtNew WHERE CASNumber='" & Trim(txtCaseNo) & "'"
    CISCLm.Open CISSql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly

    With CISCLm
        lstRec.SubItems(211) = IIf(IsNumeric(!CLMGPFeeBasicMed), !CLMGPFeeBasicMed, 0)
        lstRec.SubItems(212) = IIf(IsNumeric(!CLMGPFeeMedInj), !CLMGPFeeMedInj, 0)
        lstRec.SubItems(213) = IIf(IsNumeric(!CLMGPFeeLTMed), !CLMGPFeeLTMed, 0)
        lstRec.SubItems(214) = IIf(IsNumeric(!CLMSPFeeBasicMed), !CLMSPFeeBasicMed, 0)
        lstRec.SubItems(215) = IIf(IsNumeric(!CLMSPFeeLTMed), !CLMSPFeeLTMed, 0)
        lstRec.SubItems(216) = IIf(IsNumeric(!CLMSPFeeMedInj), !CLMSPFeeMedInj, 0)
        lstRec.SubItems(217) = IIf(IsNumeric(!CLMXRay), !CLMXRay, 0)
        lstRec.SubItems(218) = IIf(IsNumeric(!CLMUltraSound), !CLMUltraSound, 0)
        lstRec.SubItems(219) = IIf(IsNumeric(!CLMECG), !CLMECG, 0)
        lstRec.SubItems(220) = IIf(IsNumeric(!CLMDressing), !CLMDressing, 0)
        lstRec.SubItems(221) = IIf(IsNumeric(!CLMNebulizer), !CLMNebulizer, 0)
        lstRec.SubItems(222) = IIf(IsNumeric(!CLMID), !CLMID, 0)
        lstRec.SubItems(223) = IIf(IsNumeric(!CLMTSSTO), !CLMTSSTO, 0)
        lstRec.SubItems(224) = IIf(IsNumeric(!CLMMinorSurgery), !CLMMinorSurgery, 0)
        lstRec.SubItems(225) = IIf(IsNumeric(!CLMPOthers), !CLMPOthers, 0)
        lstRec.SubItems(226) = IIf(IsNumeric(!CLMBloodDrawing), !CLMBloodDrawing, 0)
        lstRec.SubItems(227) = IIf(IsNumeric(!CLMBloodGlucose), !CLMBloodGlucose, 0)
        lstRec.SubItems(228) = IIf(IsNumeric(!CLMUrine), !CLMUrine, 0)
        lstRec.SubItems(229) = IIf(IsNumeric(!CLMPapSmear), !CLMPapSmear, 0)
        lstRec.SubItems(230) = IIf(IsNumeric(!CLMLOthers), !CLMLOthers, 0)
        lstRec.SubItems(231) = IIf(IsNumeric(!CLMDTPPolio), !CLMDTPPolio, 0)
        lstRec.SubItems(232) = IIf(IsNumeric(!CLMDTPPolioHIB), !CLMDTPPolioHIB, 0)
        lstRec.SubItems(233) = IIf(IsNumeric(!CLMHepatitisB), !CLMHepatitisB, 0)
        lstRec.SubItems(234) = IIf(IsNumeric(!CLMMMR), !CLMMMR, 0)
        lstRec.SubItems(235) = IIf(IsNumeric(!CLMIOthers), !CLMIOthers, 0)
        lstRec.SubItems(236) = IIf(IsNumeric(!CLMDental), !CLMDental, 0)
        lstRec.SubItems(237) = IIf(IsNumeric(!CLMMaternity), !CLMMaternity, 0)
        lstRec.SubItems(238) = IIf(IsNumeric(!CLMPRVCheckup), !CLMPRVCheckup, 0)
        lstRec.SubItems(239) = IIf(IsNumeric(!CLMVision), !CLMVision, 0)
        lstRec.SubItems(240) = IIf(IsNumeric(!CLMTyphiadJabNoDed), !CLMTyphiadJabNoDed, 0)
        lstRec.SubItems(241) = IIf(IsNumeric(!CLMOthersNoDed), !CLMOthersNoDed, 0)
        
    End With
    CISCLm.Close
    Set CISCLm = Nothing
    
    CISSql = "SELECT Top 1 * FROM CLMClinicNew WHERE CASNumber='" & Trim(txtCaseNo) & "'"
    CISCLm.Open CISSql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    With CISCLm
        lstRec.SubItems(383) = IIf(IsNumeric(!CLMGPFeeBasicMed), !CLMGPFeeBasicMed, 0)
        lstRec.SubItems(384) = IIf(IsNumeric(!CLMGPFeeMedInj), !CLMGPFeeMedInj, 0)
        lstRec.SubItems(385) = IIf(IsNumeric(!CLMGPFeeLTMed), !CLMGPFeeLTMed, 0)
        lstRec.SubItems(386) = IIf(IsNumeric(!CLMSPFeeBasicMed), !CLMSPFeeBasicMed, 0)
        lstRec.SubItems(387) = IIf(IsNumeric(!CLMSPFeeLTMed), !CLMSPFeeLTMed, 0)
        lstRec.SubItems(388) = IIf(IsNumeric(!CLMSPFeeMedInj), !CLMSPFeeMedInj, 0)
        lstRec.SubItems(389) = IIf(IsNumeric(!CLMXRay), !CLMXRay, 0)
        lstRec.SubItems(390) = IIf(IsNumeric(!CLMUltraSound), !CLMUltraSound, 0)
        lstRec.SubItems(391) = IIf(IsNumeric(!CLMECG), !CLMECG, 0)
        lstRec.SubItems(392) = IIf(IsNumeric(!CLMDressing), !CLMDressing, 0)
        lstRec.SubItems(393) = IIf(IsNumeric(!CLMNebulizer), !CLMNebulizer, 0)
        lstRec.SubItems(394) = IIf(IsNumeric(!CLMID), !CLMID, 0)
        lstRec.SubItems(395) = IIf(IsNumeric(!CLMTSSTO), !CLMTSSTO, 0)
        lstRec.SubItems(396) = IIf(IsNumeric(!CLMMinorSurgery), !CLMMinorSurgery, 0)
        lstRec.SubItems(397) = IIf(IsNumeric(!CLMPOthers), !CLMPOthers, 0)
        lstRec.SubItems(398) = IIf(IsNumeric(!CLMBloodDrawing), !CLMBloodDrawing, 0)
        lstRec.SubItems(399) = IIf(IsNumeric(!CLMBloodGlucose), !CLMBloodGlucose, 0)
        lstRec.SubItems(400) = IIf(IsNumeric(!CLMUrine), !CLMUrine, 0)
        lstRec.SubItems(401) = IIf(IsNumeric(!CLMPapSmear), !CLMPapSmear, 0)
        lstRec.SubItems(402) = IIf(IsNumeric(!CLMLOthers), !CLMLOthers, 0)
        lstRec.SubItems(403) = IIf(IsNumeric(!CLMDTPPolio), !CLMDTPPolio, 0)
        lstRec.SubItems(404) = IIf(IsNumeric(!CLMDTPPolioHIB), !CLMDTPPolioHIB, 0)
        lstRec.SubItems(405) = IIf(IsNumeric(!CLMHepatitisB), !CLMHepatitisB, 0)
        lstRec.SubItems(406) = IIf(IsNumeric(!CLMMMR), !CLMMMR, 0)
        lstRec.SubItems(407) = IIf(IsNumeric(!CLMIOthers), !CLMIOthers, 0)
        lstRec.SubItems(408) = IIf(IsNumeric(!CLMDental), !CLMDental, 0)
        lstRec.SubItems(409) = IIf(IsNumeric(!CLMMaternity), !CLMMaternity, 0)
        lstRec.SubItems(410) = IIf(IsNumeric(!CLMPRVCheckup), !CLMPRVCheckup, 0)
        lstRec.SubItems(411) = IIf(IsNumeric(!CLMVision), !CLMVision, 0)
        lstRec.SubItems(412) = IIf(IsNumeric(!CLMTyphiadJabNoDed), !CLMTyphiadJabNoDed, 0)
        lstRec.SubItems(413) = IIf(IsNumeric(!CLMOthersNoDed), !CLMOthersNoDed, 0)

    End With
    CISCLm.Close
    Set CISCLm = Nothing
    
    CISSql = "SELECT Top 1 * FROM CLMAppNew WHERE CASNumber='" & Trim(txtCaseNo) & "'"
    CISCLm.Open CISSql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    With CISCLm
        lstRec.SubItems(254) = IIf(IsNumeric(!CLMGPFeeBasicMed), !CLMGPFeeBasicMed, 0)
        lstRec.SubItems(255) = IIf(IsNumeric(!CLMGPFeeMedInj), !CLMGPFeeMedInj, 0)
        lstRec.SubItems(256) = IIf(IsNumeric(!CLMGPFeeLTMed), !CLMGPFeeLTMed, 0)
        lstRec.SubItems(257) = IIf(IsNumeric(!CLMSPFeeBasicMed), !CLMSPFeeBasicMed, 0)
        lstRec.SubItems(258) = IIf(IsNumeric(!CLMSPFeeLTMed), !CLMSPFeeLTMed, 0)
        lstRec.SubItems(259) = IIf(IsNumeric(!CLMSPFeeMedInj), !CLMSPFeeMedInj, 0)
        lstRec.SubItems(260) = IIf(IsNumeric(!CLMXRay), !CLMXRay, 0)
        lstRec.SubItems(261) = IIf(IsNumeric(!CLMUltraSound), !CLMUltraSound, 0)
        lstRec.SubItems(262) = IIf(IsNumeric(!CLMECG), !CLMECG, 0)
        lstRec.SubItems(263) = IIf(IsNumeric(!CLMDressing), !CLMDressing, 0)
        lstRec.SubItems(264) = IIf(IsNumeric(!CLMNebulizer), !CLMNebulizer, 0)
        lstRec.SubItems(265) = IIf(IsNumeric(!CLMID), !CLMID, 0)
        lstRec.SubItems(266) = IIf(IsNumeric(!CLMTSSTO), !CLMTSSTO, 0)
        lstRec.SubItems(267) = IIf(IsNumeric(!CLMMinorSurgery), !CLMMinorSurgery, 0)
        lstRec.SubItems(268) = IIf(IsNumeric(!CLMPOthers), !CLMPOthers, 0)
        lstRec.SubItems(269) = IIf(IsNumeric(!CLMBloodDrawing), !CLMBloodDrawing, 0)
        lstRec.SubItems(270) = IIf(IsNumeric(!CLMBloodGlucose), !CLMBloodGlucose, 0)
        lstRec.SubItems(271) = IIf(IsNumeric(!CLMUrine), !CLMUrine, 0)
        lstRec.SubItems(272) = IIf(IsNumeric(!CLMPapSmear), !CLMPapSmear, 0)
        lstRec.SubItems(273) = IIf(IsNumeric(!CLMLOthers), !CLMLOthers, 0)
        lstRec.SubItems(274) = IIf(IsNumeric(!CLMDTPPolio), !CLMDTPPolio, 0)
        lstRec.SubItems(275) = IIf(IsNumeric(!CLMDTPPolioHIB), !CLMDTPPolioHIB, 0)
        lstRec.SubItems(276) = IIf(IsNumeric(!CLMHepatitisB), !CLMHepatitisB, 0)
        lstRec.SubItems(277) = IIf(IsNumeric(!CLMMMR), !CLMMMR, 0)
        lstRec.SubItems(278) = IIf(IsNumeric(!CLMIOthers), !CLMIOthers, 0)
        lstRec.SubItems(279) = IIf(IsNumeric(!CLMDental), !CLMDental, 0)
        lstRec.SubItems(280) = IIf(IsNumeric(!CLMMaternity), !CLMMaternity, 0)
        lstRec.SubItems(281) = IIf(IsNumeric(!CLMPRVCheckup), !CLMPRVCheckup, 0)
        lstRec.SubItems(282) = IIf(IsNumeric(!CLMVision), !CLMVision, 0)
        lstRec.SubItems(283) = IIf(IsNumeric(!CLMTyphiadJabNoDed), !CLMTyphiadJabNoDed, 0)
        lstRec.SubItems(284) = IIf(IsNumeric(!CLMOthersNoDed), !CLMOthersNoDed, 0)
        
    End With
    CISCLm.Close
    Set CISCLm = Nothing
    
    CISSql = "SELECT Top 1 * FROM CLMMemberNew WHERE CASNumber='" & Trim(txtCaseNo) & "'"
    CISCLm.Open CISSql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
    With CISCLm
        lstRec.SubItems(340) = IIf(IsNumeric(!CLMGPFeeBasicMed), !CLMGPFeeBasicMed, 0)
        lstRec.SubItems(341) = IIf(IsNumeric(!CLMGPFeeMedInj), !CLMGPFeeMedInj, 0)
        lstRec.SubItems(342) = IIf(IsNumeric(!CLMGPFeeLTMed), !CLMGPFeeLTMed, 0)
        lstRec.SubItems(343) = IIf(IsNumeric(!CLMSPFeeBasicMed), !CLMSPFeeBasicMed, 0)
        lstRec.SubItems(344) = IIf(IsNumeric(!CLMSPFeeLTMed), !CLMSPFeeLTMed, 0)
        lstRec.SubItems(345) = IIf(IsNumeric(!CLMSPFeeMedInj), !CLMSPFeeMedInj, 0)
        lstRec.SubItems(346) = IIf(IsNumeric(!CLMXRay), !CLMXRay, 0)
        lstRec.SubItems(347) = IIf(IsNumeric(!CLMUltraSound), !CLMUltraSound, 0)
        lstRec.SubItems(348) = IIf(IsNumeric(!CLMECG), !CLMECG, 0)
        lstRec.SubItems(349) = IIf(IsNumeric(!CLMDressing), !CLMDressing, 0)
        lstRec.SubItems(350) = IIf(IsNumeric(!CLMNebulizer), !CLMNebulizer, 0)
        lstRec.SubItems(351) = IIf(IsNumeric(!CLMID), !CLMID, 0)
        lstRec.SubItems(352) = IIf(IsNumeric(!CLMTSSTO), !CLMTSSTO, 0)
        lstRec.SubItems(353) = IIf(IsNumeric(!CLMMinorSurgery), !CLMMinorSurgery, 0)
        lstRec.SubItems(354) = IIf(IsNumeric(!CLMPOthers), !CLMPOthers, 0)
        lstRec.SubItems(355) = IIf(IsNumeric(!CLMBloodDrawing), !CLMBloodDrawing, 0)
        lstRec.SubItems(356) = IIf(IsNumeric(!CLMBloodGlucose), !CLMBloodGlucose, 0)
        lstRec.SubItems(357) = IIf(IsNumeric(!CLMUrine), !CLMUrine, 0)
        lstRec.SubItems(358) = IIf(IsNumeric(!CLMPapSmear), !CLMPapSmear, 0)
        lstRec.SubItems(359) = IIf(IsNumeric(!CLMLOthers), !CLMLOthers, 0)
        lstRec.SubItems(360) = IIf(IsNumeric(!CLMDTPPolio), !CLMDTPPolio, 0)
        lstRec.SubItems(361) = IIf(IsNumeric(!CLMDTPPolioHIB), !CLMDTPPolioHIB, 0)
        lstRec.SubItems(362) = IIf(IsNumeric(!CLMHepatitisB), !CLMHepatitisB, 0)
        lstRec.SubItems(363) = IIf(IsNumeric(!CLMMMR), !CLMMMR, 0)
        lstRec.SubItems(364) = IIf(IsNumeric(!CLMIOthers), !CLMIOthers, 0)
        lstRec.SubItems(365) = IIf(IsNumeric(!CLMDental), !CLMDental, 0)
        lstRec.SubItems(366) = IIf(IsNumeric(!CLMMaternity), !CLMMaternity, 0)
        lstRec.SubItems(367) = IIf(IsNumeric(!CLMPRVCheckup), !CLMPRVCheckup, 0)
        lstRec.SubItems(368) = IIf(IsNumeric(!CLMVision), !CLMVision, 0)
        lstRec.SubItems(369) = IIf(IsNumeric(!CLMTyphiadJabNoDed), !CLMTyphiadJabNoDed, 0)
        lstRec.SubItems(370) = IIf(IsNumeric(!CLMOthersNoDed), !CLMOthersNoDed, 0)
    End With
    CISCLm.Close
    Set CISCLm = Nothing
    
    CISSql = "SELECT Top 1 * FROM CLMLimitNew WHERE CASNumber='" & Trim(txtCaseNo) & "'"
    CISCLm.Open CISSql, MediConnCISWS, adOpenForwardOnly, adLockReadOnly
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
        If DiagRec.EOF = False Then
            tmpDiagDesc = DiagRec!ReasonDesc
        End If
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
    lblDispBillAmt = Format(tmpBillAmt, "#,##0.00")
    lblDispCLNBillAmt = Format(tmpBillAmt, "#,##0.00")
End Sub

Private Sub ClaimBordxInfo()
Const METHOD_NAME = "ClaimBordxInfo"
On Error GoTo debugs
Dim cmd2 As New ADODB.Command
Dim prm2 As New ADODB.Parameter
Dim StrCLMSql As String
Dim CLmRec As New ADODB.Recordset
Dim MedixList As ListItem
Dim tmpMedixCASNo As String
Dim tmpCAsNo As String
Dim tmpCLMVer As String
    Set cmd2.ActiveConnection = MediConnCISWS
    cmd2.CommandText = "SP_Claim_Bordx"
    cmd2.CommandType = adCmdStoredProc
    cmd2.CommandTimeout = 300
    Set prm2 = cmd2.CreateParameter("tmpCasNo", adVarChar, adParamInput, 20, Trim(txtCaseNo))
    cmd2.Parameters.Append prm2
    CLmRec.CursorLocation = adUseClient
    CLmRec.CursorType = adOpenForwardOnly
    CLmRec.CacheSize = 100
    Set CLmRec = cmd2.Execute
    Do While Not CLmRec.EOF
        A31 = IIf(IsNumeric(CLmRec!ClmNonGPFee), CLmRec!ClmNonGPFee, 0)
        txtMBMAdjAnnLmt = IIf(IsNumeric(CLmRec!CLMAdjAnnLmt), CLmRec!CLMAdjAnnLmt, 0)
        tmpFinalApp = IIf(IsNumeric(CLmRec!CLMFinalApprove), CLmRec!CLMFinalApprove, 0)
        lblBordx = IIf(Len(CLmRec!ClaimBordxStatus), CLmRec!ClaimBordxStatus, "False")
        lblBordxNo = IIf(Len(CLmRec!Bordxrefno), CLmRec!Bordxrefno, "")
        lblBordxDate = IIf(Len(CLmRec!BordxDate), CLmRec!BordxDate, "")
        Set MedixList = lstMedixPay.ListItems.Add(, , "")
        tmpCAsNo = IIf(Len(CLmRec!MedixCaseNo), CLmRec!MedixCaseNo, "")
        tmpCLMVer = IIf(Len(CLmRec!MedixClaimVer), CLmRec!MedixClaimVer, "")
        tmpMedixCASNo = tmpCAsNo & "-" & tmpCLMVer
        MedixList.SubItems(1) = tmpMedixCASNo
        MedixList.SubItems(2) = IIf(Len(CLmRec!MedixPVNO), CLmRec!MedixPVNO, "")
        MedixList.SubItems(3) = IIf(Len(CLmRec!MedixPVDate), CLmRec!MedixPVDate, "")
        CLmRec.MoveNext
    Loop
    CLmRec.Close
    Set CLmRec = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub MedicationRequestInfo()
Const METHOD_NAME = "MedicationRequestInfo"
On Error GoTo debugs
Dim n As Integer
Dim i As Integer
Dim MedReqList As ListItem
    
    lvwMedRequest.ListItems.Clear
    lvwMedRequest.ColumnHeaders.Clear
    Dim MRcmd As New ADODB.Command
    Dim MRPrm1 As New ADODB.Parameter
    Dim MRRes As New ADODB.Recordset
        Set MRcmd.ActiveConnection = MediConnCISDB
        MRcmd.CommandText = "sp_CISMedRequestInfo"
        MRcmd.CommandType = adCmdStoredProc
        MRcmd.CommandTimeout = 300
        Set MRPrm1 = MRcmd.CreateParameter("CaseNo", adVarChar, adParamInput, 100, Trim(txtCaseNo.Text))
        MRcmd.Parameters.Append MRPrm1
        MRRes.CursorLocation = adUseClient
        MRRes.CursorType = adOpenForwardOnly
        MRRes.CacheSize = 100
        Set MRRes = MRcmd.Execute
        Do While Not MRRes.EOF
            For n = 0 To MRRes.Fields.Count - 1
                If MRRes.Fields(n).Type = 6 Then
                    lvwMedRequest.ColumnHeaders.Add , , MRRes.Fields(n).Name, , lvwColumnRight
                Else
                    lvwMedRequest.ColumnHeaders.Add , , MRRes.Fields(n).Name
                End If
            Next n
            Set MedReqList = lvwMedRequest.ListItems.Add(, , MRRes(0) & vbNullString)
            For i = 1 To MRRes.Fields.Count - 1
                If MRRes(i).Type = 6 Then
                    MedReqList.SubItems(i) = Format(MRRes(i), "#,##0.00")
                Else
                    MedReqList.SubItems(i) = MRRes(i) & vbNullString
                End If
            Next i
            MRRes.MoveNext
        Loop
        MRRes.Close
        Set MRRes = Nothing

debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

Private Sub PaymentInfo()
Const METHOD_NAME = "PaymentInfo"
On Error GoTo debugs
Dim StrSql As String
Dim PayRec As New ADODB.Recordset
Dim PayList As ListItem
    lstPayment.ListItems.Clear
    StrSql = ""
    StrSql = "SELECT PVCLNNo, PVCLNDate,  CLNPaidAmt, PVCLNChequeAmt, " & _
             "CLNRemark , PVCLNChequeDate From Payment2CLN WHERE CASNumber='" & Trim(txtCaseNo) & "'"
    PayRec.Open StrSql, MediConnCISPayment, adOpenForwardOnly, adLockOptimistic
    Do While Not PayRec.EOF
        Set PayList = lstPayment.ListItems.Add(, , "")
        PayList.SubItems(1) = IIf(Len(PayRec!PVCLNNo), PayRec!PVCLNNo, "")
        PayList.SubItems(2) = IIf(Len(PayRec!PVCLNDate), PayRec!PVCLNDate, "")
        PayList.SubItems(3) = Format(IIf(Len(PayRec!PVCLNChequeAmt), PayRec!PVCLNChequeAmt, ""), "#,##0.00")
        PayList.SubItems(4) = "CLINIC" ' pay to clinic
        PayList.SubItems(5) = Format(IIf(IsNumeric(PayRec!CLNPaidAmt), PayRec!CLNPaidAmt, 0), "#,##0.00")
        PayList.SubItems(6) = IIf(Len(PayRec!CLNRemark), PayRec!CLNRemark, "")
        PayList.SubItems(7) = IIf(Len(PayRec!PVCLNChequeDate), PayRec!PVCLNChequeDate, "")
        PayRec.MoveNext
    Loop
    PayRec.Close
    Set PayRec = Nothing
    
    StrSql = ""
    StrSql = "SELECT PVNo, PVDate, MBMPaidAmt, PVPayto, PVCLNChequeAmt, " & _
             "Remark,PVChequeDate From Payment2MBM WHERE CASNumber='" & Trim(txtCaseNo) & "'"
    PayRec.Open StrSql, MediConnCISPayment, adOpenForwardOnly, adLockOptimistic
    Do While Not PayRec.EOF
        Set PayList = lstPayment.ListItems.Add(, , "")
        PayList.SubItems(1) = IIf(Len(PayRec!PVNo), PayRec!PVNo, "")
        PayList.SubItems(2) = IIf(Len(PayRec!PVDate), PayRec!PVDate, "")
        PayList.SubItems(3) = Format(IIf(Len(PayRec!PVCLNChequeAmt), PayRec!PVCLNChequeAmt, ""), "#,##0.00")
        PayList.SubItems(4) = IIf(Len(PayRec!PVPayto), PayRec!PVPayto, "")
        PayList.SubItems(5) = Format(IIf(IsNumeric(PayRec!MBMPaidAmt), PayRec!MBMPaidAmt, 0), "#,##0.00")
        PayList.SubItems(6) = IIf(Len(PayRec!Remark), PayRec!Remark, "")
        PayList.SubItems(7) = IIf(Len(PayRec!PVChequeDate), PayRec!PVChequeDate, "")
        PayRec.MoveNext
    Loop
    PayRec.Close
    Set PayRec = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub

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
Private Sub ReceiptInfo()
Const METHOD_NAME = "ReceiptInfo"
On Error GoTo debugs

Dim StrSql As String
Dim ReceiptRec As New ADODB.Recordset
Dim ReceiptList As ListItem
    lstReceipt.ListItems.Clear
    StrSql = "SELECT ReceiptNo, ReceiptDate, ReceiptAmt From ReceiptFrMBM WHERE CASNumber='" & Trim(txtCaseNo) & "'"
    ReceiptRec.Open StrSql, MediConnCISPayment, adOpenForwardOnly, adLockOptimistic
    Do While Not ReceiptRec.EOF
        Set ReceiptList = lstReceipt.ListItems.Add(, , "")
        ReceiptList.SubItems(1) = IIf(Len(ReceiptRec!ReceiptNo), ReceiptRec!ReceiptNo, "")
        ReceiptList.SubItems(2) = IIf(Len(ReceiptRec!ReceiptDate), ReceiptRec!ReceiptDate, "")
        ReceiptList.SubItems(3) = "Member" ' pay to clinic
        ReceiptList.SubItems(4) = Format(IIf(Len(ReceiptRec!ReceiptAmt), ReceiptRec!ReceiptAmt, ""), "#,##0.00")
        ReceiptRec.MoveNext
    Loop
    ReceiptRec.Close
    Set ReceiptRec = Nothing
    
    StrSql = "SELECT ReceiptNo, ReceiptDate, ReceiptAmt From ReceiptfrPAY WHERE CASNumber='" & Trim(txtCaseNo) & "'"
    ReceiptRec.Open StrSql, MediConnCISPayment, adOpenForwardOnly, adLockOptimistic
    Do While Not ReceiptRec.EOF
        Set ReceiptList = lstReceipt.ListItems.Add(, , "")
        ReceiptList.SubItems(1) = IIf(Len(ReceiptRec!ReceiptNo), ReceiptRec!ReceiptNo, "")
        ReceiptList.SubItems(2) = IIf(Len(ReceiptRec!ReceiptDate), ReceiptRec!ReceiptDate, "")
        ReceiptList.SubItems(3) = "Grp Company" ' pay to clinic
        ReceiptList.SubItems(4) = Format(IIf(Len(ReceiptRec!ReceiptAmt), ReceiptRec!ReceiptAmt, ""), "#,##0.00")
        ReceiptRec.MoveNext
    Loop
    ReceiptRec.Close
    Set ReceiptRec = Nothing
debugs:
If Err.Description <> "" Then
    Call ErrorLog(Me.Caption, METHOD_NAME, Err.Description)
End If
End Sub
