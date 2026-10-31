VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmInvDentalProcessing 
   Caption         =   "InvDentalProcessing"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   8505
   StartUpPosition =   3  'Windows Default
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
      Height          =   2130
      Left            =   7080
      TabIndex        =   223
      Top             =   1920
      Width           =   4815
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
         Left            =   315
         MaxLength       =   50
         TabIndex        =   238
         Top             =   525
         Width           =   3495
      End
      Begin VB.TextBox txtMedType2 
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
         Left            =   315
         TabIndex        =   237
         Top             =   810
         Width           =   3495
      End
      Begin VB.TextBox txtMedType3 
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
         Left            =   315
         TabIndex        =   236
         Top             =   1065
         Width           =   3495
      End
      Begin VB.TextBox txtMedType4 
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
         Left            =   315
         TabIndex        =   235
         Top             =   1350
         Width           =   3495
      End
      Begin VB.TextBox txtMedType5 
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
         Left            =   315
         TabIndex        =   234
         Top             =   1605
         Width           =   3495
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
         Left            =   3780
         MaxLength       =   3
         TabIndex        =   233
         Top             =   525
         Width           =   495
      End
      Begin VB.TextBox txtMedDay2 
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
         Left            =   3780
         MaxLength       =   3
         TabIndex        =   232
         Top             =   810
         Width           =   495
      End
      Begin VB.TextBox txtMedDay3 
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
         Left            =   3780
         MaxLength       =   3
         TabIndex        =   231
         Top             =   1065
         Width           =   495
      End
      Begin VB.TextBox txtMedDay4 
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
         Left            =   3780
         MaxLength       =   3
         TabIndex        =   230
         Top             =   1350
         Width           =   495
      End
      Begin VB.TextBox txtMedDay5 
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
         Left            =   3780
         MaxLength       =   3
         TabIndex        =   229
         Top             =   1605
         Width           =   495
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
         Left            =   4245
         MaxLength       =   3
         TabIndex        =   228
         Top             =   525
         Width           =   495
      End
      Begin VB.TextBox txtMedQty2 
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
         Left            =   4245
         MaxLength       =   3
         TabIndex        =   227
         Top             =   810
         Width           =   495
      End
      Begin VB.TextBox txtMedQty3 
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
         Left            =   4245
         MaxLength       =   3
         TabIndex        =   226
         Top             =   1065
         Width           =   495
      End
      Begin VB.TextBox txtMedQty4 
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
         Left            =   4245
         MaxLength       =   3
         TabIndex        =   225
         Top             =   1350
         Width           =   495
      End
      Begin VB.TextBox txtMedQty5 
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
         Left            =   4245
         MaxLength       =   3
         TabIndex        =   224
         Top             =   1605
         Width           =   495
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
         Left            =   3840
         TabIndex        =   246
         Top             =   240
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
         Left            =   4320
         TabIndex        =   245
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   244
         Top             =   1350
         Width           =   975
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000013&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   243
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   242
         Top             =   825
         Width           =   375
      End
      Begin VB.Label Label41 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   241
         Top             =   525
         Width           =   375
      End
      Begin VB.Label Label50 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   240
         Top             =   1635
         Width           =   975
      End
      Begin VB.Label Label70 
         Caption         =   "Type"
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
         TabIndex        =   239
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame FrameDiag 
      Height          =   1930
      Left            =   7080
      TabIndex        =   200
      Top             =   4080
      Width           =   4815
      Begin VB.CommandButton cmdSearchICD2 
         Caption         =   ">"
         Height          =   255
         Left            =   4440
         TabIndex        =   216
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton cmdSearchICD3 
         Caption         =   ">"
         Height          =   255
         Left            =   4440
         TabIndex        =   215
         Top             =   1010
         Width           =   255
      End
      Begin VB.CommandButton cmdSearchICD1 
         Caption         =   ">"
         Height          =   255
         Left            =   4440
         TabIndex        =   214
         Top             =   440
         Width           =   255
      End
      Begin VB.CommandButton cmdSearchICD4 
         Caption         =   ">"
         Height          =   255
         Left            =   4440
         TabIndex        =   213
         Top             =   1290
         Width           =   255
      End
      Begin VB.CommandButton cmdSearchICD5 
         Caption         =   ">"
         Height          =   255
         Left            =   4440
         TabIndex        =   212
         Top             =   1580
         Width           =   255
      End
      Begin VB.TextBox txtDiagnosis1 
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
         Left            =   315
         MaxLength       =   10
         TabIndex        =   211
         Top             =   440
         Width           =   615
      End
      Begin VB.TextBox txtDiagnosis2 
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
         Left            =   315
         MaxLength       =   10
         TabIndex        =   210
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtDiagnosis3 
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
         Left            =   315
         MaxLength       =   10
         TabIndex        =   209
         Top             =   1000
         Width           =   615
      End
      Begin VB.TextBox txtDiagnosis4 
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
         Left            =   315
         MaxLength       =   10
         TabIndex        =   208
         Top             =   1280
         Width           =   615
      End
      Begin VB.TextBox TxtDiagnosis5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   320
         MaxLength       =   100
         TabIndex        =   207
         Top             =   1560
         Width           =   615
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
         Left            =   960
         MaxLength       =   10
         TabIndex        =   206
         Top             =   440
         Width           =   3495
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
         Left            =   960
         MaxLength       =   10
         TabIndex        =   205
         Top             =   720
         Width           =   3495
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
         Left            =   960
         MaxLength       =   10
         TabIndex        =   204
         Top             =   1000
         Width           =   3495
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
         Left            =   960
         MaxLength       =   10
         TabIndex        =   203
         Top             =   1280
         Width           =   3495
      End
      Begin VB.TextBox TxtDiagnosisOth 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         MaxLength       =   100
         TabIndex        =   202
         Top             =   1560
         Width           =   3495
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
         Left            =   960
         MaxLength       =   10
         TabIndex        =   201
         Top             =   1560
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "Diagnosis"
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
         TabIndex        =   222
         Top             =   170
         Width           =   735
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
         TabIndex        =   221
         Top             =   440
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
         TabIndex        =   220
         Top             =   740
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
         TabIndex        =   219
         Top             =   1040
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
         TabIndex        =   218
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label10 
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
         TabIndex        =   217
         Top             =   1580
         Width           =   255
      End
   End
   Begin VB.Frame FrameMBMCLN 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   0
      TabIndex        =   187
      Top             =   360
      Width           =   5040
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
         Left            =   5640
         TabIndex        =   194
         Text            =   "N - No"
         Top             =   1320
         Width           =   855
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
         Height          =   270
         Left            =   4260
         TabIndex        =   193
         Top             =   195
         Width           =   660
      End
      Begin VB.CommandButton cmdSearch 
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
         Height          =   270
         Left            =   4260
         TabIndex        =   192
         Top             =   450
         Width           =   660
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
         Left            =   1080
         TabIndex        =   191
         Top             =   195
         Width           =   3135
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
         Left            =   1080
         TabIndex        =   190
         Top             =   450
         Width           =   3135
      End
      Begin VB.ComboBox CboClinic 
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
         Left            =   2040
         TabIndex        =   189
         Top             =   700
         Width           =   2895
      End
      Begin VB.TextBox TxtClinicCode 
         Height          =   315
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   188
         Top             =   740
         Width           =   855
      End
      Begin VB.Label Label54 
         Caption         =   "MBM No."
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
         TabIndex        =   199
         Top             =   210
         Width           =   1095
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
         TabIndex        =   198
         Top             =   740
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "MBM Name"
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
         TabIndex        =   197
         Top             =   480
         Width           =   1095
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
         Left            =   4560
         TabIndex        =   196
         Top             =   1140
         Width           =   825
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
         Left            =   4800
         TabIndex        =   195
         Top             =   1380
         Width           =   660
      End
   End
   Begin VB.Frame Frame7 
      Height          =   555
      Left            =   0
      TabIndex        =   173
      Top             =   6000
      Width           =   11895
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
         Left            =   8280
         TabIndex        =   186
         Top             =   160
         Width           =   615
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
         Left            =   8280
         TabIndex        =   185
         Top             =   160
         Width           =   615
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
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
         Left            =   9600
         TabIndex        =   184
         Top             =   160
         Width           =   615
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
         Left            =   9600
         TabIndex        =   183
         Top             =   160
         Width           =   615
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
         Left            =   5160
         TabIndex        =   182
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
         Left            =   5760
         TabIndex        =   181
         Top             =   160
         Width           =   975
      End
      Begin VB.CommandButton CmdClearInv 
         Caption         =   "Clear &Bill"
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
         Left            =   7320
         TabIndex        =   180
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
         Left            =   10200
         TabIndex        =   179
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
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
         Left            =   11040
         TabIndex        =   178
         Top             =   160
         Width           =   615
      End
      Begin VB.CommandButton CmdAddRec 
         Caption         =   "&Add"
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
         Left            =   6720
         TabIndex        =   177
         Top             =   160
         Width           =   615
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
         Left            =   8880
         TabIndex        =   176
         Top             =   160
         Width           =   735
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
         Left            =   4320
         TabIndex        =   175
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdTotInv 
         Caption         =   "&Tot Inv."
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
         Left            =   3480
         TabIndex        =   174
         Top             =   160
         Width           =   855
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
      Height          =   435
      Left            =   0
      TabIndex        =   164
      Top             =   1440
      Width           =   8745
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
         Height          =   255
         Left            =   7800
         MaxLength       =   3
         TabIndex        =   166
         Top             =   120
         Width           =   855
      End
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
         Height          =   255
         Left            =   960
         MaxLength       =   20
         TabIndex        =   165
         Top             =   120
         Width           =   2175
      End
      Begin MSMask.MaskEdBox txtVisitDate 
         Height          =   240
         Left            =   3960
         TabIndex        =   167
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   423
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
         Height          =   255
         Left            =   5760
         TabIndex        =   168
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
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
         TabIndex        =   172
         Top             =   135
         Width           =   975
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
         Left            =   3240
         TabIndex        =   171
         Top             =   120
         Width           =   975
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
         Left            =   5040
         TabIndex        =   170
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "MC Days"
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
         Left            =   6840
         TabIndex        =   169
         Top             =   120
         Width           =   975
      End
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
      Height          =   4215
      Left            =   0
      TabIndex        =   31
      Top             =   1800
      Width           =   7065
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
         Height          =   2835
         Left            =   8880
         TabIndex        =   75
         Top             =   120
         Visible         =   0   'False
         Width           =   2370
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
            Height          =   255
            Left            =   1080
            TabIndex        =   84
            Top             =   1920
            Visible         =   0   'False
            Width           =   975
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
            Left            =   1080
            TabIndex        =   83
            Top             =   720
            Width           =   975
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
            Left            =   1080
            TabIndex        =   82
            Top             =   240
            Width           =   495
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
            Left            =   1080
            TabIndex        =   81
            Top             =   480
            Width           =   495
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
            Left            =   1080
            TabIndex        =   80
            Top             =   1200
            Width           =   975
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
            Left            =   1080
            TabIndex        =   79
            Top             =   1680
            Width           =   975
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
            Left            =   1080
            TabIndex        =   78
            Top             =   2160
            Width           =   975
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
            Left            =   1080
            TabIndex        =   77
            Top             =   2400
            Width           =   975
         End
         Begin VB.TextBox txtMBMStatus 
            Height          =   285
            Left            =   1920
            TabIndex        =   76
            Top             =   240
            Width           =   375
         End
         Begin MSMask.MaskEdBox txtExpDate 
            Height          =   315
            Left            =   1080
            TabIndex        =   85
            Top             =   1440
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
            Left            =   1080
            TabIndex        =   86
            Top             =   960
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
            Left            =   120
            TabIndex        =   97
            Top             =   2160
            Width           =   735
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
            Left            =   120
            TabIndex        =   96
            Top             =   2400
            Width           =   615
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
            Left            =   120
            TabIndex        =   95
            Top             =   1920
            Visible         =   0   'False
            Width           =   855
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
            TabIndex        =   94
            Top             =   240
            Width           =   975
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
            Left            =   120
            TabIndex        =   93
            Top             =   1680
            Width           =   975
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
            Left            =   120
            TabIndex        =   92
            Top             =   1440
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
            Left            =   120
            TabIndex        =   91
            Top             =   960
            Width           =   975
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
            Left            =   120
            TabIndex        =   90
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label1 
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
            TabIndex        =   89
            Top             =   480
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
            Left            =   120
            TabIndex        =   88
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label102 
            Caption         =   "Status"
            Height          =   255
            Left            =   1560
            TabIndex        =   87
            Top             =   240
            Width           =   375
         End
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
         Left            =   4920
         MaxLength       =   6
         TabIndex        =   74
         Top             =   360
         Width           =   495
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
         Left            =   4920
         MaxLength       =   6
         TabIndex        =   73
         Top             =   600
         Width           =   495
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
         Left            =   2640
         MaxLength       =   6
         TabIndex        =   72
         Top             =   360
         Width           =   495
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
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   71
         Top             =   360
         Width           =   495
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
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   70
         Top             =   600
         Width           =   495
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
         Left            =   720
         MaxLength       =   6
         TabIndex        =   69
         Top             =   360
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
         Left            =   720
         MaxLength       =   6
         TabIndex        =   68
         Top             =   600
         Width           =   495
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
         Height          =   300
         Left            =   3960
         MaxLength       =   6
         TabIndex        =   67
         Top             =   360
         Width           =   495
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
         Left            =   3960
         MaxLength       =   6
         TabIndex        =   66
         Top             =   600
         Width           =   495
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
         Left            =   2640
         MaxLength       =   6
         TabIndex        =   65
         Top             =   600
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
         Height          =   315
         Left            =   6360
         MaxLength       =   6
         TabIndex        =   64
         Top             =   360
         Visible         =   0   'False
         Width           =   495
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
         Left            =   5400
         MaxLength       =   6
         TabIndex        =   63
         Top             =   360
         Visible         =   0   'False
         Width           =   495
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
         Left            =   5880
         MaxLength       =   6
         TabIndex        =   62
         Top             =   360
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
         Left            =   5400
         MaxLength       =   6
         TabIndex        =   61
         Top             =   600
         Visible         =   0   'False
         Width           =   495
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
         Left            =   5880
         MaxLength       =   6
         TabIndex        =   60
         Top             =   600
         Visible         =   0   'False
         Width           =   495
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
         Left            =   6360
         MaxLength       =   6
         TabIndex        =   59
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox A29 
         Height          =   285
         Left            =   6480
         MaxLength       =   6
         TabIndex        =   58
         Top             =   3120
         Width           =   495
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
         Left            =   1485
         MaxLength       =   6
         TabIndex        =   57
         Top             =   1320
         Width           =   495
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
         Height          =   285
         Left            =   3600
         MaxLength       =   6
         TabIndex        =   56
         Top             =   1260
         Width           =   495
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
         Left            =   3600
         MaxLength       =   6
         TabIndex        =   55
         Top             =   1440
         Width           =   495
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
         Left            =   3600
         MaxLength       =   6
         TabIndex        =   54
         Top             =   1695
         Width           =   495
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
         Left            =   3600
         MaxLength       =   6
         TabIndex        =   53
         Top             =   1950
         Width           =   495
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
         Left            =   3600
         MaxLength       =   6
         TabIndex        =   52
         Top             =   2220
         Width           =   495
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
         Left            =   6480
         MaxLength       =   6
         TabIndex        =   51
         Top             =   1275
         Width           =   495
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
         Left            =   6480
         MaxLength       =   6
         TabIndex        =   50
         Top             =   1515
         Width           =   495
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
         Left            =   6480
         MaxLength       =   6
         TabIndex        =   49
         Top             =   1785
         Width           =   495
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
         Left            =   6480
         MaxLength       =   6
         TabIndex        =   48
         Top             =   2025
         Width           =   495
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
         Left            =   6480
         MaxLength       =   6
         TabIndex        =   47
         Top             =   2280
         Width           =   495
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
         Left            =   3660
         MaxLength       =   6
         TabIndex        =   46
         Top             =   3120
         Width           =   450
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
         Left            =   3660
         MaxLength       =   6
         TabIndex        =   45
         Top             =   3360
         Width           =   450
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
         Left            =   3660
         MaxLength       =   6
         TabIndex        =   44
         Top             =   3600
         Width           =   450
      End
      Begin VB.TextBox txtLOthDesc 
         Height          =   285
         Left            =   3240
         MaxLength       =   50
         TabIndex        =   43
         Top             =   2505
         Width           =   855
      End
      Begin VB.TextBox txtIOthDesc 
         Height          =   285
         Left            =   5880
         MaxLength       =   50
         TabIndex        =   42
         Top             =   2535
         Width           =   1095
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
         Left            =   1485
         MaxLength       =   6
         TabIndex        =   41
         Top             =   1560
         Width           =   495
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
         Left            =   1485
         MaxLength       =   6
         TabIndex        =   40
         Top             =   1815
         Width           =   495
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
         Left            =   1485
         MaxLength       =   6
         TabIndex        =   39
         Top             =   2055
         Width           =   495
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
         Left            =   1485
         MaxLength       =   6
         TabIndex        =   38
         Top             =   2280
         Width           =   495
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
         Left            =   1485
         MaxLength       =   6
         TabIndex        =   37
         Top             =   2565
         Width           =   495
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
         Left            =   1485
         MaxLength       =   6
         TabIndex        =   36
         Top             =   2850
         Width           =   495
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
         Left            =   1485
         MaxLength       =   6
         TabIndex        =   35
         Top             =   3090
         Width           =   495
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
         Left            =   1485
         MaxLength       =   6
         TabIndex        =   34
         Top             =   3360
         Width           =   495
      End
      Begin VB.TextBox txtPOthDesc 
         Height          =   285
         Left            =   1005
         MaxLength       =   50
         TabIndex        =   33
         Top             =   3600
         Width           =   975
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
         Left            =   3660
         MaxLength       =   6
         TabIndex        =   32
         Top             =   3840
         Width           =   450
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
         Left            =   120
         TabIndex        =   163
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "X-Ray "
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
         TabIndex        =   162
         Top             =   1335
         Width           =   615
      End
      Begin VB.Label Label73 
         Caption         =   "Charges Covered"
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
         Left            =   120
         TabIndex        =   161
         Top             =   120
         Width           =   1455
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
         Left            =   2160
         TabIndex        =   160
         Top             =   3600
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
         Left            =   2160
         TabIndex        =   159
         Top             =   3360
         Width           =   735
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
         Left            =   2160
         TabIndex        =   158
         Top             =   3105
         Width           =   735
      End
      Begin VB.Label Label90 
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
         Height          =   180
         Left            =   4320
         TabIndex        =   157
         Top             =   2320
         Width           =   1095
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
         Left            =   4320
         TabIndex        =   156
         Top             =   1050
         Width           =   1455
      End
      Begin VB.Label Label87 
         Caption         =   "DTP && Polio && HIB"
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
         TabIndex        =   155
         Top             =   1560
         Width           =   1455
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
         Left            =   4320
         TabIndex        =   154
         Top             =   1800
         Width           =   975
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
         Left            =   4320
         TabIndex        =   153
         Top             =   2075
         Width           =   1095
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
         Left            =   8880
         TabIndex        =   152
         Top             =   1920
         Visible         =   0   'False
         Width           =   855
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
         Left            =   8880
         TabIndex        =   151
         Top             =   2040
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label63 
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
         Left            =   120
         TabIndex        =   150
         Top             =   1605
         Width           =   735
      End
      Begin VB.Label Label58 
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
         Left            =   120
         TabIndex        =   149
         Top             =   1875
         Width           =   375
      End
      Begin VB.Label Label20 
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
         Left            =   120
         TabIndex        =   148
         Top             =   2115
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "Nebulizer"
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
         TabIndex        =   147
         Top             =   2355
         Width           =   855
      End
      Begin VB.Label Label31 
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
         Left            =   120
         TabIndex        =   146
         Top             =   2835
         Width           =   615
      End
      Begin VB.Label Label27 
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
         Left            =   120
         TabIndex        =   145
         Top             =   2595
         Width           =   495
      End
      Begin VB.Label Label26 
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
         Height          =   255
         Left            =   120
         TabIndex        =   144
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label13 
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
         Left            =   120
         TabIndex        =   143
         Top             =   3075
         Width           =   735
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
         Left            =   2160
         TabIndex        =   142
         Top             =   1280
         Width           =   975
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
         Left            =   2160
         TabIndex        =   141
         Top             =   1040
         Width           =   1455
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
         Left            =   2160
         TabIndex        =   140
         Top             =   1520
         Width           =   975
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
         Left            =   2160
         TabIndex        =   139
         Top             =   1760
         Width           =   735
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
         Left            =   2160
         TabIndex        =   138
         Top             =   2010
         Width           =   855
      End
      Begin VB.Label Label53 
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
         Height          =   255
         Left            =   2160
         TabIndex        =   137
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label32 
         Caption         =   "DTP && Polio"
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
         TabIndex        =   136
         Top             =   1305
         Width           =   1575
      End
      Begin VB.Label Label21 
         Caption         =   "601"
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
         Left            =   3240
         TabIndex        =   135
         Top             =   3165
         Width           =   375
      End
      Begin VB.Label Label22 
         Caption         =   "602"
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
         Left            =   3240
         TabIndex        =   134
         Top             =   3420
         Width           =   255
      End
      Begin VB.Label Label24 
         Caption         =   "603"
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
         Left            =   3240
         TabIndex        =   133
         Top             =   3660
         Width           =   255
      End
      Begin VB.Label Label46 
         Caption         =   "301"
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
         Left            =   1080
         TabIndex        =   132
         Top             =   1365
         Width           =   375
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         Caption         =   "307"
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
         Left            =   960
         TabIndex        =   131
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         Caption         =   "308"
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
         Left            =   960
         TabIndex        =   130
         Top             =   3130
         Width           =   375
      End
      Begin VB.Label Label56 
         Caption         =   "302"
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
         Left            =   1080
         TabIndex        =   129
         Top             =   1635
         Width           =   375
      End
      Begin VB.Label Label57 
         Caption         =   "303"
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
         Left            =   1080
         TabIndex        =   128
         Top             =   1875
         Width           =   375
      End
      Begin VB.Label Label59 
         Caption         =   "304"
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
         Left            =   1080
         TabIndex        =   127
         Top             =   2115
         Width           =   375
      End
      Begin VB.Label Label80 
         Caption         =   "305"
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
         Left            =   1080
         TabIndex        =   126
         Top             =   2355
         Width           =   375
      End
      Begin VB.Label Label81 
         Alignment       =   1  'Right Justify
         Caption         =   "306"
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
         Left            =   1080
         TabIndex        =   125
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label Label92 
         Caption         =   "404"
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
         Left            =   3240
         TabIndex        =   124
         Top             =   1995
         Width           =   375
      End
      Begin VB.Label Label93 
         Caption         =   "403"
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
         Left            =   3240
         TabIndex        =   123
         Top             =   1740
         Width           =   375
      End
      Begin VB.Label Label94 
         Caption         =   "402"
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
         Left            =   3240
         TabIndex        =   122
         Top             =   1485
         Width           =   375
      End
      Begin VB.Label Label95 
         Caption         =   "401"
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
         Left            =   3240
         TabIndex        =   121
         Top             =   1245
         Width           =   375
      End
      Begin VB.Label Label96 
         Caption         =   "505"
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
         Left            =   6120
         TabIndex        =   120
         Top             =   2325
         Width           =   375
      End
      Begin VB.Label Label97 
         Caption         =   "504"
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
         Left            =   6120
         TabIndex        =   119
         Top             =   2085
         Width           =   375
      End
      Begin VB.Label Label98 
         Caption         =   "503"
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
         Left            =   6120
         TabIndex        =   118
         Top             =   1815
         Width           =   375
      End
      Begin VB.Label Label99 
         Caption         =   "502"
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
         Left            =   6120
         TabIndex        =   117
         Top             =   1575
         Width           =   375
      End
      Begin VB.Label Label100 
         Caption         =   "501"
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
         Left            =   6120
         TabIndex        =   116
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         Caption         =   "309"
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
         Left            =   960
         TabIndex        =   115
         Top             =   3375
         Width           =   375
      End
      Begin VB.Label Label91 
         Caption         =   "405"
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
         Left            =   3240
         TabIndex        =   114
         Top             =   2260
         Width           =   375
      End
      Begin VB.Label Label101 
         Caption         =   "Co-Payment"
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
         TabIndex        =   113
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label103 
         Caption         =   "Oth Desc"
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
         TabIndex        =   112
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label104 
         Caption         =   "Others Desc"
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
         TabIndex        =   111
         Top             =   2560
         Width           =   1095
      End
      Begin VB.Label Label105 
         Caption         =   "Others Desc"
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
         Left            =   2160
         TabIndex        =   110
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label106 
         Caption         =   "Co-Payment"
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
         Left            =   4320
         TabIndex        =   109
         Top             =   2850
         Width           =   1455
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
         Left            =   2160
         TabIndex        =   108
         Top             =   2850
         Width           =   1455
      End
      Begin VB.Label Label108 
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
         Left            =   3240
         TabIndex        =   107
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label74 
         Caption         =   "GP Fee + LT Med."
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
         TabIndex        =   106
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label109 
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
         Left            =   3240
         TabIndex        =   105
         Top             =   680
         Width           =   735
      End
      Begin VB.Label Label77 
         Caption         =   "SP Fee + LT Med."
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
         TabIndex        =   104
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label115 
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
         Left            =   2320
         TabIndex        =   103
         Top             =   400
         Width           =   375
      End
      Begin VB.Label Label116 
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
         Left            =   1320
         TabIndex        =   102
         Top             =   405
         Width           =   375
      End
      Begin VB.Label Label117 
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
         Left            =   2320
         TabIndex        =   101
         Top             =   680
         Width           =   315
      End
      Begin VB.Label Label118 
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
         Left            =   1320
         TabIndex        =   100
         Top             =   645
         Width           =   375
      End
      Begin VB.Label Label28 
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
         Left            =   2160
         TabIndex        =   99
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Label33 
         Caption         =   "604"
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
         Left            =   3240
         TabIndex        =   98
         Top             =   3880
         Width           =   255
      End
   End
   Begin VB.Frame FrameDis 
      Enabled         =   0   'False
      Height          =   1550
      Left            =   8880
      TabIndex        =   11
      Top             =   360
      Width           =   3015
      Begin VB.TextBox txtMepLmt 
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
         Left            =   2320
         TabIndex        =   21
         Top             =   1160
         Width           =   615
      End
      Begin VB.TextBox txtSpecLmt 
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
         Left            =   2320
         TabIndex        =   20
         Top             =   920
         Width           =   615
      End
      Begin VB.TextBox txtDentLmt 
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
         Left            =   2320
         TabIndex        =   19
         Top             =   680
         Width           =   615
      End
      Begin VB.TextBox TxtVisitLimit 
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
         Left            =   2320
         TabIndex        =   18
         Top             =   420
         Width           =   615
      End
      Begin VB.TextBox TxtCoPay 
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
         Left            =   920
         TabIndex        =   17
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtLTM 
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
         Left            =   920
         TabIndex        =   16
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtDentStatus 
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
         Left            =   720
         TabIndex        =   15
         Top             =   1200
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox TxtBalLimit 
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
         Left            =   2320
         TabIndex        =   14
         Top             =   180
         Width           =   615
      End
      Begin VB.TextBox TxtClnContract 
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
         Left            =   920
         TabIndex        =   13
         Top             =   480
         Width           =   615
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
         Height          =   315
         Left            =   920
         TabIndex        =   12
         Top             =   220
         Width           =   615
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
         Left            =   120
         TabIndex        =   30
         Top             =   260
         Width           =   735
      End
      Begin VB.Label Label40 
         Caption         =   "Card Type"
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
         TabIndex        =   29
         Top             =   560
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000013&
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
         Left            =   1620
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label34 
         Caption         =   "Vst Lmt"
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
         Left            =   1620
         TabIndex        =   27
         Top             =   490
         Width           =   735
      End
      Begin VB.Label Label68 
         Caption         =   "Co-Pay"
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
         Height          =   210
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label69 
         BackColor       =   &H80000013&
         Caption         =   "LTM"
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
         Height          =   135
         Left            =   120
         TabIndex        =   25
         Top             =   820
         Width           =   375
      End
      Begin VB.Label Label42 
         Caption         =   "Dent Lmt"
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
         Left            =   1620
         TabIndex        =   24
         Top             =   740
         Width           =   735
      End
      Begin VB.Label Label39 
         Caption         =   "Spec Lmt"
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
         Left            =   1620
         TabIndex        =   23
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label43 
         Caption         =   "MEP Lmt"
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
         Left            =   1620
         TabIndex        =   22
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.Frame FrameInvoice 
      Height          =   1125
      Left            =   5040
      TabIndex        =   0
      Top             =   360
      Width           =   3735
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
         Left            =   840
         MaxLength       =   20
         TabIndex        =   4
         Top             =   195
         Width           =   975
      End
      Begin VB.TextBox TxtTotInvoiceAmt 
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
         Left            =   2760
         MaxLength       =   20
         TabIndex        =   3
         Top             =   195
         Width           =   855
      End
      Begin VB.TextBox txtAllNoOfBill 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   2760
         MaxLength       =   20
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtAllBillAmt 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   2760
         MaxLength       =   20
         TabIndex        =   1
         Top             =   720
         Width           =   855
      End
      Begin MSMask.MaskEdBox txtInvDate 
         Height          =   300
         Left            =   840
         TabIndex        =   5
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
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
      Begin VB.Label Label18 
         Caption         =   "Inv. Amt"
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
         Left            =   2040
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label35 
         Caption         =   "Inv. No"
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
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label67 
         Caption         =   "Inv. Date"
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
         TabIndex        =   8
         Top             =   540
         Width           =   735
      End
      Begin VB.Label lblAllBillAmt 
         Caption         =   "Tot. Amt"
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
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   765
         Width           =   735
      End
      Begin VB.Label LblAllNoOfBill 
         Caption         =   "No.of.Bill"
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
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Top             =   525
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView lstInvoice 
      Height          =   1905
      Left            =   0
      TabIndex        =   247
      Top             =   6600
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   3360
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
      NumItems        =   468
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Batch No"
         Object.Width           =   1411
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
         Text            =   "Inv. No"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Bill No."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Inv. Date"
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
         Text            =   "OtherDiagCode"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(95) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   94
         Text            =   "OtherDiagind"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(96) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   95
         Text            =   "OthDiagDesc"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(97) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   96
         Text            =   "DiagDesc1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(98) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   97
         Text            =   "DiagDesc2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(99) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   98
         Text            =   "DiagDesc3"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(100) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   99
         Text            =   "DiagDesc4"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(101) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   100
         Text            =   "DiagDesc5"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(102) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   101
         Text            =   "Vision"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(103) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   102
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(104) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   103
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(105) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   104
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(106) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   105
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(107) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   106
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(108) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   107
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(109) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   108
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(110) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   109
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(111) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   110
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(112) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   111
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(113) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   112
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(114) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   113
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(115) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   114
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(116) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   115
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(117) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   116
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(118) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   117
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(119) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   118
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(120) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   119
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(121) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   120
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(122) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   121
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(123) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   122
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(124) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   123
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(125) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   124
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(126) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   125
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(127) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   126
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(128) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   127
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(129) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   128
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(130) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   129
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(131) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   130
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(132) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   131
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(133) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   132
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(134) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   133
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(135) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   134
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(136) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   135
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(137) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   136
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(138) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   137
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(139) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   138
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(140) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   139
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(141) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   140
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(142) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   141
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(143) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   142
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(144) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   143
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(145) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   144
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(146) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   145
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(147) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   146
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(148) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   147
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(149) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   148
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(150) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   149
         Object.Width           =   2540
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
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(198) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   197
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
         Object.Width           =   0
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
         Object.Width           =   0
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
         Object.Width           =   0
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
   End
   Begin MSComctlLib.ListView lstErr 
      Height          =   255
      Left            =   240
      TabIndex        =   248
      Top             =   6720
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
   Begin VB.Label lblInvoice 
      BackColor       =   &H00404040&
      Caption         =   " Invoice Batch Processing - Dental"
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
      Height          =   345
      Left            =   0
      TabIndex        =   249
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmInvDentalProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bDataStatus As Boolean
Dim bExit As Boolean
Dim intIndex As String
Dim EditFlag As Boolean
Dim ScanCode As String
Dim AA, CC As Integer
Dim strPayTo, strPayee As String
Dim strMBMno As String
Dim IntInvAmt As Double
Dim CASENumber, NewCaseNo, MembershipNo, CovPersonID As String
Dim strPolNo, strPatientName, strMBMDataStatus, strPOlStatus As String
Dim strEffdate, strExpDate, strClnPln, strClnIns, strPatientDOB As String
Dim strPayor, strPatientID, strPatientIC, strBalLimit, strVisitLimit, strCLNContract As String
Dim strMBMCLNLTM, strHLTCode As String
Dim strVisitIND, strAnnualInd, strCoPay As String
Dim MBM As New ADODB.Recordset
Dim lstErrRec As ListItem
Dim VPlan, VExpDate, VEffDate, VDiagCode, VDiagIND As String
Dim SuppLmtType As String
Dim Z1, Z2, Z3, Z4, Z5, Z6, Z7, Z8 As Double
Dim OPBatchNo As String
Dim TempOPVar As String
Dim ReasonOtherIND As String
Dim StrDentStatus, StrDentLmt As String
Dim StrSpecLmt, StrMepLmt As String
Public lstRec As ListItem

Private Sub A1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub A3_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A30_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
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

Private Sub CboClinic_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboClinic.Text, CboClinic.SelStart) + Chr(KeyAscii) + Mid(CboClinic.Text, CboClinic.SelStart + CboClinic.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboClinic.ListCount - 1
       
        If NewText = LCase(Left(CboClinic.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboClinic.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboClinic.Text = ValidValue
            CboClinic.SelStart = 0
            CboClinic.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CmdAddRec_Click()
    frmWorksheet.InvoiceClick = 1
    Call ChkData
End Sub

Private Sub CmdClearInv_Click()
    Call ClearEntries
    Call ClearMBMInfo
    FrameMBMCLN.Enabled = True
    EditFlag = False
    Call EntriesEnabled(False)
    CmdSave.Visible = False
    CmdVerify.Visible = True
End Sub

Private Sub CmdClearAll_Click()
    Call ClearEntries
    Call ClearMBMInfo
    txtInvDate = "__/__/____"
    txtInvNo = ""
    CboClinic.Clear
    CboClinic.Text = ""
    TxtClinicCode = ""
    txtAllBillAmt = 0
    txtAllNoOfBill = 0
    CmdMBMInfo.Enabled = False
    Call EntriesEnabled(False)
    lstInvoice.ListItems.Clear
    lstErr.ListItems.Clear
    EditFlag = False
    FrameMBMCLN.Enabled = True
End Sub

Private Sub cmdCompute_Click()
    frmWorksheet.InvoiceClick = 2
    frmWorksheet.Show
End Sub

Private Sub CmdEdit_Click()
    If Trim(txtMBMNo) <> "" Then
        CmdEdit.Visible = False
        CmdEdit.Enabled = False
        CmdAddRec.Enabled = False
        CmdUpdate.Visible = True
        CmdUpdate.Enabled = True
        Call EntriesEnabled(True)
        FrameMBMCLN.Enabled = False
        EditFlag = True
        txtBillNo.SetFocus
    End If
    CmdSave.Visible = False
    CmdVerify.Visible = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Public Sub CmdGo_Click()
Dim ZZ As String
Dim QQ As Integer
    If lstInvoice.ListItems.count > 0 Then
        For QQ = 1 To lstInvoice.ListItems.count
            If Trim(lstInvoice.ListItems(QQ).SubItems(1)) <> "" Then
                lstInvoice.ListItems.Clear
                Exit For
            End If
        Next QQ
    End If
    txtMBMNo = ChkStr(txtMBMNo)
    TempOPVar = "OP"
    If Trim(txtMBMNo) <> "" Then
        If Mid(Trim(txtMBMNo), 1, 1) = "Z" Then
            TempOPVar = "ZOP"
        End If
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
            CmdAddRec.Enabled = True
            Call ClearEntries
            CmdEdit.Visible = True
            Call EntriesEnabled(True)
            txtTabMBMNo = MembershipNo
            txtMBMName = strPatientName
            TxtPayor = strPayor
            txtPlanCode = strClnPln
            TxtInsType = strClnIns
            txtPatientName = strPatientName
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
            txtLTM = strMBMCLNLTM
            txtDentStatus = StrDentStatus
            txtDentLmt = StrDentLmt
            txtSpecLmt = StrSpecLmt
            txtMepLmt = StrMepLmt
            'TxtClinicCode.Enabled = True
            'TxtClinicCode.SetFocus
        End If
    End If
End Sub

Private Sub CmdMBMInfo_Click()
    FrmINVMBMinfo.Source = 1
    FrmINVMBMinfo.Show vbModal
End Sub

Private Sub CmdPrint_Click()
    If lstInvoice.ListItems.count > 0 Then
        Call ExportToExcelError(lstInvoice)
    End If
End Sub

Private Sub CmdRemove_Click()
Dim RecRemove
    
    RecRemove = MsgBox("Are you sure?", vbYesNo + vbExclamation)
    If RecRemove = vbYes Then
        Call RemoveRecFrListV
        Call ClearEntries
        Call ClearMBMInfo
    End If
End Sub

Private Sub CmdSave_Click()
Dim RecSaved

    If lstInvoice.ListItems.count <> 0 Then
        RecSaved = MsgBox("Are you sure ? ", vbYesNo + vbExclamation)
        If RecSaved = vbYes Then
            OPBatchNo = GenerateOPSeq(TempOPVar)
            For AA = 1 To lstInvoice.ListItems.count
                Call SaveRecord
            Next AA
            
            MsgBox "Record Saved ! ", vbOKOnly
    
            CmdSave.Enabled = False
            CmdSave.Visible = False
            CmdEdit.Enabled = False
            CmdVerify.Visible = True
            CmdVerify.Enabled = True
        End If
    End If
End Sub

Private Sub cmdSearch_Click()
    bDataStatus = True
    frmSearchMBM.Source = 1
    frmSearchMBM.Show vbModal
End Sub

Private Sub cmdSearchICD1_Click()
    frmSearchICD.Source = 1
    frmSearchICD.Show vbModal
End Sub

Private Sub cmdSearchICD2_Click()
    frmSearchICD.Source = 2
    frmSearchICD.Show vbModal
End Sub

Private Sub cmdSearchICD3_Click()
    frmSearchICD.Source = 3
    frmSearchICD.Show vbModal
End Sub

Private Sub cmdSearchICD4_Click()
    frmSearchICD.Source = 4
    frmSearchICD.Show vbModal
End Sub

Private Sub cmdSearchICD5_Click()
Dim TreatRec As New ADODB.Recordset
Dim TreatSql As String

    frmSearchICD.Source = 5
    frmSearchICD.Show vbModal
    
    TxtDiagnosisOth.Enabled = False
    TxtDiagDesc5.Visible = True
    If Trim(TxtDiagnosis5) <> "" Then
        TreatSql = "Select ReasonOtherDiag from TreatmentReason WHERE ReasonCode='" & Trim(TxtDiagnosis5) & "'"
        TreatRec.Open TreatSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        If TreatRec.RecordCount Then
            ReasonOtherIND = IIf(Len(Trim(TreatRec!ReasonOtherDiag)), Trim(TreatRec!ReasonOtherDiag), "N")
            If ReasonOtherIND = "Y" Then
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
End Sub

Private Sub CmdTotInv_Click()
Dim lstCnt As Integer
    txtAllBillAmt = 0
    txtAllNoOfBill = 0
    If lstInvoice.ListItems.count > 0 Then
        txtAllNoOfBill.Visible = True
        txtAllBillAmt.Visible = True
        txtAllNoOfBill = lstInvoice.ListItems.count
        For lstCnt = 1 To lstInvoice.ListItems.count
            txtAllBillAmt = txtAllBillAmt + CDbl(lstInvoice.ListItems(lstCnt).SubItems(13))
        Next lstCnt
    End If
End Sub

Private Sub CmdUpdate_Click()
    frmWorksheet.InvoiceClick = 1
    Call ChkData
End Sub

Private Sub CmdVerify_Click()
Dim ChkSql As String
Dim ChKItemRec As New ADODB.Recordset
Dim TotBillAmt As Double
    
    TotBillAmt = 0
    lstErr.ListItems.Clear
    If lstInvoice.ListItems.count = 0 Then
        MsgBox "No Record to be Verified ! ", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        
        For CC = 1 To lstInvoice.ListItems.count
            TotBillAmt = TotBillAmt + lstInvoice.ListItems(CC).SubItems(13)
            VPlan = lstInvoice.ListItems(CC).SubItems(64)
            
            VEffDate = lstInvoice.ListItems(CC).SubItems(78)
            VExpDate = lstInvoice.ListItems(CC).SubItems(79)
            If Len(Trim(lstInvoice.ListItems(CC).SubItems(62))) Then
                VDiagCode = Trim(lstInvoice.ListItems(CC).SubItems(62))
                Call VerifyDiagnosis
            End If
            If Len(Trim(lstInvoice.ListItems(CC).SubItems(63))) Then
                VDiagCode = Trim(lstInvoice.ListItems(CC).SubItems(63))
                Call VerifyDiagnosis
            End If
            If Len(Trim(lstInvoice.ListItems(CC).SubItems(64))) Then
                VDiagCode = Trim(lstInvoice.ListItems(CC).SubItems(64))
                Call VerifyDiagnosis
            End If
            If Len(Trim(lstInvoice.ListItems(CC).SubItems(65))) Then
                VDiagCode = Trim(lstInvoice.ListItems(CC).SubItems(65))
                Call VerifyDiagnosis
            End If
            If Len(Trim(lstInvoice.ListItems(CC).SubItems(93))) Then
                VDiagCode = Trim(lstInvoice.ListItems(CC).SubItems(93))
                Call VerifyDiagnosis
            End If
            If Trim(lstInvoice.ListItems(CC).SubItems(6)) = "" Then
                Call ErrRpt
                lstErrRec.SubItems(9) = "Invoice No Can't be null ! "
            End If
            If IsDate(lstInvoice.ListItems(CC).SubItems(8)) Then
                If CDate(lstInvoice.ListItems(CC).SubItems(8)) < CDate(VEffDate) Then
                    Call ErrRpt
                    lstErrRec.SubItems(9) = "Invoice Date MUST BE GREATER than Clinical Effective date ! "
                End If
                If CDate(lstInvoice.ListItems(CC).SubItems(8)) > CDate(VExpDate) + 14 Then
                    Call ErrRpt
                    lstErrRec.SubItems(9) = "Invoice date IS GREATER THAN Clinical Expiry date + 14 days ! "
                End If
    '           If IsDate(lstInvoice.ListItems(CC).SubItems(8)) Then
                If (CDate(Date) - CDate(lstInvoice.ListItems(CC).SubItems(8))) > 45 Then
                    Call ErrRpt
                    lstErrRec.SubItems(9) = "45 Days LATE ! "
                End If
            End If
            If IsDate(lstInvoice.ListItems(CC).SubItems(9)) Then
                If CDate(lstInvoice.ListItems(CC).SubItems(9)) < CDate(VEffDate) Then
                    Call ErrRpt
                    lstErrRec.SubItems(9) = "Visit date MUST BE AFTER the Effective date ! "
                End If
                If CDate(lstInvoice.ListItems(CC).SubItems(9)) > CDate(VExpDate) + 14 Then
                    Call ErrRpt
                    lstErrRec.SubItems(9) = "Visit date IS GREATER THAN Expiry date + 14 days ! "
                End If
            ElseIf lstInvoice.ListItems(CC).SubItems(9) <> "__/__/____" Then
            
            End If
            If IsNumeric(lstInvoice.ListItems(CC).SubItems(70)) Then ' Bill Amt > Balance Limit
                If CDbl(lstInvoice.ListItems(CC).SubItems(13)) > CDbl(lstInvoice.ListItems(CC).SubItems(70)) Then
                    Call ErrRpt
                    lstErrRec.SubItems(9) = "Bill Amt > Balance Limit ! "
                End If
            End If
            
            If IsNumeric(lstInvoice.ListItems(CC).SubItems(71)) Then 'Bill Amt > Visit Limit
                If CDbl(lstInvoice.ListItems(CC).SubItems(13)) > CDbl(lstInvoice.ListItems(CC).SubItems(71)) Then
                    Call ErrRpt
                    lstErrRec.SubItems(9) = "Bill Amt > Visit Limit ! "
                End If
            End If
            
            ChkSql = "SELECT Catcode, Amount FROM CheckItems WHERE CatCode = 'CSMD'" & _
                    " and CatName ='" & lstInvoice.ListItems(CC).SubItems(72) & "'"
            ChKItemRec.Open ChkSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            If ChKItemRec.RecordCount Then
                If IsNumeric(ChKItemRec!Amount) Then
                    If CDbl(lstInvoice.ListItems(CC).SubItems(13)) > CDbl(ChKItemRec!Amount) Then
                        Call ErrRpt
                        lstErrRec.SubItems(9) = "Bill Amount > Consultation Medication - " & lstInvoice.ListItems(CC).SubItems(72) & "Card"
                    End If
                End If
            End If
            ChKItemRec.Close
            Set ChKItemRec = Nothing
        Next CC
        TxtTotInvoiceAmt = IIf(IsNumeric(TxtTotInvoiceAmt), TxtTotInvoiceAmt, 0)
        If TotBillAmt <> TxtTotInvoiceAmt Then
            Set lstErrRec = lstErr.ListItems.Add(, , " ")
            lstErrRec.SubItems(1) = " "
            lstErrRec.SubItems(2) = " "
            lstErrRec.SubItems(3) = " "
            lstErrRec.SubItems(4) = " "
            lstErrRec.SubItems(5) = " "
            lstErrRec.SubItems(6) = " "
            lstErrRec.SubItems(7) = "Invoice Amt = " & CDbl(TxtTotInvoiceAmt)
            lstErrRec.SubItems(8) = "Total Bill Amt = " & CDbl(TotBillAmt)
            lstErrRec.SubItems(9) = "Invoice Amt is different from Total Bill Amt"
            CmdVerify.Visible = True
            CmdSave.Visible = False
            CmdSave.Enabled = False
        Else
            CmdVerify.Visible = False
            CmdSave.Visible = True
            CmdSave.Enabled = True
        End If
        
        If lstErr.ListItems.count > 0 Then
            Call ExportToExcelError(lstErr)
        Else
            MsgBox "Verifed Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
        End If
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
    CmdSave.Enabled = False
    CmdSave.Visible = False
    intIndex = 0
'sl    If CheckAccess(117) = True Then
        strPayTo = "C"
'sl    ElseIf CheckAccess(118) = True Then
'sl        strPayTo = "M"
'sl    End If
End Sub

Private Sub ClearEntries()
   On Error Resume Next
   Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            If ctl.Name <> "txtMBMNo" And ctl.Name <> "txtInvNo" And ctl.Name <> "TxtClinicCode" And ctl.Name <> "TxtTotInvoiceAmt" Then
               ctl.Text = ""
            End If
        ElseIf TypeOf ctl Is ComboBox Then
        
'        ElseIf TypeOf ctl Is ListView Then
'            ctl.ListItems.Clear
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
    Z1 = 0
    Z2 = 0
    Z3 = 0
    Z4 = 0
    Z5 = 0
    Z6 = 0
    Z7 = 0
    Z8 = 0
    ReasonOtherIND = ""
    TxtDiagnosisOth.Enabled = False
    TxtDiagDesc5.Visible = True
    CmdUpdate.Visible = False
    CmdEdit.Visible = True
    CmdEdit.Enabled = True
End Sub

Private Sub MBMInfo()
Dim MBMRec As New ADODB.Recordset
Dim CLNOth As New ADODB.Recordset
Dim MBMSql, CLNSql, COVSql, COVLmtSql As String
Dim COVRec As New ADODB.Recordset
Dim Cnt As Integer
Dim CISPln As New ADODB.Recordset
Dim CISAnnLMt As New ADODB.Recordset
Dim CISPlnSql, CISAnnLmtSql As String
Dim COVLmt As New ADODB.Recordset
Dim CoPayment As New ADODB.Recordset
Dim CoSql As String
Dim TmpMBMNo As String
Dim TempCovID As String

    If Len(Trim(MembershipNo)) Then
        MBMSql = "SELECT MBMNumber, MBMName, MBMICBcPp, PAYCode, MBMCLNStatus, MBMCOVID,  " & _
                " MBMDOB, MBMCLNDataStatus, HLTCode FROM VIEW_MBM_Clinical WHERE MBMNumber='" & Trim(MembershipNo) & " ' AND MBMCLNDataStatus <> 'D' " & _
                " AND MBMCLNStatus <> 'D' AND (HSCLNInd = 'C' OR HSCLNInd = 'M' OR HSCLNInd = 'B')"
        MBMRec.Open MBMSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If MBMRec.RecordCount <= 0 Then
            MsgBox "Membership No not found ! ", vbOKOnly + vbCritical, DialogBoxTitle
           
            TmpMBMNo = MembershipNo & "-" & CovPersonID
            MembershipNo = ""
            CovPersonID = ""
            Call ClearEntries
            Call ClearMBMInfo
            txtMBMNo = TmpMBMNo
            CmdAddRec.Enabled = False
            CmdMBMInfo.Enabled = False
            txtMBMNo.SetFocus
            
        Else
            SuppLmtType = ""
            CLNSql = "SELECT MBMNumber, MBMPolNumber, MBMPolEffDate, MBMPolExpDate,MBMCLNBalLmt,MBMCLNVisitLmt, " & _
                    "INSType, PLNType,MBMCLNLTM, PAYCode, " & _
                    " MBMDentalLmt, MBMSpecsLmt, MBMMEPLmt FROM MBMCLNOthers WHERE MBMNumber='" & Trim(MembershipNo) & "'"
            CLNOth.Open CLNSql, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            CISPlnSql = "SELECT AnnLmtIND, VisLmtIND, ClnContract, PDentalStatus, SDentalStatus, PDentalLimit, SDentalLimit " & _
                        " from PLN WHERE PLNCode ='" & Trim(CLNOth!PLNType) & "' and HLTCode='" & Trim(MBMRec!HLTCode) & "'"
            CISPln.Open CISPlnSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            
            CoSql = "SELECT * FROM CoPaymentMBM WHERE HLTType='" & MBMRec!HLTCode & "' And PlnCode='" & CLNOth!PLNType & "'"
            CoPayment.Open CoSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            
            If Trim(CISPln!AnnLmtIND) = "Y" Then
                CISAnnLmtSql = "SELECT SuppLimitStatus FROM AnnualLimit WHERE PLNCode='" & Trim(CLNOth!PLNType) & "' and HLTCode='" & Trim(MBMRec!HLTCode) & "' AND INSCode='" & Trim(CLNOth!INSType) & "'"
                CISAnnLMt.Open CISAnnLmtSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            
                SuppLmtType = CISAnnLMt!SuppLimitStatus
                CISAnnLMt.Close
                Set CISAnnLMt = Nothing
            End If
            strVisitIND = IIf(Len(CISPln!AnnLmtIND), CISPln!AnnLmtIND, "")
            strAnnualInd = IIf(Len(CISPln!VisLmtIND), CISPln!VisLmtIND, "")
            If Trim(CISPln!AnnLmtIND) = "N" Then
                strBalLimit = "FULL"
            Else
                strBalLimit = CLNOth!MBMCLNBalLmt
            End If
            If Trim(CISPln!VisLmtIND) = "N" Then
                strVisitLimit = "FULL"
            Else
                strVisitLimit = CLNOth!MBMCLNVisitLmt
            End If
            StrDentStatus = IIf(Len(Trim(CISPln!PDentalStatus)), Trim(CISPln!PDentalStatus), "")
            
            strCLNContract = CISPln!ClnContract
            CISPln.Close
            Set CISPln = Nothing
            
            If CoPayment.RecordCount Then
                TxtCoPay = IIf(CoPayment!CoAmount > 0, CoPayment!CoAmount, 0)
                A29 = IIf(TxtCoPay > 0, TxtCoPay, 0)
            End If
            CoPayment.Close
            Set CoPayment = Nothing
            
            strMBMno = MBMRec!MBMNumber
            strPayor = CLNOth!PAYCode
            strHLTCode = MBMRec!HLTCode
            strClnPln = CLNOth!PLNType
            strClnIns = CLNOth!INSType
            strEffdate = CLNOth!MBMPolEffDate
            strExpDate = CLNOth!MBMPolExpDate
            strPolNo = CLNOth!MBMPolNumber
            strMBMDataStatus = MBMRec!MBMCLNDataStatus
            strPOlStatus = MBMRec!MBMCLNStatus
            strPatientName = MBMRec!MBMName
            strPatientDOB = MBMRec!MBMDOB
            strPatientIC = MBMRec!MBMIcBcPp
            strMBMCLNLTM = IIf(Len(Trim(CLNOth!MBMCLNLTM)), CLNOth!MBMCLNLTM, "")
            StrDentLmt = IIf(Len(Trim(CLNOth!MBMDentalLmt)), CLNOth!MBMDentalLmt, "")
            StrSpecLmt = IIf(Len(Trim(CLNOth!MBMSpecsLmt)), CLNOth!MBMSpecsLmt, "")
            StrMepLmt = IIf(Len(Trim(CLNOth!MBMMEPLmt)), CLNOth!MBMMEPLmt, "")
            CmdAddRec.Enabled = True
            CmdMBMInfo.Enabled = True
            If CovPersonID <> "00" Then
                COVSql = "SELECT MBMCName, MBMCCoverID, MBMCICBCPPNo, MBMCDOB, MBMCCLNBalLmt, MBMCCLNVisitLmt FROM MBMCoveredPersons WHERE MBMCNumber='" & Trim(MembershipNo) & "' AND MBMCCoverID = '" & Trim(CovPersonID) & "'"
                COVRec.Open COVSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                If COVRec.RecordCount <= 0 Then
                    MsgBox "No such Cover Person ! ", vbOKOnly + vbCritical, DialogBoxTitle
                    txtMBMNo = MembershipNo & "-" & CovPersonID
                    MembershipNo = ""
                    CovPersonID = ""
                    Call ClearEntries
                    Call ClearMBMInfo
                    CmdAddRec.Enabled = False
                    txtMBMNo.SetFocus
                Else
                    TempCovID = Trim(CovPersonID) ' supp separate limit type 'C'
                    If SuppLmtType = "B" Then     ' supp shared limit type 'B'
                        TempCovID = "01"
                    End If
                    COVLmtSql = "SELECT MBMCCLNBalLmt, MBMCCLNVisitLmt, MBMCCLNDentalLmt, MBMCCLNSpecLmt, MBMCCLNMEPLmt " & _
                                "FROM MBMCoveredPersonsCLN WHERE MBMCNumber='" & Trim(MembershipNo) & "' AND MBMCCoverID = '" & Trim(TempCovID) & "'"
                    COVLmt.Open COVLmtSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                    strBalLimit = COVLmt!MBMCCLNBalLmt
                    strVisitLimit = COVLmt!MBMCCLNVisitLmt
                    StrDentLmt = IIf(Len(Trim(COVLmt!MBMCCLNDentalLmt)), COVLmt!MBMCCLNDentalLmt, "")
                    StrSpecLmt = IIf(Len(Trim(COVLmt!MBMCCLNSpecLmt)), COVLmt!MBMCCLNSpecLmt, "")
                    StrMepLmt = IIf(Len(Trim(COVLmt!MBMCCLNMEPLmt)), COVLmt!MBMCCLNMEPLmt, "")

                    COVLmt.Close
                    Set COVLmt = Nothing
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
    strCLNContract = ""
    strMBMCLNLTM = ""
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
    TxtPolNo = ""
    txtPolStatus = ""
    txtMBMStatus = ""
    txtdob = ""
    txtPatientIC = ""
    txtPatientName = ""
    TxtBalLimit = ""
    TxtVisitLimit = ""
    TxtClnContract = ""
    txtLTM = ""
End Sub
Private Sub ComboListing()
'Dim ClinicRec As New ADODB.Recordset
        
'    CboClinic.Clear
'    ClinicRec.Open "SELECT CLNCode, CLNName FROM CLN WHERE CLNPanel = 'Y'", medixmasterconn, adOpenKeyset, adLockOptimistic
'    Do Until ClinicRec.EOF
'        CboClinic.AddItem ClinicRec!CLNCode & " - " & ClinicRec!CLNName
'        ClinicRec.MoveNext
'    Loop
'    ClinicRec.Close
'    Set ClinicRec = Nothing
           
End Sub



Private Sub lstInvoice_Click()
 CmdRemove.Enabled = True
 '   If EditFlag = True Then
        If lstInvoice.ListItems.count Then
            CmdUpdate.Enabled = True
            
            txtMBMNo = ""
            txtMBMNo = lstInvoice.SelectedItem.SubItems(2) & "-" & lstInvoice.SelectedItem.SubItems(3)
            MembershipNo = lstInvoice.SelectedItem.SubItems(2)
            Call MBMInfo
            txtTabMBMNo = lstInvoice.SelectedItem.SubItems(2)
            TxtPayor = strPayor
            
            txtPlanCode = lstInvoice.SelectedItem.SubItems(67)
            TxtInsType = lstInvoice.SelectedItem.SubItems(68)
            txtMBMName = lstInvoice.SelectedItem.SubItems(4)
            txtEffDate = lstInvoice.SelectedItem.SubItems(78)
            txtExpDate = lstInvoice.SelectedItem.SubItems(79)
            TxtPolNo = lstInvoice.SelectedItem.SubItems(76)
            txtPolStatus = strPOlStatus
            txtMBMStatus = strMBMDataStatus
            txtdob = strPatientDOB
            txtPatientIC = lstInvoice.SelectedItem.SubItems(77)
            txtPatientName = lstInvoice.SelectedItem.SubItems(4)
            TxtBalLimit = lstInvoice.SelectedItem.SubItems(70)
            TxtVisitLimit = lstInvoice.SelectedItem.SubItems(71)
            TxtClnContract = lstInvoice.SelectedItem.SubItems(72)
            txtBillNo = lstInvoice.SelectedItem.SubItems(7)
            txtInvDate = lstInvoice.SelectedItem.SubItems(8)
            txtVisitDate = lstInvoice.SelectedItem.SubItems(9)
            txtVisitTime = lstInvoice.SelectedItem.SubItems(10)
            txtMC = lstInvoice.SelectedItem.SubItems(11)
            CboClinic.Text = lstInvoice.SelectedItem.SubItems(12)
            strPayTo = lstInvoice.SelectedItem.SubItems(80)
            strPayee = lstInvoice.SelectedItem.SubItems(81)
            A1 = lstInvoice.SelectedItem.SubItems(15)
            A2 = lstInvoice.SelectedItem.SubItems(16)
            A3 = lstInvoice.SelectedItem.SubItems(17)
            A4 = lstInvoice.SelectedItem.SubItems(18)
            A5 = lstInvoice.SelectedItem.SubItems(19)
            A6 = lstInvoice.SelectedItem.SubItems(20)
            A7 = lstInvoice.SelectedItem.SubItems(21)
            A8 = lstInvoice.SelectedItem.SubItems(22)
            A9 = lstInvoice.SelectedItem.SubItems(23)
            A10 = lstInvoice.SelectedItem.SubItems(24)
            A11 = lstInvoice.SelectedItem.SubItems(25)
            A12 = lstInvoice.SelectedItem.SubItems(26)
            A13 = lstInvoice.SelectedItem.SubItems(27)
            A14 = lstInvoice.SelectedItem.SubItems(28)
            A15 = lstInvoice.SelectedItem.SubItems(29)
            A16 = lstInvoice.SelectedItem.SubItems(30)
            A17 = lstInvoice.SelectedItem.SubItems(31)
            A18 = lstInvoice.SelectedItem.SubItems(32)
            A19 = lstInvoice.SelectedItem.SubItems(33)
            A20 = lstInvoice.SelectedItem.SubItems(34)
            A21 = lstInvoice.SelectedItem.SubItems(35)
            A22 = lstInvoice.SelectedItem.SubItems(36)
            A23 = lstInvoice.SelectedItem.SubItems(37)
            A24 = lstInvoice.SelectedItem.SubItems(38)
            A25 = lstInvoice.SelectedItem.SubItems(39)
            A26 = lstInvoice.SelectedItem.SubItems(40)
            A27 = lstInvoice.SelectedItem.SubItems(41)
            A28 = lstInvoice.SelectedItem.SubItems(42)
            A29 = lstInvoice.SelectedItem.SubItems(43)
            A30 = lstInvoice.SelectedItem.SubItems(101)
            txtPOthDesc = lstInvoice.SelectedItem.SubItems(44)
            txtLOthDesc = lstInvoice.SelectedItem.SubItems(45)
            txtIOthDesc = lstInvoice.SelectedItem.SubItems(46)
            
            txtMedType1 = lstInvoice.SelectedItem.SubItems(47)
            txtMedType2 = lstInvoice.SelectedItem.SubItems(48)
            txtMedType3 = lstInvoice.SelectedItem.SubItems(49)
            txtMedType4 = lstInvoice.SelectedItem.SubItems(50)
            txtMedType5 = lstInvoice.SelectedItem.SubItems(51)
            txtMedDay1 = lstInvoice.SelectedItem.SubItems(52)
            txtMedDay2 = lstInvoice.SelectedItem.SubItems(53)
            txtMedDay3 = lstInvoice.SelectedItem.SubItems(54)
            txtMedDay4 = lstInvoice.SelectedItem.SubItems(55)
            txtMedDay5 = lstInvoice.SelectedItem.SubItems(56)
            txtMedQty1 = lstInvoice.SelectedItem.SubItems(57)
            txtMedQty2 = lstInvoice.SelectedItem.SubItems(58)
            txtMedQty3 = lstInvoice.SelectedItem.SubItems(59)
            txtMedQty4 = lstInvoice.SelectedItem.SubItems(60)
            txtMedQty5 = lstInvoice.SelectedItem.SubItems(61)
            txtDiagnosis1 = lstInvoice.SelectedItem.SubItems(62)
            txtDiagnosis2 = lstInvoice.SelectedItem.SubItems(63)
            txtDiagnosis3 = lstInvoice.SelectedItem.SubItems(64)
            txtDiagnosis4 = lstInvoice.SelectedItem.SubItems(65)
            TxtPayor = lstInvoice.SelectedItem.SubItems(66)
            SuppLmtType = lstInvoice.SelectedItem.SubItems(82)
            B1 = lstInvoice.SelectedItem.SubItems(83)
            B2 = lstInvoice.SelectedItem.SubItems(84)
            B3 = lstInvoice.SelectedItem.SubItems(85)
            B4 = lstInvoice.SelectedItem.SubItems(86)
            B5 = lstInvoice.SelectedItem.SubItems(87)
            B6 = lstInvoice.SelectedItem.SubItems(88)
            B7 = lstInvoice.SelectedItem.SubItems(89)
            B8 = lstInvoice.SelectedItem.SubItems(90)
            Z1 = lstInvoice.SelectedItem.SubItems(83)
            Z2 = lstInvoice.SelectedItem.SubItems(84)
            Z3 = lstInvoice.SelectedItem.SubItems(85)
            Z4 = lstInvoice.SelectedItem.SubItems(86)
            Z5 = lstInvoice.SelectedItem.SubItems(87)
            Z6 = lstInvoice.SelectedItem.SubItems(88)
            Z7 = lstInvoice.SelectedItem.SubItems(89)
            Z8 = lstInvoice.SelectedItem.SubItems(90)
            C1 = lstInvoice.SelectedItem.SubItems(91)
            C2 = lstInvoice.SelectedItem.SubItems(92)
            TxtDiagDesc5.Visible = True
            TxtDiagnosis5 = lstInvoice.SelectedItem.SubItems(93)
            ReasonOtherIND = lstInvoice.SelectedItem.SubItems(94)
            If ReasonOtherIND = "Y" Then
                TxtDiagDesc5.Visible = False
                TxtDiagnosisOth.Enabled = True
            Else
                TxtDiagnosisOth = ""
            End If
            TxtDiagnosisOth = lstInvoice.SelectedItem.SubItems(95)
            TxtDiagDesc1 = lstInvoice.SelectedItem.SubItems(96)
            TxtDiagDesc2 = lstInvoice.SelectedItem.SubItems(97)
            TxtDiagDesc3 = lstInvoice.SelectedItem.SubItems(98)
            TxtDiagDesc4 = lstInvoice.SelectedItem.SubItems(99)
            TxtDiagDesc5 = lstInvoice.SelectedItem.SubItems(100)
            CmdAddRec.Enabled = False
        End If
'    End If
End Sub


Private Sub TxtClinicCode_LostFocus()
Dim ClnRec As New ADODB.Recordset
Dim CLNSql As String
    TxtClinicCode = ChkStr(TxtClinicCode)
    If Trim(TxtClinicCode) <> "" Then
             CboClinic.Clear
        CLNSql = "SELECT CLNCode, CLNName FROM CLN WHERE (CLNPanel = 'C' OR CLNPanel = 'Y')" & _
                " AND (CLNCode ='" & Trim(TxtClinicCode) & "')"

        ClnRec.Open CLNSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If ClnRec.RecordCount <= 0 Then
            MsgBox "No such Clinic Code ! ", vbOKOnly + vbCritical, DialogBoxTitle
            ClnRec.Close
            Set ClnRec = Nothing
            TxtClinicCode.SetFocus
        Else
            CboClinic.AddItem ClnRec!CLNCode & " - " & ClnRec!CLNname
            CboClinic.ListIndex = 0
            ClnRec.Close
            Set ClnRec = Nothing
        End If
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

Private Sub txtMBMNo_GotFocus()
    cmdGo.Default = True
End Sub

Private Sub txtMBMNo_LostFocus()
    cmdGo.Default = False
End Sub

Private Sub EntriesEnabled(EStatus As Boolean)
    FrameInv.Enabled = EStatus
    FrameDet.Enabled = EStatus
    FrameDiag.Enabled = EStatus
    FrameMed.Enabled = EStatus
'    FrameInvoice.Enabled = EStatus
End Sub


Private Sub txtMC_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtMedDay1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtMedDay2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub txtMedDay3_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub txtMedDay4_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub txtMedDay5_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub txtMedqty1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtMedqty2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub txtMedqty3_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub txtMedqty4_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub txtMedqty5_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
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
    ElseIf txtVisitDate = "__/__/____" Or txtVisitDate = "" Then
        MsgBox "Please Enter Visit Date! ", vbCritical
        txtVisitDate.SetFocus
    ElseIf Trim(CboClinic) = "" Then
        MsgBox "Please keyin the Clinic code ! ", vbCritical
        TxtClinicCode.SetFocus
    ElseIf Trim(txtInvNo) = "" Then
        MsgBox "Please Enter Invoice No ! ", vbCritical
        txtInvNo.SetFocus
    ElseIf txtInvDate = "__/__/____" Or txtInvDate = "" Then
        MsgBox "Please Enter Invoice Date! ", vbCritical
        txtInvDate.SetFocus
    ElseIf Trim(TxtTotInvoiceAmt) = "" Then
        MsgBox "Please Keyin Invoice Amount ! ", vbCritical
        TxtTotInvoiceAmt.SetFocus
    ElseIf IntInvAmt = 0 Then
        MsgBox "Total Invoice Amount cannot be Zero ! ", vbCritical
        B1.SetFocus
    ElseIf Trim(ReasonOtherIND) = "Y" And Trim(TxtDiagnosisOth) = "" Then
        MsgBox "Please Keyin Others Diagnosis ! ", vbCritical
        TxtDiagnosisOth.Enabled = True
        TxtDiagnosisOth.SetFocus
    Else
    
        Erase ArrChk
        ReDim ArrChk(0)
        QQ = 0
        CISCasSql = "SELECT * FROM View_INV_CLM where MBMNumber = '" & Trim(MembershipNo) & _
              "' AND COVERID ='" & CovPersonID & "' AND BillDate = '" & Format(txtVisitDate, "mm/dd/yyyy") & "'"
        If Trim(txtBillNo) <> "" Then
            CISCasSql = CISCasSql & " AND BillNo = '" & Trim(txtBillNo) & "'"
        End If
        CISCasRec.Open CISCasSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
        If CISCasRec.RecordCount Then
            Do Until CISCasRec.EOF
                QQ = QQ + 1
                ReDim Preserve ArrChk(QQ)
                ArrChk(QQ) = Trim(CISCasRec!MBMNumber & "~" & CISCasRec!CoverID & "#" & CISCasRec!DateVisit & "*" & CISCasRec!BillNo)
                CISCasRec.MoveNext
            Loop
        End If
        CISCasRec.Close
        Set CISCasRec = Nothing
        For BB = 1 To lstInvoice.ListItems.count
            ReDim Preserve ArrChk(QQ + BB)
            ArrChk(QQ + BB) = Trim(lstInvoice.ListItems(BB).SubItems(2) & "~" & lstInvoice.ListItems(BB).SubItems(3) & "#" & lstInvoice.ListItems(BB).SubItems(9) & "*" & lstInvoice.ListItems(BB).SubItems(7))
        Next BB
        ChkDupKey = MembershipNo & "~" & CovPersonID & "#" & txtVisitDate & "*" & txtBillNo
        AddNewInv = True
        For XX = 0 To UBound(ArrChk)
            If EditFlag = False Then
            If Trim(ChkDupKey) = ArrChk(XX) Then
                AddNewInv = False
                If Trim(txtBillNo) = "" Then
                    DupInvoice = MsgBox("Invoice has been entered !, Press 'Yes' to Proceed, 'No' to Cancel", vbYesNo + vbExclamation)
                    If DupInvoice = vbYes Then
                        If EditFlag = True Then
                            Call RemoveRecFrListV
                        End If
                        Call AssignData2ListView
                        Load frmWorksheet
                        Unload frmWorksheet
                        EditFlag = False
                        Call ClearEntries
                        Call ClearMBMInfo
                        FrameMBMCLN.Enabled = True
                        txtMBMNo.SetFocus
                    ElseIf DupInvoice = vbNo Then
                        Call EntriesEnabled(False)
                    End If
                Else
                    MsgBox "Record Duplicated, Same Invoice has been entered ", vbCritical, "Error"
                End If
            End If
            End If
        Next XX
        If AddNewInv = True Then
            If EditFlag = True Then
                Call RemoveRecFrListV
            End If
            Call AssignData2ListView
            Load frmWorksheet
            Unload frmWorksheet
            EditFlag = False
            Call ClearEntries
            Call ClearMBMInfo
            FrameMBMCLN.Enabled = True
            txtMBMNo.SetFocus
        End If
        CmdVerify.Visible = True
        CmdVerify.Enabled = True
        CmdEdit.Visible = True
        CmdEdit.Enabled = True
        CmdSave.Visible = False
        CmdSave.Enabled = False
        CmdMBMInfo.Enabled = False
        CmdAddRec.Enabled = True
        CmdUpdate.Visible = False
        CmdUpdate.Enabled = False
        FrameMBMCLN.Enabled = True
    End If
End Sub

Private Function getTotal() As Double
Dim ctl  As Control
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
End Function

Private Sub AssignData2ListView()
Dim strPayee As String

    CmdSave.Enabled = True
    Set lstRec = lstInvoice.ListItems.Add(, , "")
    'frmInvoice
    If strPayTo = "C" Then
'        strPayee = Trim(Left(CboClinic, IIf(InStr(1, CboClinic, "-") - 1 < 0, Len(CboClinic), InStr(1, CboClinic, "-") - 1)))
        strPayee = TxtClinicCode
    Else
        strPayee = "MEMBER"
    End If
       
    Call getTotal
    lstRec.SubItems(2) = MembershipNo
    lstRec.SubItems(3) = CovPersonID
    lstRec.SubItems(4) = txtPatientName
    lstRec.SubItems(6) = IIf(Len(Trim(txtInvNo)), Trim(txtInvNo), "")
    lstRec.SubItems(7) = IIf(Len(Trim(txtBillNo)), txtBillNo, "")
    lstRec.SubItems(8) = IIf(IsDate(txtInvDate), txtInvDate, "__/__/____")
    lstRec.SubItems(9) = IIf(IsDate(txtVisitDate), txtVisitDate, "__/__/____")
    lstRec.SubItems(10) = IIf(isTime(txtVisitTime), txtVisitTime, "__:__")
    lstRec.SubItems(11) = IIf(IsNumeric(txtMC), txtMC, 0)
    lstRec.SubItems(12) = Trim(Left(CboClinic, IIf(InStr(1, CboClinic, "-") - 1 < 0, Len(CboClinic), InStr(1, CboClinic, "-") - 1)))
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
    lstRec.SubItems(43) = IIf(IsNumeric(A29), A29, 0)
    lstRec.SubItems(44) = IIf(Len(Trim(txtPOthDesc)), Trim(txtPOthDesc), "")
    lstRec.SubItems(45) = IIf(Len(Trim(txtLOthDesc)), Trim(txtLOthDesc), "")
    lstRec.SubItems(46) = IIf(Len(Trim(txtIOthDesc)), Trim(txtIOthDesc), "")
    lstRec.SubItems(47) = IIf(Len(Trim(txtMedType1)), Trim(txtMedType1), "")
    lstRec.SubItems(48) = IIf(Len(Trim(txtMedType2)), Trim(txtMedType2), "")
    lstRec.SubItems(49) = IIf(Len(Trim(txtMedType3)), Trim(txtMedType3), "")
    lstRec.SubItems(50) = IIf(Len(Trim(txtMedType4)), Trim(txtMedType4), "")
    lstRec.SubItems(51) = IIf(Len(Trim(txtMedType5)), Trim(txtMedType5), "")
    lstRec.SubItems(52) = IIf(Len(Trim(txtMedDay1)), Trim(txtMedDay1), "")
    lstRec.SubItems(53) = IIf(Len(Trim(txtMedDay2)), Trim(txtMedDay2), "")
    lstRec.SubItems(54) = IIf(Len(Trim(txtMedDay3)), Trim(txtMedDay3), "")
    lstRec.SubItems(55) = IIf(Len(Trim(txtMedDay4)), Trim(txtMedDay4), "")
    lstRec.SubItems(56) = IIf(Len(Trim(txtMedDay5)), Trim(txtMedDay5), "")
    lstRec.SubItems(57) = IIf(Len(Trim(txtMedQty1)), Trim(txtMedQty1), "")
    lstRec.SubItems(58) = IIf(Len(Trim(txtMedQty2)), Trim(txtMedQty2), "")
    lstRec.SubItems(59) = IIf(Len(Trim(txtMedQty3)), Trim(txtMedQty3), "")
    lstRec.SubItems(60) = IIf(Len(Trim(txtMedQty4)), Trim(txtMedQty4), "")
    lstRec.SubItems(61) = IIf(Len(Trim(txtMedQty5)), Trim(txtMedQty5), "")
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
    lstRec.SubItems(72) = TxtClnContract
    lstRec.SubItems(73) = strVisitIND
    lstRec.SubItems(74) = strAnnualInd
    lstRec.SubItems(75) = TxtCoPay
    lstRec.SubItems(76) = TxtPolNo
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
    lstRec.SubItems(96) = IIf(Len(Trim(TxtDiagDesc1)), Trim(TxtDiagDesc1), "")
    lstRec.SubItems(97) = IIf(Len(Trim(TxtDiagDesc2)), Trim(TxtDiagDesc2), "")
    lstRec.SubItems(98) = IIf(Len(Trim(TxtDiagDesc3)), Trim(TxtDiagDesc3), "")
    lstRec.SubItems(99) = IIf(Len(Trim(TxtDiagDesc4)), Trim(TxtDiagDesc4), "")
    lstRec.SubItems(100) = IIf(Len(Trim(TxtDiagDesc5)), Trim(TxtDiagDesc5), "")
    lstRec.SubItems(101) = IIf(IsNumeric(A30), A30, 0)
    Call ClearEntries
    Call ClearMBMInfo
    txtMBMNo = ""
    Call EntriesEnabled(False)
    'txtMBMNo.SetFocus
End Sub

Private Sub txtMedType1_LostFocus()
    txtMedType1 = ChkStr(txtMedType1)
End Sub

Private Sub txtMedType2_LostFocus()
    txtMedType2 = ChkStr(txtMedType2)
End Sub

Private Sub txtMedType3_LostFocus()
    txtMedType3 = ChkStr(txtMedType3)
End Sub

Private Sub txtMedType4_LostFocus()
    txtMedType4 = ChkStr(txtMedType4)
End Sub

Private Sub txtMedType5_LostFocus()
    txtMedType5 = ChkStr(txtMedType5)
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
        End If
    End If
End Sub

Private Sub txtInvNo_LostFocus()
    txtInvNo = ChkStr(txtInvNo)
End Sub

Private Sub TxtTotInvoiceAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtVisitDate_LostFocus()
    If txtVisitDate <> "__/__/____" Then
        If IsDate(txtVisitDate) = False Then
            MsgBox "Invalid date format!", vbCritical
            txtVisitDate.SetFocus
        ElseIf CDate(txtVisitDate) > CDate(Date) Then
            MsgBox "Visit Date cannot greater than system date ! ", vbCritical
            txtVisitDate.SetFocus
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

Private Sub ErrRpt()
        
    Set lstErrRec = lstErr.ListItems.Add(, , lstInvoice.ListItems(CC).SubItems(2))
        lstErrRec.SubItems(1) = lstInvoice.ListItems(CC).SubItems(3)
        lstErrRec.SubItems(2) = lstInvoice.ListItems(CC).SubItems(4)
        lstErrRec.SubItems(3) = lstInvoice.ListItems(CC).SubItems(76)
        lstErrRec.SubItems(4) = lstInvoice.ListItems(CC).SubItems(77)
        lstErrRec.SubItems(5) = lstInvoice.ListItems(CC).SubItems(78)
        lstErrRec.SubItems(6) = lstInvoice.ListItems(CC).SubItems(79)
        lstErrRec.SubItems(7) = lstInvoice.ListItems(CC).SubItems(70)
        lstErrRec.SubItems(8) = lstInvoice.ListItems(CC).SubItems(71)
End Sub
Private Sub SaveRecord()
On Err GoTo ErrorSave
Dim startTrans As Boolean
    
    MediConnCISWS.beginTrans
        startTrans = True
        Call SaveCase
        Call SaveInvoice
        Call saveActual
        Call saveActualOthers
        Call SaveINVClaim
        Call SaveINVCLNApp '426-454
        Call SaveINVCLNContract '211-239
        Call SaveINVActual '168-196
        Call SaveINVCLNExcess '383-411
        Call SaveINVMBMApp '254-282
        Call SaveINVMBMExcess '340 - 368
        Call SaveINVMBMLmt '297-325
        Call UpdateMBMBalance

    MediConnCISWS.CommitTrans
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
        MediConnCISWS.RollbackTrans
    End If
    MsgBox Err.Description
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

Private Function GenerateOPSeq(OPVar)
Dim OPRec As New ADODB.Recordset
Dim strsql As String
Dim SeqNo As String
    
    strsql = "Select OPCode,OPBatchSeq from OPBatch Where OPCode = '" & OPVar & "'"
    OPRec.Open strsql, MediConnCISDB, adOpenStatic, adLockOptimistic
    
    If OPRec.RecordCount <= 0 Then
        With OPRec
            .AddNew
            !OPCode = OPVar
            !OPBAtchSeq = "00001"
            .Update
            SeqNo = "00001"
        End With
    Else
        OPRec!OPBAtchSeq = Format(OPRec!OPBAtchSeq + 1, "00000")
        SeqNo = OPRec!OPBAtchSeq
        OPRec.Update
    End If

    GenerateOPSeq = OPVar & SeqNo
    OPRec.Close
    Set OPRec = Nothing
End Function

Private Sub SaveCase()
Dim CISCAS As New ADODB.Recordset
Dim CISSql As String
Dim tmpDiag1, tmpDiag2, tmpDiag3, tmpDiag4, tmpDiag5 As String

    CISSql = "SELECT * FROM CISCAS "
    CISCAS.Open CISSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    CISCAS.AddNew
    NewCaseNo = GenerateCISCaseNo(lstInvoice.ListItems(AA).SubItems(66))

    With CISCAS
        !CASNumber = NewCaseNo
        lstInvoice.ListItems(AA).Text = OPBatchNo
        lstInvoice.ListItems(AA).SubItems(1) = NewCaseNo
        !MBMNumber = lstInvoice.ListItems(AA).SubItems(2)
        !CoverID = lstInvoice.ListItems(AA).SubItems(3)
        !CASStatus = "O"
        !Claimability = "A"
        !Pending = "N"
        !REgisterDate = Date
        !REgisterBy = Logon
        !LastUpdateDate = Date
        !LastUpdateBy = Logon
        !ClinicCode = lstInvoice.ListItems(AA).SubItems(12)
        !InvoiceNo = lstInvoice.ListItems(AA).SubItems(6)
        !InvoiceDate = lstInvoice.ListItems(AA).SubItems(8)
        !InvoiceStmtAmt = IIf(IsNumeric(TxtTotInvoiceAmt), TxtTotInvoiceAmt, 0)
        !DateVisit = lstInvoice.ListItems(AA).SubItems(9)
        !MCDays = lstInvoice.ListItems(AA).SubItems(11)
        !OPBatchNo = OPBatchNo
        If isTime(lstInvoice.ListItems(AA).SubItems(10)) Then
            !TimeVisit = lstInvoice.ListItems(AA).SubItems(10)
        End If
        
        !MedInjType1 = lstInvoice.ListItems(AA).SubItems(47)
        !MedInjType2 = lstInvoice.ListItems(AA).SubItems(48)
        !MedInjType3 = lstInvoice.ListItems(AA).SubItems(49)
        !MedInjType4 = lstInvoice.ListItems(AA).SubItems(50)
        !MedInjType5 = lstInvoice.ListItems(AA).SubItems(51)
        !MedInjDays1 = lstInvoice.ListItems(AA).SubItems(52)
        !MedInjDays2 = lstInvoice.ListItems(AA).SubItems(53)
        !MedInjDays3 = lstInvoice.ListItems(AA).SubItems(54)
        !MedInjDays4 = lstInvoice.ListItems(AA).SubItems(55)
        !MedInjDays5 = lstInvoice.ListItems(AA).SubItems(56)
        !MedInjQty1 = lstInvoice.ListItems(AA).SubItems(57)
        !MedInjQty2 = lstInvoice.ListItems(AA).SubItems(58)
        !MedInjQty3 = lstInvoice.ListItems(AA).SubItems(59)
        !MedInjQty4 = lstInvoice.ListItems(AA).SubItems(60)
        !MedInjQty5 = lstInvoice.ListItems(AA).SubItems(61)
        
        !CASTreatReason1 = lstInvoice.ListItems(AA).SubItems(62)
        !CASTreatReason2 = lstInvoice.ListItems(AA).SubItems(63)
        !CASTreatReason3 = lstInvoice.ListItems(AA).SubItems(64)
        !CASTreatReason4 = lstInvoice.ListItems(AA).SubItems(65)
        !CASTreatReason5 = lstInvoice.ListItems(AA).SubItems(93)
        !CASDiagnosis5 = lstInvoice.ListItems(AA).SubItems(95)
    End With
    CISCAS.Update
    
    CISCAS.Close
    Set CISCAS = Nothing
End Sub

Private Sub SaveInvoice()
Dim ClaimInvoice  As New ADODB.Recordset
    
    ClaimInvoice.Open "Select INVINdex , * from ClaimInvoice", MediConnCISWS, adOpenKeyset, adLockOptimistic
    ClaimInvoice.AddNew
    With ClaimInvoice
        !CASNumber = NewCaseNo
        !INVEnteredDate = Date
        !INVEnteredBy = Logon
        !INVLastUpdateBy = Logon
        !INVLastUpdateDate = Date
        !BillNo = lstInvoice.ListItems(AA).SubItems(7)
        !BillDate = lstInvoice.ListItems(AA).SubItems(9)
        !INVAmount = lstInvoice.ListItems(AA).SubItems(13)
        !INVPayTo = lstInvoice.ListItems(AA).SubItems(80)
        !INVPayee = lstInvoice.ListItems(AA).SubItems(81)

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
 
    ClaimActual.Open "Select INVINdex , * from ClaimInvoiceActualNew ", MediConnCISWS, adOpenKeyset, adLockOptimistic
    ClaimActual.AddNew
    ClaimActual!InvINdex = intIndex
    
    With ClaimActual
        !CASNumber = NewCaseNo
        !CLMGPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(15)
        !CLMGPFeeLTMed = lstInvoice.ListItems(AA).SubItems(16)
        !CLMGPFeeMedInj = lstInvoice.ListItems(AA).SubItems(17)
        !CLMSPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(18)
        !CLMSPFeeLTMed = lstInvoice.ListItems(AA).SubItems(19)
        !CLMSPFeeMedInj = lstInvoice.ListItems(AA).SubItems(20)
        !CLMDental = lstInvoice.ListItems(AA).SubItems(21)
        !CLMMaternity = lstInvoice.ListItems(AA).SubItems(22)
        !CLMPRVCheckup = lstInvoice.ListItems(AA).SubItems(23)
        !CLMXRay = lstInvoice.ListItems(AA).SubItems(24)
        !CLMUltraSound = lstInvoice.ListItems(AA).SubItems(25)
        !CLMECG = lstInvoice.ListItems(AA).SubItems(26)
        !CLMDressing = lstInvoice.ListItems(AA).SubItems(27)
        !CLMNebulizer = lstInvoice.ListItems(AA).SubItems(28)
        !CLMID = lstInvoice.ListItems(AA).SubItems(29)
        !CLMTSSTO = lstInvoice.ListItems(AA).SubItems(30)
        !CLMMinorSurgery = lstInvoice.ListItems(AA).SubItems(31)
        !CLMPOthers = lstInvoice.ListItems(AA).SubItems(32)
        !CLMBloodDrawing = lstInvoice.ListItems(AA).SubItems(33)
        !CLMBloodGlucose = lstInvoice.ListItems(AA).SubItems(34)
        !CLMUrine = lstInvoice.ListItems(AA).SubItems(35)
        !CLMPapSmear = lstInvoice.ListItems(AA).SubItems(36)
        !CLMLOthers = lstInvoice.ListItems(AA).SubItems(37)
        !CLMDTPPolio = lstInvoice.ListItems(AA).SubItems(38)
        !CLMDTPPolioHIB = lstInvoice.ListItems(AA).SubItems(39)
        !CLMHepatitisB = lstInvoice.ListItems(AA).SubItems(40)
        !CLMMMR = lstInvoice.ListItems(AA).SubItems(41)
        !CLMIOthers = lstInvoice.ListItems(AA).SubItems(42)
        !CLMCoPayment = lstInvoice.ListItems(AA).SubItems(43)
        !CLMPOthersDesc = lstInvoice.ListItems(AA).SubItems(44)
        !CLMLOthersDesc = lstInvoice.ListItems(AA).SubItems(45)
        !CLMIOthersDesc = lstInvoice.ListItems(AA).SubItems(46)
        !CLMVision = lstInvoice.ListItems(AA).SubItems(101)
    End With
    ClaimActual.Update
    ClaimActual.Close
    Set ClaimActual = Nothing
End Sub
Private Sub saveActualOthers()
    Dim ClaimActOTH As New ADODB.Recordset
    Dim strsql As String
 
    ClaimActOTH.Open "Select INVINdex , * from ClaimInvoiceActualOthers ", MediConnCISWS, adOpenKeyset, adLockOptimistic
    ClaimActOTH.AddNew
    ClaimActOTH!InvINdex = intIndex
    
    With ClaimActOTH
        !CASNumber = NewCaseNo
        !GPFee = lstInvoice.ListItems(AA).SubItems(83)
        !GPLTMed = lstInvoice.ListItems(AA).SubItems(84)
        !GPInj = lstInvoice.ListItems(AA).SubItems(85)
        !GPMed = lstInvoice.ListItems(AA).SubItems(86)
        !SPFee = lstInvoice.ListItems(AA).SubItems(87)
        !SPLTMed = lstInvoice.ListItems(AA).SubItems(88)
        !SPInj = lstInvoice.ListItems(AA).SubItems(89)
        !SPMed = lstInvoice.ListItems(AA).SubItems(90)
    End With
    ClaimActOTH.Update
    ClaimActOTH.Close
    Set ClaimActOTH = Nothing
End Sub

Private Sub SaveINVActual()
Dim CISAct As New ADODB.Recordset
Dim CISSql As String
    
    CISSql = "SELECT * FROM CLMActualNew "
    CISAct.Open CISSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    CISAct.AddNew
    'NewCaseNo = GenerateCISCaseNo(lstInvoice.ListItems(AA).SubItems(65))

    With CISAct
        !CLMNo = NewCaseNo
        !CLMGPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(168)
        !CLMGPFeeMedInj = lstInvoice.ListItems(AA).SubItems(169)
        !CLMGPFeeLTMed = lstInvoice.ListItems(AA).SubItems(170)
        !CLMSPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(171)
        !CLMSPFeeLTMed = lstInvoice.ListItems(AA).SubItems(172)
        !CLMSPFeeMedInj = lstInvoice.ListItems(AA).SubItems(173)
        !CLMXRay = lstInvoice.ListItems(AA).SubItems(174)
        !CLMUltraSound = lstInvoice.ListItems(AA).SubItems(175)
        !CLMECG = lstInvoice.ListItems(AA).SubItems(176)
        !CLMDressing = lstInvoice.ListItems(AA).SubItems(177)
        !CLMNebulizer = lstInvoice.ListItems(AA).SubItems(178)
        !CLMID = lstInvoice.ListItems(AA).SubItems(179)
        !CLMTSSTO = lstInvoice.ListItems(AA).SubItems(180)
        !CLMMinorSurgery = lstInvoice.ListItems(AA).SubItems(181)
        !CLMPOthers = lstInvoice.ListItems(AA).SubItems(182)
        !CLMBloodDrawing = lstInvoice.ListItems(AA).SubItems(183)
        !CLMBloodGlucose = lstInvoice.ListItems(AA).SubItems(184)
        !CLMUrine = lstInvoice.ListItems(AA).SubItems(185)
        !CLMPapSmear = lstInvoice.ListItems(AA).SubItems(186)
        !CLMLOthers = lstInvoice.ListItems(AA).SubItems(187)
        !CLMDTPPolio = lstInvoice.ListItems(AA).SubItems(188)
        !CLMDTPPolioHIB = lstInvoice.ListItems(AA).SubItems(189)
        !CLMHepatitisB = lstInvoice.ListItems(AA).SubItems(190)
        !CLMMMR = lstInvoice.ListItems(AA).SubItems(191)
        !CLMIOthers = lstInvoice.ListItems(AA).SubItems(192)
        !CLMDental = lstInvoice.ListItems(AA).SubItems(193)
        !CLMMaternity = lstInvoice.ListItems(AA).SubItems(194)
        !CLMPRVCheckup = lstInvoice.ListItems(AA).SubItems(195)
        !CLMVision = lstInvoice.ListItems(AA).SubItems(196)
    End With
    CISAct.Update
    
    CISAct.Close
    Set CISAct = Nothing
End Sub

Private Sub SaveINVCLNContract()
Dim CISCLNLmt As New ADODB.Recordset
Dim CISSql As String
    
    CISSql = "SELECT * FROM CLMClinicLmtNew "
    CISCLNLmt.Open CISSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    CISCLNLmt.AddNew
    'NewCaseNo = GenerateCISCaseNo(lstInvoice.ListItems(AA).SubItems(65))

    With CISCLNLmt
        !CASNumber = NewCaseNo
        !CLMGPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(211)
        !CLMGPFeeMedInj = lstInvoice.ListItems(AA).SubItems(212)
        !CLMGPFeeLTMed = lstInvoice.ListItems(AA).SubItems(213)
        !CLMSPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(214)
        !CLMSPFeeLTMed = lstInvoice.ListItems(AA).SubItems(215)
        !CLMSPFeeMedInj = lstInvoice.ListItems(AA).SubItems(216)
        !CLMXRay = lstInvoice.ListItems(AA).SubItems(217)
        !CLMUltraSound = lstInvoice.ListItems(AA).SubItems(218)
        !CLMECG = lstInvoice.ListItems(AA).SubItems(219)
        !CLMDressing = lstInvoice.ListItems(AA).SubItems(220)
        !CLMNebulizer = lstInvoice.ListItems(AA).SubItems(221)
        !CLMID = lstInvoice.ListItems(AA).SubItems(222)
        !CLMTSSTO = lstInvoice.ListItems(AA).SubItems(223)
        !CLMMinorSurgery = lstInvoice.ListItems(AA).SubItems(224)
        !CLMPOthers = lstInvoice.ListItems(AA).SubItems(225)
        !CLMBloodDrawing = lstInvoice.ListItems(AA).SubItems(226)
        !CLMBloodGlucose = lstInvoice.ListItems(AA).SubItems(227)
        !CLMUrine = lstInvoice.ListItems(AA).SubItems(228)
        !CLMPapSmear = lstInvoice.ListItems(AA).SubItems(229)
        !CLMLOthers = lstInvoice.ListItems(AA).SubItems(230)
        !CLMDTPPolio = lstInvoice.ListItems(AA).SubItems(231)
        !CLMDTPPolioHIB = lstInvoice.ListItems(AA).SubItems(232)
        !CLMHepatitisB = lstInvoice.ListItems(AA).SubItems(233)
        !CLMMMR = lstInvoice.ListItems(AA).SubItems(234)
        !CLMIOthers = lstInvoice.ListItems(AA).SubItems(235)
        !CLMDental = lstInvoice.ListItems(AA).SubItems(236)
        !CLMMaternity = lstInvoice.ListItems(AA).SubItems(237)
        !CLMPRVCheckup = lstInvoice.ListItems(AA).SubItems(238)
        !CLMVision = lstInvoice.ListItems(AA).SubItems(239)
    End With
    CISCLNLmt.Update
    
    CISCLNLmt.Close
    Set CISCLNLmt = Nothing
End Sub

Private Sub SaveINVMBMApp()
Dim CISAPP As New ADODB.Recordset
Dim CISSql As String
    
    CISSql = "SELECT * FROM CLMAppNew "
    CISAPP.Open CISSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    CISAPP.AddNew
    'NewCaseNo = GenerateCISCaseNo(lstInvoice.ListItems(AA).SubItems(65))

    With CISAPP
        !CASNumber = NewCaseNo
        !CLMGPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(254)
        !CLMGPFeeMedInj = lstInvoice.ListItems(AA).SubItems(255)
        !CLMGPFeeLTMed = lstInvoice.ListItems(AA).SubItems(256)
        !CLMSPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(257)
        !CLMSPFeeLTMed = lstInvoice.ListItems(AA).SubItems(258)
        !CLMSPFeeMedInj = lstInvoice.ListItems(AA).SubItems(259)
        !CLMXRay = lstInvoice.ListItems(AA).SubItems(260)
        !CLMUltraSound = lstInvoice.ListItems(AA).SubItems(261)
        !CLMECG = lstInvoice.ListItems(AA).SubItems(262)
        !CLMDressing = lstInvoice.ListItems(AA).SubItems(263)
        !CLMNebulizer = lstInvoice.ListItems(AA).SubItems(264)
        !CLMID = lstInvoice.ListItems(AA).SubItems(265)
        !CLMTSSTO = lstInvoice.ListItems(AA).SubItems(266)
        !CLMMinorSurgery = lstInvoice.ListItems(AA).SubItems(267)
        !CLMPOthers = lstInvoice.ListItems(AA).SubItems(268)
        !CLMBloodDrawing = lstInvoice.ListItems(AA).SubItems(269)
        !CLMBloodGlucose = lstInvoice.ListItems(AA).SubItems(270)
        !CLMUrine = lstInvoice.ListItems(AA).SubItems(271)
        !CLMPapSmear = lstInvoice.ListItems(AA).SubItems(272)
        !CLMLOthers = lstInvoice.ListItems(AA).SubItems(273)
        !CLMDTPPolio = lstInvoice.ListItems(AA).SubItems(274)
        !CLMDTPPolioHIB = lstInvoice.ListItems(AA).SubItems(275)
        !CLMHepatitisB = lstInvoice.ListItems(AA).SubItems(276)
        !CLMMMR = lstInvoice.ListItems(AA).SubItems(277)
        !CLMIOthers = lstInvoice.ListItems(AA).SubItems(278)
        !CLMDental = lstInvoice.ListItems(AA).SubItems(279)
        !CLMMaternity = lstInvoice.ListItems(AA).SubItems(280)
        !CLMPRVCheckup = lstInvoice.ListItems(AA).SubItems(281)
        !CLMVision = lstInvoice.ListItems(AA).SubItems(282)
    End With
    CISAPP.Update
    
    CISAPP.Close
    Set CISAPP = Nothing
End Sub

Private Sub SaveINVMBMLmt()
Dim CISLmt As New ADODB.Recordset
Dim CISSql As String
    
    CISSql = "SELECT * FROM CLMLimitNew "
    CISLmt.Open CISSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    CISLmt.AddNew
    'NewCaseNo = GenerateCISCaseNo(lstInvoice.ListItems(AA).SubItems(65))

    With CISLmt
        !CASNumber = NewCaseNo
        !CLMGPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(297)
        !CLMGPFeeMedInj = lstInvoice.ListItems(AA).SubItems(298)
        !CLMGPFeeLTMed = lstInvoice.ListItems(AA).SubItems(299)
        !CLMSPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(300)
        !CLMSPFeeLTMed = lstInvoice.ListItems(AA).SubItems(301)
        !CLMSPFeeMedInj = lstInvoice.ListItems(AA).SubItems(302)
        !CLMXRay = lstInvoice.ListItems(AA).SubItems(303)
        !CLMUltraSound = lstInvoice.ListItems(AA).SubItems(304)
        !CLMECG = lstInvoice.ListItems(AA).SubItems(305)
        !CLMDressing = lstInvoice.ListItems(AA).SubItems(306)
        !CLMNebulizer = lstInvoice.ListItems(AA).SubItems(307)
        !CLMID = lstInvoice.ListItems(AA).SubItems(308)
        !CLMTSSTO = lstInvoice.ListItems(AA).SubItems(309)
        !CLMMinorSurgery = lstInvoice.ListItems(AA).SubItems(310)
        !CLMPOthers = lstInvoice.ListItems(AA).SubItems(311)
        !CLMBloodDrawing = lstInvoice.ListItems(AA).SubItems(312)
        !CLMBloodGlucose = lstInvoice.ListItems(AA).SubItems(313)
        !CLMUrine = lstInvoice.ListItems(AA).SubItems(314)
        !CLMPapSmear = lstInvoice.ListItems(AA).SubItems(315)
        !CLMLOthers = lstInvoice.ListItems(AA).SubItems(316)
        !CLMDTPPolio = lstInvoice.ListItems(AA).SubItems(317)
        !CLMDTPPolioHIB = lstInvoice.ListItems(AA).SubItems(318)
        !CLMHepatitisB = lstInvoice.ListItems(AA).SubItems(319)
        !CLMMMR = lstInvoice.ListItems(AA).SubItems(320)
        !CLMIOthers = lstInvoice.ListItems(AA).SubItems(321)
        !CLMDental = lstInvoice.ListItems(AA).SubItems(322)
        !CLMMaternity = lstInvoice.ListItems(AA).SubItems(323)
        !CLMPRVCheckup = lstInvoice.ListItems(AA).SubItems(324)
        !CLMVision = lstInvoice.ListItems(AA).SubItems(325)
    End With
    CISLmt.Update
    
    CISLmt.Close
    Set CISLmt = Nothing
End Sub

Private Sub SaveINVMBMExcess()
Dim CISExc As New ADODB.Recordset
Dim CISSql As String
    
    CISSql = "SELECT * FROM CLMMemberNew "
    CISExc.Open CISSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    CISExc.AddNew
    'NewCaseNo = GenerateCISCaseNo(lstInvoice.ListItems(AA).SubItems(65))

    With CISExc
        !CASNumber = NewCaseNo
        !CLMGPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(340)
        !CLMGPFeeMedInj = lstInvoice.ListItems(AA).SubItems(341)
        !CLMGPFeeLTMed = lstInvoice.ListItems(AA).SubItems(342)
        !CLMSPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(343)
        !CLMSPFeeLTMed = lstInvoice.ListItems(AA).SubItems(344)
        !CLMSPFeeMedInj = lstInvoice.ListItems(AA).SubItems(345)
        !CLMXRay = lstInvoice.ListItems(AA).SubItems(346)
        !CLMUltraSound = lstInvoice.ListItems(AA).SubItems(347)
        !CLMECG = lstInvoice.ListItems(AA).SubItems(348)
        !CLMDressing = lstInvoice.ListItems(AA).SubItems(349)
        !CLMNebulizer = lstInvoice.ListItems(AA).SubItems(350)
        !CLMID = lstInvoice.ListItems(AA).SubItems(351)
        !CLMTSSTO = lstInvoice.ListItems(AA).SubItems(352)
        !CLMMinorSurgery = lstInvoice.ListItems(AA).SubItems(353)
        !CLMPOthers = lstInvoice.ListItems(AA).SubItems(354)
        !CLMBloodDrawing = lstInvoice.ListItems(AA).SubItems(355)
        !CLMBloodGlucose = lstInvoice.ListItems(AA).SubItems(356)
        !CLMUrine = lstInvoice.ListItems(AA).SubItems(357)
        !CLMPapSmear = lstInvoice.ListItems(AA).SubItems(358)
        !CLMLOthers = lstInvoice.ListItems(AA).SubItems(359)
        !CLMDTPPolio = lstInvoice.ListItems(AA).SubItems(360)
        !CLMDTPPolioHIB = lstInvoice.ListItems(AA).SubItems(361)
        !CLMHepatitisB = lstInvoice.ListItems(AA).SubItems(362)
        !CLMMMR = lstInvoice.ListItems(AA).SubItems(363)
        !CLMIOthers = lstInvoice.ListItems(AA).SubItems(364)
        !CLMDental = lstInvoice.ListItems(AA).SubItems(365)
        !CLMMaternity = lstInvoice.ListItems(AA).SubItems(366)
        !CLMPRVCheckup = lstInvoice.ListItems(AA).SubItems(367)
        !CLMVision = lstInvoice.ListItems(AA).SubItems(368)
    End With
    CISExc.Update
    
    CISExc.Close
    Set CISExc = Nothing
End Sub

Private Sub SaveINVCLNExcess()
Dim CISExc As New ADODB.Recordset
Dim CISSql As String
    
    CISSql = "SELECT * FROM CLMClinicNew "
    CISExc.Open CISSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    CISExc.AddNew
    'NewCaseNo = GenerateCISCaseNo(lstInvoice.ListItems(AA).SubItems(65))

    With CISExc
        !CASNumber = NewCaseNo
        !CLMGPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(383)
        !CLMGPFeeMedInj = lstInvoice.ListItems(AA).SubItems(384)
        !CLMGPFeeLTMed = lstInvoice.ListItems(AA).SubItems(385)
        !CLMSPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(386)
        !CLMSPFeeLTMed = lstInvoice.ListItems(AA).SubItems(387)
        !CLMSPFeeMedInj = lstInvoice.ListItems(AA).SubItems(388)
        !CLMXRay = lstInvoice.ListItems(AA).SubItems(389)
        !CLMUltraSound = lstInvoice.ListItems(AA).SubItems(390)
        !CLMECG = lstInvoice.ListItems(AA).SubItems(391)
        !CLMDressing = lstInvoice.ListItems(AA).SubItems(392)
        !CLMNebulizer = lstInvoice.ListItems(AA).SubItems(393)
        !CLMID = lstInvoice.ListItems(AA).SubItems(394)
        !CLMTSSTO = lstInvoice.ListItems(AA).SubItems(395)
        !CLMMinorSurgery = lstInvoice.ListItems(AA).SubItems(396)
        !CLMPOthers = lstInvoice.ListItems(AA).SubItems(397)
        !CLMBloodDrawing = lstInvoice.ListItems(AA).SubItems(398)
        !CLMBloodGlucose = lstInvoice.ListItems(AA).SubItems(399)
        !CLMUrine = lstInvoice.ListItems(AA).SubItems(400)
        !CLMPapSmear = lstInvoice.ListItems(AA).SubItems(401)
        !CLMLOthers = lstInvoice.ListItems(AA).SubItems(402)
        !CLMDTPPolio = lstInvoice.ListItems(AA).SubItems(403)
        !CLMDTPPolioHIB = lstInvoice.ListItems(AA).SubItems(404)
        !CLMHepatitisB = lstInvoice.ListItems(AA).SubItems(405)
        !CLMMMR = lstInvoice.ListItems(AA).SubItems(406)
        !CLMIOthers = lstInvoice.ListItems(AA).SubItems(407)
        !CLMDental = lstInvoice.ListItems(AA).SubItems(408)
        !CLMMaternity = lstInvoice.ListItems(AA).SubItems(409)
        !CLMPRVCheckup = lstInvoice.ListItems(AA).SubItems(410)
        !CLMVision = lstInvoice.ListItems(AA).SubItems(411)
    End With
    CISExc.Update
    
    CISExc.Close
    Set CISExc = Nothing
End Sub


Private Sub SaveINVCLNApp()
Dim CISExc As New ADODB.Recordset
Dim CISSql As String
    
    CISSql = "SELECT * FROM CLMClinicAppNew "
    CISExc.Open CISSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    CISExc.AddNew
    'NewCaseNo = GenerateCISCaseNo(lstInvoice.ListItems(AA).SubItems(65))

    With CISExc
        !CASNumber = NewCaseNo
        !CLMGPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(426)
        !CLMGPFeeMedInj = lstInvoice.ListItems(AA).SubItems(427)
        !CLMGPFeeLTMed = lstInvoice.ListItems(AA).SubItems(428)
        !CLMSPFeeBasicMed = lstInvoice.ListItems(AA).SubItems(429)
        !CLMSPFeeLTMed = lstInvoice.ListItems(AA).SubItems(430)
        !CLMSPFeeMedInj = lstInvoice.ListItems(AA).SubItems(431)
        !CLMXRay = lstInvoice.ListItems(AA).SubItems(432)
        !CLMUltraSound = lstInvoice.ListItems(AA).SubItems(433)
        !CLMECG = lstInvoice.ListItems(AA).SubItems(434)
        !CLMDressing = lstInvoice.ListItems(AA).SubItems(435)
        !CLMNebulizer = lstInvoice.ListItems(AA).SubItems(436)
        !CLMID = lstInvoice.ListItems(AA).SubItems(437)
        !CLMTSSTO = lstInvoice.ListItems(AA).SubItems(438)
        !CLMMinorSurgery = lstInvoice.ListItems(AA).SubItems(439)
        !CLMPOthers = lstInvoice.ListItems(AA).SubItems(440)
        !CLMBloodDrawing = lstInvoice.ListItems(AA).SubItems(441)
        !CLMBloodGlucose = lstInvoice.ListItems(AA).SubItems(442)
        !CLMUrine = lstInvoice.ListItems(AA).SubItems(443)
        !CLMPapSmear = lstInvoice.ListItems(AA).SubItems(444)
        !CLMLOthers = lstInvoice.ListItems(AA).SubItems(445)
        !CLMDTPPolio = lstInvoice.ListItems(AA).SubItems(446)
        !CLMDTPPolioHIB = lstInvoice.ListItems(AA).SubItems(447)
        !CLMHepatitisB = lstInvoice.ListItems(AA).SubItems(448)
        !CLMMMR = lstInvoice.ListItems(AA).SubItems(449)
        !CLMIOthers = lstInvoice.ListItems(AA).SubItems(450)
        !CLMDental = lstInvoice.ListItems(AA).SubItems(451)
        !CLMMaternity = lstInvoice.ListItems(AA).SubItems(452)
        !CLMPRVCheckup = lstInvoice.ListItems(AA).SubItems(453)
        !CLMVision = lstInvoice.ListItems(AA).SubItems(454)
    End With
    CISExc.Update
    
    CISExc.Close
    Set CISExc = Nothing
End Sub

Private Sub SaveINVClaim()
Dim CISExc As New ADODB.Recordset
Dim CISSql As String
    
    CISSql = "SELECT * FROM Claim "
    CISExc.Open CISSql, MediConnCISWS, adOpenKeyset, adLockOptimistic
    CISExc.AddNew
    'NewCaseNo = GenerateCISCaseNo(lstInvoice.ListItems(AA).SubItems(65))

    With CISExc
        !CASNumber = NewCaseNo
        !ClaimOpenDate = Date
        !ClaimOpenBy = Logon
        !ClaimCloseDate = Date
        !ClaimClosedBy = Logon
        !ClaimLastUpdateDate = Date
        !ClaimLastUpdateUser = Logon
        !CLMPayableByMedix2C = lstInvoice.ListItems(AA).SubItems(158)
        !CLMCLNExcess = lstInvoice.ListItems(AA).SubItems(159)
        !CLMTotalActual = lstInvoice.ListItems(AA).SubItems(160)
        !CLMFinalApprove = lstInvoice.ListItems(AA).SubItems(161)
        !CLMExcess = lstInvoice.ListItems(AA).SubItems(162)
        !CLMBalanceLimit = lstInvoice.ListItems(AA).SubItems(163)
        !CLMPreviousBalance = lstInvoice.ListItems(AA).SubItems(164)
        !CLMPaidBYMedix2M = lstInvoice.ListItems(AA).SubItems(165)
        !CLMPaidBYMedix2C = lstInvoice.ListItems(AA).SubItems(166)
    End With
    CISExc.Update
    
    CISExc.Close
    Set CISExc = Nothing
End Sub

Private Function ExportToExcelError(lvw As MSComctlLib.ListView) As Boolean
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
        For i = 1 To lvw.ColumnHeaders.count
            If lvw.ColumnHeaders(i).Width <> 0 Then
                ' Create an array of visible column indexes
                intColumns = intColumns + 1
                ReDim Preserve intVisibleColumns(1 To intColumns)
                intVisibleColumns(intColumns) = i
                objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
            End If
        Next i
        If lvw.ListItems.count <> 0 Then
            ReDim strArray(1 To lvw.ListItems.count, 1 To intColumns)
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
End Function

Private Sub UpdateMBMBalance()
    Dim MBM As New ADODB.Recordset
    Dim strsql As String
    Dim FinalBalance  As Double
    Dim StrsqlAnnLmt As String
    Dim AnnLmt As New ADODB.Recordset
    Dim StrsqlNewAnnLmt As String
    Dim NewAnnLmt As New ADODB.Recordset
    Dim FinalNewBalance As Double
    Dim MBMAvailableLimit As String
    Dim TotalApprove As Double
    
    If lstInvoice.ListItems(AA).SubItems(70) <> "FULL" Then

        strsql = "Select MBMNUmber , MBMCLNBalLmt from MBMCLNOthers " & _
                " where MBMNUmber ='" & Trim(lstInvoice.ListItems(AA).SubItems(2)) & "'"
        
        If Trim(lstInvoice.ListItems(AA).SubItems(3)) <> "00" Then
            If lstInvoice.ListItems(AA).SubItems(82) = "B" Then 'SuppLimitType
                strsql = "Select MBMCNUmber , MBMCCLNBalLmt as MBMCLNBalLmt from MBMCoveredPersons " & _
                " where MBMCNUmber ='" & Trim(lstInvoice.ListItems(AA).SubItems(2)) & "' And MBMCCoverID   ='" & "01" & "'"
            ElseIf lstInvoice.ListItems(AA).SubItems(82) = "C" Then
                strsql = "Select MBMCNUmber , MBMCCLNBalLmt as MBMCLNBalLmt from MBMCoveredPersons " & _
                " where MBMCNUmber ='" & Trim(lstInvoice.ListItems(AA).SubItems(2)) & "' And MBMCCoverID   ='" & Trim(lstInvoice.ListItems(AA).SubItems(3)) & "'"
            End If
        End If
    
        MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        TotalApprove = 0
        If IsNumeric(lstInvoice.ListItems(AA).SubItems(161)) Then
            TotalApprove = lstInvoice.ListItems(AA).SubItems(161)
        End If
        
        FinalBalance = MBM!MBMCLNBalLmt - TotalApprove
        If FinalBalance < 0 Then
            MBM!MBMCLNBalLmt = 0
        Else
            MBM!MBMCLNBalLmt = FinalBalance
        End If
        
        MBM.Update
        MBM.Close
        Set MBM = Nothing
    End If
End Sub






