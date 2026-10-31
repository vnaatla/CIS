VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmWorksheet 
   Caption         =   "Worksheet"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   15
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   7215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Worksheet "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   60
         Width           =   6615
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   11775
      _ExtentX        =   20765
      _ExtentY        =   13568
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&WorkSheet"
      TabPicture(0)   =   "frmWorksheet.frx":0000
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
         TabIndex        =   5
         Top             =   7320
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid Fgrid 
         Height          =   6855
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   11535
         _ExtentX        =   20341
         _ExtentY        =   12086
         _Version        =   393216
         Rows            =   44
         Cols            =   9
         FixedCols       =   2
         BackColorFixed  =   -2147483638
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmWorksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TotalActual, TotalApprove, totalCLN, totalMem, totalExceed, Excess, totalFMemOriginal, Approved, PaytoHospital As Double
Dim TotalCLNExc As Double
Dim TotalFActual, totalFApprove, totalFCLN, totalFMem, TotalFinalCLNExc As Double
Dim CoPayApp, CoPayMBM, TotalExcess, copay, COPayColl As Double
Dim NewPLNCode As String
Dim NewPolEffDate, AdmissionDate, PolEffDate As String
Dim ClaimStatus As Boolean
Dim CLNCode As String
Dim VISInd As Boolean
Dim ANNInd As Boolean
Dim VisitLimit, BalLimit As String
Dim PAYTo As String
Dim PayCLNValue  As Double
Dim CLNBnfCard As String
Dim FinalBalLmt As Double
Dim FinalVisitLmt As Double
Dim CLNRedStatus As String
Dim CLNRedAmt As Double
Dim AdjAnnLmt As Double
Dim NONGPAmt As Double
Dim GPStatus As Boolean
Dim TmpCLNCode As String


Dim I1, I2, I3, I4, I5, I6 As Double
Dim I7, I8, I9, I10, I11, I12 As Double
Dim I13, I14, I15, I16, I17, I18, I19 As Double
Dim I20, I21, I22, I23, I24, I25, I26, I27, I28, I29, I30, I31 As Double

Dim CB1, CB2, CB3, CB4, CB5, CB6 As Double
Dim CB7, CB8, CB9, CB10, CB11, CB12 As Double
Dim CB13, CB14, CB15, CB16, CB17, CB18, CB19 As Double
Dim CB20, CB21, CB22, CB23, CB24, CB25, CB26, CB27, CB28, CB29, CB30, CB31 As Double

Dim ComPLNCode As String
Dim CLMPByM2M As Double
Dim FTotalMember As Double
Dim BNFStatus As Boolean
Dim ClaimOpen As String
Dim bExit As Boolean
Dim CLMVersionNo As Integer
Dim debitNo As String
Dim creditno As String
Dim HdebitNo As String
Dim HCreditNo As String
Dim MCreditNo As String
Dim bChangesMade As Boolean
Dim RBDays As String
Dim ICUDays As String
Dim CashDays, GuardianDays, InHosPhycisianDays, KidneyDCDays As String
Dim strsqlSameLmt As String
Dim SameLmt As New ADODB.Recordset
Dim EXCNumber As String
Dim HOSCode As String
Dim COPreHosp, COMPreHosp As Double
Dim PLNPercentage As String
Dim PLNCode As String
Dim INSCode As String
Dim HLTCode As String
' Approved
Dim App1, App2, App3, App4, App5 As String
Dim App6, App7, App8, App9, App10 As String
Dim App11, App12, App13, App14, App15 As String
Dim App16, App17, App18, App19, App20 As String
Dim App21, App22, App23, App24, APP25, APP26, App27, APP28, APP29, APP30, APP31 As String

'CalReimburse
Dim RGPFeeBMed As Double
Dim RGPFeeLTMed As Double
Dim RGPFeeMedInj As Double
Dim RSPFeeBMed As Double
Dim RSPFeeLTMed As Double
Dim RSPFeeMedInj As Double
Dim RDental As Double
Dim RVision As Double
Dim RMaternity As Double
Dim RPRVCheckup As Double
Dim RXRay As Double
Dim RUltraSound As Double
Dim RECG As Double
Dim RDressing As Double
Dim RNebulizer As Double
Dim RID As Double
Dim RTSSTO As Double
Dim RMinorSurgery As Double
Dim RPOthers As Double
Dim RBloodDrawing As Double
Dim RBloodGlucose As Double
Dim RUrine As Double
Dim RPapSmear As Double
Dim RLOthers As Double
Dim RDTPPolio As Double
Dim RDTPPolioHIB As Double
Dim RHepatitisB As Double
Dim RMMR As Double
Dim RIOthers As Double
Dim RTyphoidNoDed As Double
Dim ROthersNoDed As Double

'CalCashless
Dim CGPFeeBMed As Double
Dim CGPFeeLTMed As Double
Dim CGPFeeMedInj As Double
Dim CSPFeeBMed As Double
Dim CSPFeeLTMed As Double
Dim CSPFeeMedInj As Double
Dim CDental As Double
Dim CVision As Double
Dim CMaternity As Double
Dim CPRVCheckup As Double
Dim CXRay As Double
Dim CUltraSound As Double
Dim CECG As Double
Dim CDressing As Double
Dim CNebulizer As Double
Dim CID As Double
Dim CTSSTO As Double
Dim CMinorSurgery As Double
Dim CPOthers As Double
Dim CBloodDrawing As Double
Dim CBloodGlucose As Double
Dim CUrine As Double
Dim CPapSmear As Double
Dim CLOthers As Double
Dim CDTPPolio As Double
Dim CDTPPolioHIB As Double
Dim CHepatitisB As Double
Dim CMMR As Double
Dim CIOthers As Double
Dim CTyphoidNoDed As Double
Dim COthersNoDed As Double

'CalTotal
Dim TRBDay As Double
Dim TRBLmt As Double
Dim TICUDay As Double
Dim TICULmt As Double
Dim TGurdianDay As Double
Dim TGurdianLmt As Double
Dim TGorvDay As Double
Dim TGorvLmt As Double
Dim CBInd As String
Dim ClickTExcess As Boolean
Public InvoiceClick As Integer
Public INVCLNForm As Integer
Public PostageStatus As Boolean
Dim MCOProcStatus As Boolean
Dim MCOProcFee As Double
Dim MCOProcFeePerc As String

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub FGrid_Click()
    Dim temp, tempCol, tempRow, tempTitle As String
    Dim Appvalue, PayHosValue, ActualValue As Double
    Dim Col7, Row32, ClickGrid As Integer
    
    
    If (Fgrid.Col = 3 And Fgrid.Row <= 31) Or (Fgrid.Col = 6 And Fgrid.Row <= 31) And Fgrid.Row <> 0 Then
        If Fgrid.Col = 3 Then
            ClickGrid = 3
        ElseIf Fgrid.Col = 6 Then
            ClickGrid = 6
        End If
            tempCol = Fgrid.Col
            Fgrid.Col = 1
            tempTitle = Fgrid.Text
            temp = InputBox(tempTitle, Fgrid.Text)
            Fgrid.Col = tempCol
            
            
            If temp <> "" Then
                If IsNumeric(temp) = True Then
                    Fgrid = Format(temp, "#,#0.00;(#,#0.00)")
                Else
                    Fgrid = UCase(temp)
                End If
                
                Fgrid.Col = 2
                If IsNumeric(Fgrid.Text) Then
                    ActualValue = Fgrid.Text
                Else
                    ActualValue = 0
                End If
                
                If ClickGrid = 3 Then
                    Fgrid.Col = 4
                    If IsNumeric(temp) = True Then
                        Fgrid.Text = GetMin(CDbl(temp), CDbl(ActualValue))
                     Else
                        Fgrid.Text = Format(ActualValue, "#,#.00")
                    End If
                    
                    Appvalue = Fgrid.Text
                    
                    Fgrid.Col = 5
                    Fgrid.Text = DueToFromCLN((ActualValue), (Appvalue))
                    
            ElseIf ClickGrid = 6 Then
                    Fgrid.Col = 7
                    If IsNumeric(temp) = True Then
                        Fgrid.Text = GetMin(CDbl(temp), CDbl(ActualValue))
                     Else
                        Fgrid.Text = Format(ActualValue, "#,#.00")
                    End If
                    
                    Appvalue = Fgrid.Text
                    Fgrid.Col = 8
                    Fgrid.Text = DueToFromMBM((ActualValue), (Appvalue))
            
                    
            End If
                    'Approve
                    
                    Fgrid.Col = 7
                    Fgrid = getTotal(7)
                    TotalApprove = Fgrid
                    
                    Fgrid.Col = 8
                    Fgrid = getTotal(8)
                    If IsNumeric(Fgrid) Then
                        totalMem = Fgrid
                    Else
                        totalMem = 0
                    End If
                    
                    Fgrid.Col = 4
                    Fgrid = getTotal(4)
                    totalCLN = Fgrid
                    
                    Fgrid.Col = 5
                    Fgrid = getTotal(5)
                    If IsNumeric(Fgrid) = True Then
                        TotalCLNExc = Fgrid
                    Else
                        TotalCLNExc = 0
                    End If
            
                Call getFinalTotal
            End If
        End If
        Call SaveCLMActual
        Call SaveCLMContRate
        Call SavePayCLN
        Call SaveNotPayCLN
        Call SaveMBMBnfLmt
        Call SaveMBMApp
        Call SaveMBMExc
        Call SaveClaim
    

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
Call ActionLogUpdate(12, "frmWorksheet")
Dim CAS As New ADODB.Recordset
Dim StrSql, sqlPolStatus As String
Dim PolStatus As New ADODB.Recordset
Dim OpenNew
Dim StrsqlAnnLmt As String
Dim AnnLmt As New ADODB.Recordset
Dim strsqlCOV As String
Dim MBMCov As New ADODB.Recordset
Dim strsqlCLM As String
Dim CLM As New ADODB.Recordset
Dim strsqlPLN As String
Dim PLN As New ADODB.Recordset
Dim strsqlDiag As String
Dim Diag As New ADODB.Recordset
Dim DiagInd As String
Dim DiagRec As Integer

    BNFcombolisting
    
        Screen.MousePointer = vbHourglass
        Call ClearEntries
        TotalFActual = 0
        totalFApprove = 0
        totalFCLN = 0
        totalFMem = 0
        TotalFinalCLNExc = 0
        GPStatus = False
        'Call Inv Clinic form
        If INVCLNForm = 1 Then
             If frmInvoice.lstInvoice.ListItems.Count > 0 Then
                'click on invoice update button
                If InvoiceClick = 1 Then
                
                    PLNCode = frmInvoice.lstRec.SubItems(67)
                    INSCode = frmInvoice.lstRec.SubItems(68)
                    HLTCode = frmInvoice.lstRec.SubItems(69)
                    CLNCode = frmInvoice.lstRec.SubItems(12)
                    CLNBnfCard = frmInvoice.lstRec.SubItems(72)
                    BalLimit = frmInvoice.lstRec.SubItems(70)
                    VisitLimit = frmInvoice.lstRec.SubItems(71)
                    If frmInvoice.lstRec.SubItems(74) = "N" Then
                        ANNInd = False
                    Else
                        ANNInd = True
                    End If
                    If frmInvoice.lstRec.SubItems(73) = "N" Then
                        VISInd = False
                    Else
                        VISInd = True
                    End If
                    PAYTo = frmInvoice.lstRec.SubItems(80)
                    COPayColl = IIf(IsNumeric(frmInvoice.lstRec.SubItems(43)) = True, frmInvoice.lstRec.SubItems(43), 0)
                    copay = IIf(IsNumeric(frmInvoice.lstRec.SubItems(75)) = True, frmInvoice.lstRec.SubItems(75), 0)
                    AdjAnnLmt = IIf(IsNumeric(frmInvoice.lstRec.SubItems(113)) = True, frmInvoice.lstRec.SubItems(113), 0)
                    If CDbl(frmInvoice.lstRec.SubItems(91)) > 0 Then
                        GPStatus = True
                    Else
                        GPStatus = False
                    End If
                    If Trim(frmInvoice.lstRec.SubItems(469)) = "Y" Then
                        MCOProcStatus = True
                    Else
                        MCOProcStatus = False
                    End If
                    
                Else
                    'click on invoice compute button
                    PLNCode = frmInvoice.lstInvoice.SelectedItem.SubItems(67)
                    INSCode = frmInvoice.lstInvoice.SelectedItem.SubItems(68)
                    HLTCode = frmInvoice.lstInvoice.SelectedItem.SubItems(69)
                    CLNCode = frmInvoice.lstInvoice.SelectedItem.SubItems(12)
                    CLNBnfCard = frmInvoice.lstInvoice.SelectedItem.SubItems(72)
                    BalLimit = frmInvoice.lstInvoice.SelectedItem.SubItems(70)
                    VisitLimit = frmInvoice.lstInvoice.SelectedItem.SubItems(71)
                    If frmInvoice.lstInvoice.SelectedItem.SubItems(74) = "N" Then
                        ANNInd = False
                    Else
                        ANNInd = True
                    End If
                    If frmInvoice.lstInvoice.SelectedItem.SubItems(73) = "N" Then
                        VISInd = False
                    Else
                        VISInd = True
                    End If
                    PAYTo = frmInvoice.lstInvoice.SelectedItem.SubItems(80)
                    COPayColl = IIf(IsNumeric(frmInvoice.lstInvoice.SelectedItem.SubItems(43)) = True, frmInvoice.lstInvoice.SelectedItem.SubItems(43), 0)
                    copay = IIf(IsNumeric(frmInvoice.lstInvoice.SelectedItem.SubItems(75)) = True, frmInvoice.lstInvoice.SelectedItem.SubItems(75), 0)
                    AdjAnnLmt = IIf(IsNumeric(frmInvoice.lstInvoice.SelectedItem.SubItems(113)) = True, frmInvoice.lstInvoice.SelectedItem.SubItems(113), 0)
                    If CDbl(frmInvoice.lstInvoice.SelectedItem.SubItems(91)) > 0 Then
                        GPStatus = True
                    Else
                        GPStatus = False
                    End If
                    If Trim(frmInvoice.lstInvoice.SelectedItem.SubItems(469)) = "Y" Then
                        MCOProcStatus = True
                    Else
                        MCOProcStatus = False
                    End If
                End If
            End If
        'call Inv MBM Form
        ElseIf INVCLNForm = 2 Then
             If frmInvoiceMBM.lstInvoice.ListItems.Count > 0 Then
                If InvoiceClick = 1 Then
                    PLNCode = frmInvoiceMBM.lstRec.SubItems(67)
                    INSCode = frmInvoiceMBM.lstRec.SubItems(68)
                    HLTCode = frmInvoiceMBM.lstRec.SubItems(69)
                    CLNCode = frmInvoiceMBM.lstRec.SubItems(12)
                    CLNBnfCard = frmInvoiceMBM.lstRec.SubItems(72)
                    BalLimit = frmInvoiceMBM.lstRec.SubItems(70)
                    VisitLimit = frmInvoiceMBM.lstRec.SubItems(71)
                    If frmInvoiceMBM.lstRec.SubItems(74) = "N" Then
                        ANNInd = False
                    Else
                        ANNInd = True
                    End If
                    If frmInvoiceMBM.lstRec.SubItems(73) = "N" Then
                        VISInd = False
                    Else
                        VISInd = True
                    End If
                    PAYTo = frmInvoiceMBM.lstRec.SubItems(80)
                    COPayColl = IIf(IsNumeric(frmInvoiceMBM.lstRec.SubItems(43)) = True, frmInvoiceMBM.lstRec.SubItems(43), 0)
                    copay = IIf(IsNumeric(frmInvoiceMBM.lstRec.SubItems(75)) = True, frmInvoiceMBM.lstRec.SubItems(75), 0)
                    PostageStatus = IIf(frmInvoiceMBM.lstRec.SubItems(108) = "True", True, False)
                    AdjAnnLmt = IIf(IsNumeric(frmInvoiceMBM.lstRec.SubItems(113)) = True, frmInvoiceMBM.lstRec.SubItems(113), 0)

                Else
                    PLNCode = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(67)
                    INSCode = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(68)
                    HLTCode = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(69)
                    CLNCode = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(12)
                    CLNBnfCard = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(72)
                    BalLimit = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(70)
                    VisitLimit = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(71)
                    If frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(74) = "N" Then
                        ANNInd = False
                    Else
                        ANNInd = True
                    End If
                    If frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(73) = "N" Then
                        VISInd = False
                    Else
                        VISInd = True
                    End If
                    PAYTo = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(80)
                    COPayColl = IIf(IsNumeric(frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(43)) = True, frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(43), 0)
                    copay = IIf(IsNumeric(frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(75)) = True, frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(75), 0)
                    PostageStatus = IIf(frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(108) = "True", True, False)
                    AdjAnnLmt = IIf(IsNumeric(frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(113)) = True, frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(113), 0)
                End If
           End If
           
           
        'Call Inv Adj Form
        ElseIf INVCLNForm = 3 Then
             If frmInvoiceAdj.lstInvoice.ListItems.Count > 0 Then
                 PLNCode = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(67)
                 INSCode = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(68)
                 HLTCode = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(69)
                 CLNCode = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(12)
                 CLNBnfCard = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(72)
                 BalLimit = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(164)
                 VisitLimit = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(71)
                 If frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(74) = "N" Then
                     ANNInd = False
                 Else
                     ANNInd = True
                 End If
                 If frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(73) = "N" Then
                     VISInd = False
                 Else
                     VISInd = True
                 End If
                 PAYTo = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(80)
                 COPayColl = IIf(IsNumeric(frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(43)) = True, frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(43), 0)
                 copay = IIf(IsNumeric(frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(75)) = True, frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(75), 0)
                 AdjAnnLmt = IIf(IsNumeric(frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(113)) = True, frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(113), 0)
                If CDbl(frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(91)) > 0 Then
                    GPStatus = True
                Else
                    GPStatus = False
                End If
                If Trim(frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(469)) = "Y" Then
                    MCOProcStatus = True
                Else
                    MCOProcStatus = False
                End If

           End If
        ElseIf INVCLNForm = 4 Then
             If frmInvoiceAdjAftBordx.lstInvoice.ListItems.Count > 0 Then
                 PLNCode = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(67)
                 INSCode = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(68)
                 HLTCode = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(69)
                 CLNCode = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(12)
                 CLNBnfCard = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(72)
                 BalLimit = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(164)
                 VisitLimit = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(71)
                 If frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(74) = "N" Then
                     ANNInd = False
                 Else
                     ANNInd = True
                 End If
                 If frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(73) = "N" Then
                     VISInd = False
                 Else
                     VISInd = True
                 End If
                 PAYTo = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(80)
                 COPayColl = IIf(IsNumeric(frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(43)) = True, frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(43), 0)
                 copay = IIf(IsNumeric(frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(75)) = True, frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(75), 0)
                 AdjAnnLmt = IIf(IsNumeric(frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(113)) = True, frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(113), 0)
                If CDbl(frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(91)) > 0 Then
                    GPStatus = True
                Else
                    GPStatus = False
                End If
                If Trim(frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(469)) = "Y" Then
                    MCOProcStatus = True
                Else
                    MCOProcStatus = False
                End If

           End If
        
        ElseIf INVCLNForm = 5 Then
             If frmInvoiceAdjAfterPost.lstInvoice.ListItems.Count > 0 Then
                 PLNCode = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(67)
                 INSCode = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(68)
                 HLTCode = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(69)
                 CLNCode = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(12)
                 CLNBnfCard = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(72)
                 BalLimit = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(164)
                 VisitLimit = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(71)
                 If frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(74) = "N" Then
                     ANNInd = False
                 Else
                     ANNInd = True
                 End If
                 If frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(73) = "N" Then
                     VISInd = False
                 Else
                     VISInd = True
                 End If
                 PAYTo = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(80)
                 COPayColl = IIf(IsNumeric(frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(43)) = True, frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(43), 0)
                 copay = IIf(IsNumeric(frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(75)) = True, frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(75), 0)
                 AdjAnnLmt = IIf(IsNumeric(frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(113)) = True, frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(113), 0)
                If CDbl(frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(91)) > 0 Then
                    GPStatus = True
                Else
                    GPStatus = False
                End If
                If Trim(frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(469)) = "Y" Then
                    MCOProcStatus = True
                Else
                    MCOProcStatus = False
                End If

           End If
        
        End If
            
                If MCOProcStatus = True Then
                    Call GetCLNItems
                End If
                
                Call showInvoice
                Call display_benefit
                
                If PAYTo = "C" Then
                    Call display_CLNbenefit
                End If
                
                Call getDatatoWS
                Call SaveCLMActual
                Call SaveCLMContRate
                Call SavePayCLN
                Call SaveNotPayCLN
                Call SaveMBMBnfLmt
                Call SaveMBMApp
                Call SaveMBMExc
                Call SaveClaim
    Screen.MousePointer = vbDefault
    Exit Sub
ErrorShow:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description

End Sub

Private Sub BNFcombolisting()
Screen.MousePointer = vbHourglass

Dim subCat As New ADODB.Recordset
Dim Cnt As Integer
Call DoInitialSettings

    subCat.Open "select BNFCATCode , BNFCATDescription   from BNFCategoryNew order by BNFCATCode", MediConnCISDB, adOpenKeyset, adLockOptimistic
    Cnt = 1
    Fgrid.Rows = subCat.RecordCount + 9
    
    Do Until subCat.EOF
         Fgrid.Col = 0
         Fgrid.Row = Cnt
         Fgrid.Text = subCat!BNFCatCode
         
         Fgrid.Col = 1
         Fgrid.Row = Cnt
         Fgrid.Text = subCat!BNFCATDescription
    
         Cnt = Cnt + 1
         subCat.MoveNext
    Loop
Screen.MousePointer = Default
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
' Benefits
Sub DoInitialSettings()
    Dim i, Y    As Integer
    Fgrid.MergeCol(1) = True     ' Allow merge on Columns 0 thru 3
    
    Fgrid.ColAlignment(1) = 1
    Fgrid.MergeCells = flexMergeRestrictColumns
    Fgrid.ColAlignment(3) = 7
    Fgrid.ColAlignment(4) = 7
    Fgrid.ColAlignment(5) = 7
    Fgrid.ColAlignment(6) = 7
    Fgrid.ColAlignment(7) = 7
    Fgrid.ColAlignment(8) = 7
    Fgrid.Row = 0
    'Fgrid.Font = Bold
    Fgrid.ColWidth(0) = 0     ' Set column's width
    Fgrid.ColWidth(1) = 2200
    Fgrid.ColWidth(2) = 1000
    Fgrid.ColWidth(3) = 1400
    Fgrid.ColWidth(4) = 1200
    Fgrid.ColWidth(5) = 1500
    Fgrid.ColWidth(6) = 1000
    Fgrid.ColWidth(7) = 1400
    Fgrid.ColWidth(8) = 1000

    Fgrid.Col = 1
    Fgrid = "Benfits"
    Fgrid.Col = 2
    Fgrid = "Actual"
    Fgrid.Col = 3
    Fgrid = "Contracted Rate"
    Fgrid.Col = 4
    Fgrid = "Payable to CLN"
    Fgrid.Col = 5
    Fgrid = "Not Payable to CLN"
    Fgrid.Col = 6
    Fgrid = "Benefit Lmt"
    Fgrid.Col = 7
    Fgrid = "Approved Benefit"
    Fgrid.Col = 8
    Fgrid = "Excess"
    
    Fgrid.Col = 1
    Fgrid.Row = 33
    Fgrid = "Total"

    Fgrid.Row = 34
    Fgrid = "Amount Not Contracted"
    
    Fgrid.Row = 35
    Fgrid = "Less Co-Payment"
    
    Fgrid.Row = 36
    Fgrid = "Visit Limit"
    
    Fgrid.Row = 37
    Fgrid = "Balance Limit"
    
    Fgrid.Row = 38
    Fgrid = "MCO Proc. Fee"
    
    Fgrid.Row = 39
    Fgrid = "Final Approved"
    
End Sub

Private Sub ClearEntries()
    Dim ctl As Control
    Dim FCol, Frow As Integer
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            If ctl.Name <> "txtCASNumber" Then
                ctl.Text = ""
            End If
        End If
    Next
    
    For FCol = 3 To 8
        Fgrid.Col = FCol
        For Frow = 1 To 31
            Fgrid.Row = Frow
            Fgrid = "-"
        Next
    Next
        For FCol = 2 To 8
        Fgrid.Col = FCol
        
        For Frow = 33 To 39
            Fgrid.CellFontBold = True
            Fgrid.CellForeColor = vbBlue
            Fgrid.Row = Frow
            Fgrid = "-"
        Next
    Next
End Sub

Private Sub showInvoice()
    Dim Invoice As New ADODB.Recordset
    Dim INVlist As ListItem
    Dim StrSql As String
    I1 = 0
    I2 = 0
    I3 = 0
    I4 = 0
    I5 = 0
    I6 = 0
    I7 = 0
    I8 = 0
    I9 = 0
    I10 = 0
    I11 = 0
    I12 = 0
    I13 = 0
    I14 = 0
    I15 = 0
    I16 = 0
    I17 = 0
    I18 = 0
    I19 = 0
    I20 = 0
    I21 = 0
    I22 = 0
    I23 = 0
    I24 = 0
    I25 = 0
    I26 = 0
    I27 = 0
    I28 = 0
    I29 = 0
    I30 = 0
    I31 = 0
    NONGPAmt = 0
    
    If INVCLNForm = 1 Then
        If frmInvoice.lstInvoice.ListItems.Count > 0 Then
            If InvoiceClick = 1 Then
                I1 = frmInvoice.lstRec.SubItems(15)
                I2 = frmInvoice.lstRec.SubItems(16)
                I3 = frmInvoice.lstRec.SubItems(17)
                I4 = frmInvoice.lstRec.SubItems(18)
                I5 = frmInvoice.lstRec.SubItems(19)
                I6 = frmInvoice.lstRec.SubItems(20)
                I7 = frmInvoice.lstRec.SubItems(24)
                I8 = frmInvoice.lstRec.SubItems(25)
                I9 = frmInvoice.lstRec.SubItems(26)
                I10 = frmInvoice.lstRec.SubItems(27)
                I11 = frmInvoice.lstRec.SubItems(28)
                I12 = frmInvoice.lstRec.SubItems(29)
                I13 = frmInvoice.lstRec.SubItems(30)
                I14 = frmInvoice.lstRec.SubItems(31)
                I15 = frmInvoice.lstRec.SubItems(32)
                I16 = frmInvoice.lstRec.SubItems(33)
                I17 = frmInvoice.lstRec.SubItems(34)
                I18 = frmInvoice.lstRec.SubItems(35)
                I19 = frmInvoice.lstRec.SubItems(36)
                I20 = frmInvoice.lstRec.SubItems(37)
                I21 = frmInvoice.lstRec.SubItems(38)
                I22 = frmInvoice.lstRec.SubItems(39)
                I23 = frmInvoice.lstRec.SubItems(40)
                I24 = frmInvoice.lstRec.SubItems(41)
                I25 = frmInvoice.lstRec.SubItems(42)
                I26 = frmInvoice.lstRec.SubItems(21)
                I27 = frmInvoice.lstRec.SubItems(22)
                I28 = frmInvoice.lstRec.SubItems(23)
                I29 = frmInvoice.lstRec.SubItems(101)
                I30 = IIf(IsNumeric(frmInvoice.lstRec.SubItems(474)), frmInvoice.lstRec.SubItems(474), 0)
                I31 = IIf(IsNumeric(frmInvoice.lstRec.SubItems(475)), frmInvoice.lstRec.SubItems(475), 0)
                
                NONGPAmt = IIf(IsNumeric(frmInvoice.lstRec.SubItems(114)) = True, frmInvoice.lstRec.SubItems(114), 0)
            Else
                I1 = frmInvoice.lstInvoice.SelectedItem.SubItems(15)
                I2 = frmInvoice.lstInvoice.SelectedItem.SubItems(16)
                I3 = frmInvoice.lstInvoice.SelectedItem.SubItems(17)
                I4 = frmInvoice.lstInvoice.SelectedItem.SubItems(18)
                I5 = frmInvoice.lstInvoice.SelectedItem.SubItems(19)
                I6 = frmInvoice.lstInvoice.SelectedItem.SubItems(20)
                I7 = frmInvoice.lstInvoice.SelectedItem.SubItems(24)
                I8 = frmInvoice.lstInvoice.SelectedItem.SubItems(25)
                I9 = frmInvoice.lstInvoice.SelectedItem.SubItems(26)
                I10 = frmInvoice.lstInvoice.SelectedItem.SubItems(27)
                I11 = frmInvoice.lstInvoice.SelectedItem.SubItems(28)
                I12 = frmInvoice.lstInvoice.SelectedItem.SubItems(29)
                I13 = frmInvoice.lstInvoice.SelectedItem.SubItems(30)
                I14 = frmInvoice.lstInvoice.SelectedItem.SubItems(31)
                I15 = frmInvoice.lstInvoice.SelectedItem.SubItems(32)
                I16 = frmInvoice.lstInvoice.SelectedItem.SubItems(33)
                I17 = frmInvoice.lstInvoice.SelectedItem.SubItems(34)
                I18 = frmInvoice.lstInvoice.SelectedItem.SubItems(35)
                I19 = frmInvoice.lstInvoice.SelectedItem.SubItems(36)
                I20 = frmInvoice.lstInvoice.SelectedItem.SubItems(37)
                I21 = frmInvoice.lstInvoice.SelectedItem.SubItems(38)
                I22 = frmInvoice.lstInvoice.SelectedItem.SubItems(39)
                I23 = frmInvoice.lstInvoice.SelectedItem.SubItems(40)
                I24 = frmInvoice.lstInvoice.SelectedItem.SubItems(41)
                I25 = frmInvoice.lstInvoice.SelectedItem.SubItems(42)
                I26 = frmInvoice.lstInvoice.SelectedItem.SubItems(21)
                I27 = frmInvoice.lstInvoice.SelectedItem.SubItems(22)
                I28 = frmInvoice.lstInvoice.SelectedItem.SubItems(23)
                I29 = frmInvoice.lstInvoice.SelectedItem.SubItems(101)
                I30 = IIf(IsNumeric(frmInvoice.lstInvoice.SelectedItem.SubItems(474)), frmInvoice.lstInvoice.SelectedItem.SubItems(474), 0)
                I31 = IIf(IsNumeric(frmInvoice.lstInvoice.SelectedItem.SubItems(475)), frmInvoice.lstInvoice.SelectedItem.SubItems(475), 0)

                NONGPAmt = IIf(IsNumeric(frmInvoice.lstInvoice.SelectedItem.SubItems(114)) = True, frmInvoice.lstInvoice.SelectedItem.SubItems(114), 0)
            End If
        End If
    ElseIf INVCLNForm = 2 Then
        If frmInvoiceMBM.lstInvoice.ListItems.Count > 0 Then
            If InvoiceClick = 1 Then
                I1 = frmInvoiceMBM.lstRec.SubItems(15)
                I2 = frmInvoiceMBM.lstRec.SubItems(16)
                I3 = frmInvoiceMBM.lstRec.SubItems(17)
                I4 = frmInvoiceMBM.lstRec.SubItems(18)
                I5 = frmInvoiceMBM.lstRec.SubItems(19)
                I6 = frmInvoiceMBM.lstRec.SubItems(20)
                I7 = frmInvoiceMBM.lstRec.SubItems(24)
                I8 = frmInvoiceMBM.lstRec.SubItems(25)
                I9 = frmInvoiceMBM.lstRec.SubItems(26)
                I10 = frmInvoiceMBM.lstRec.SubItems(27)
                I11 = frmInvoiceMBM.lstRec.SubItems(28)
                I12 = frmInvoiceMBM.lstRec.SubItems(29)
                I13 = frmInvoiceMBM.lstRec.SubItems(30)
                I14 = frmInvoiceMBM.lstRec.SubItems(31)
                I15 = frmInvoiceMBM.lstRec.SubItems(32)
                I16 = frmInvoiceMBM.lstRec.SubItems(33)
                I17 = frmInvoiceMBM.lstRec.SubItems(34)
                I18 = frmInvoiceMBM.lstRec.SubItems(35)
                I19 = frmInvoiceMBM.lstRec.SubItems(36)
                I20 = frmInvoiceMBM.lstRec.SubItems(37)
                I21 = frmInvoiceMBM.lstRec.SubItems(38)
                I22 = frmInvoiceMBM.lstRec.SubItems(39)
                I23 = frmInvoiceMBM.lstRec.SubItems(40)
                I24 = frmInvoiceMBM.lstRec.SubItems(41)
                I25 = frmInvoiceMBM.lstRec.SubItems(42)
                I26 = frmInvoiceMBM.lstRec.SubItems(21)
                I27 = frmInvoiceMBM.lstRec.SubItems(22)
                I28 = frmInvoiceMBM.lstRec.SubItems(23)
                I29 = frmInvoiceMBM.lstRec.SubItems(101)
                I30 = IIf(IsNumeric(frmInvoiceMBM.lstRec.SubItems(474)), frmInvoiceMBM.lstRec.SubItems(474), 0)
                I31 = IIf(IsNumeric(frmInvoiceMBM.lstRec.SubItems(475)), frmInvoiceMBM.lstRec.SubItems(475), 0)
                
                NONGPAmt = IIf(IsNumeric(frmInvoiceMBM.lstRec.SubItems(114)) = True, frmInvoiceMBM.lstRec.SubItems(114), 0)
            Else
                I1 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(15)
                I2 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(16)
                I3 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(17)
                I4 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(18)
                I5 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(19)
                I6 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(20)
                I7 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(24)
                I8 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(25)
                I9 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(26)
                I10 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(27)
                I11 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(28)
                I12 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(29)
                I13 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(30)
                I14 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(31)
                I15 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(32)
                I16 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(33)
                I17 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(34)
                I18 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(35)
                I19 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(36)
                I20 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(37)
                I21 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(38)
                I22 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(39)
                I23 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(40)
                I24 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(41)
                I25 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(42)
                I26 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(21)
                I27 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(22)
                I28 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(23)
                I29 = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(101)
                I30 = IIf(IsNumeric(frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(474)), frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(474), 0)
                I31 = IIf(IsNumeric(frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(475)), frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(475), 0)
                
                NONGPAmt = IIf(IsNumeric(frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(114)) = True, frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(114), 0)
            End If
        End If
    ElseIf INVCLNForm = 3 Then
        If frmInvoiceAdj.lstInvoice.ListItems.Count > 0 Then
            I1 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(15)
            I2 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(16)
            I3 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(17)
            I4 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(18)
            I5 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(19)
            I6 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(20)
            I7 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(24)
            I8 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(25)
            I9 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(26)
            I10 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(27)
            I11 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(28)
            I12 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(29)
            I13 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(30)
            I14 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(31)
            I15 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(32)
            I16 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(33)
            I17 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(34)
            I18 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(35)
            I19 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(36)
            I20 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(37)
            I21 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(38)
            I22 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(39)
            I23 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(40)
            I24 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(41)
            I25 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(42)
            I26 = 0
            I27 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(21)
            I28 = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(22)
            I29 = IIf(IsNumeric(frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(95)) = True, frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(95), 0)
            I30 = IIf(IsNumeric(frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(473)), frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(473), 0)
            I31 = IIf(IsNumeric(frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(474)), frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(474), 0)
            NONGPAmt = IIf(IsNumeric(frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(114)) = True, frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(114), 0)
        End If
    ElseIf INVCLNForm = 4 Then
        If frmInvoiceAdjAftBordx.lstInvoice.ListItems.Count > 0 Then
            I1 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(15)
            I2 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(16)
            I3 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(17)
            I4 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(18)
            I5 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(19)
            I6 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(20)
            I7 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(24)
            I8 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(25)
            I9 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(26)
            I10 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(27)
            I11 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(28)
            I12 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(29)
            I13 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(30)
            I14 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(31)
            I15 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(32)
            I16 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(33)
            I17 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(34)
            I18 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(35)
            I19 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(36)
            I20 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(37)
            I21 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(38)
            I22 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(39)
            I23 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(40)
            I24 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(41)
            I25 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(42)
            I26 = 0
            I27 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(21)
            I28 = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(22)
            I29 = IIf(IsNumeric(frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(95)) = True, frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(95), 0)
            I30 = IIf(IsNumeric(frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(473)), frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(473), 0)
            I31 = IIf(IsNumeric(frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(474)), frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(474), 0)
            NONGPAmt = IIf(IsNumeric(frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(114)) = True, frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(114), 0)
        End If
        
    ElseIf INVCLNForm = 5 Then
        If frmInvoiceAdjAfterPost.lstInvoice.ListItems.Count > 0 Then
            I1 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(15)
            I2 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(16)
            I3 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(17)
            I4 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(18)
            I5 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(19)
            I6 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(20)
            I7 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(24)
            I8 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(25)
            I9 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(26)
            I10 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(27)
            I11 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(28)
            I12 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(29)
            I13 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(30)
            I14 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(31)
            I15 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(32)
            I16 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(33)
            I17 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(34)
            I18 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(35)
            I19 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(36)
            I20 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(37)
            I21 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(38)
            I22 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(39)
            I23 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(40)
            I24 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(41)
            I25 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(42)
            I26 = 0
            I27 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(21)
            I28 = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(22)
            I29 = IIf(IsNumeric(frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(95)) = True, frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(95), 0)
            I30 = IIf(IsNumeric(frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(473)), frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(473), 0)
            I31 = IIf(IsNumeric(frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(474)), frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(474), 0)
            NONGPAmt = IIf(IsNumeric(frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(114)) = True, frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(114), 0)
        End If
        
    End If
End Sub

Public Sub display_CLNbenefit()
    Dim strsqlCOPay As String
    Dim copay As New ADODB.Recordset
    Dim CLNBNF As New ADODB.Recordset
    Dim StrSql As String
    Dim counter As Integer
    Dim Lmt As Integer
    Dim CLN As New ADODB.Recordset
    Dim strsqlCLN As String
    Dim BNFType As String
    Dim strsqlBNFSpe As String
    Dim BNFSpe As New ADODB.Recordset
    Dim a As Integer
    BNFType = ""
    CB1 = 0
    CB2 = 0
    CB3 = 0
    CB4 = 0
    CB5 = 0
    CB6 = 0
    CB7 = 0
    CB8 = 0
    CB9 = 0
    CB10 = 0
    CB11 = 0
    CB12 = 0
    CB13 = 0
    CB14 = 0
    CB15 = 0
    CB16 = 0
    CB17 = 0
    CB18 = 0
    CB19 = 0
    CB20 = 0
    CB21 = 0
    CB22 = 0
    CB23 = 0
    CB24 = 0
    CB25 = 0
    CB26 = 0
    CB27 = 0
    CB28 = 0
    CB29 = 0
    CB30 = 0
    CB31 = 0
    
        strsqlCLN = "Select CLNCode ,CLNName, CLNBNFType, CLNRedStatus from CLN  " & _
        " where CLNCode = '" & Trim(CLNCode) & "'"
        CLN.Open strsqlCLN, medixmasterconn, adOpenKeyset, adLockOptimistic
        If CLN.RecordCount Then
            If Len(CLN!CLNBNFType) Then
                CBInd = CLN!CLNBNFType
            End If
            
            If Len(CLN!CLNRedStatus) Then
                CLNRedStatus = CLN!CLNRedStatus
            End If
        End If
        CLN.Close
        Set CLN = Nothing
        If INVCLNForm = 1 Then
            If CBInd = "2" Then
                StrSql = "Select BNFCode ,CLNBNFRMPerVisit AS BNFRMPerVisit, CLNBNFEffDate from STDClinicBNF  "
                StrSql = StrSql & " order by BNFCOde "
                CLNBNF.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
                
                strsqlBNFSpe = "Select BNFCode ,CLNBNFRMPerVisit AS BNFRMPerVisit, CLNBNFEffDate,CLNBNFType from STDClinicBNFSPE  " & _
                " where CLNBNFType = '" & Trim(CLNBnfCard) & "'"
                strsqlBNFSpe = strsqlBNFSpe & " order by BNFCOde "
                BNFSpe.Open strsqlBNFSpe, MediConnCISDB, adOpenKeyset, adLockOptimistic
            End If
                If CBInd = "2" Then
                    Do Until BNFSpe.EOF
                        For counter = 1 To 31 Step 1
                         Fgrid.Row = counter
                         Fgrid.Col = 0
                            If BNFSpe!BNFCode = "101" Then
                                If IsNumeric(BNFSpe!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = BNFSpe!BNFRMPerVisit
                                End If
                                    CB1 = Format(Lmt, "#,##0.00")
                            ElseIf BNFSpe!BNFCode = "102" Then
                                If IsNumeric(BNFSpe!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = BNFSpe!BNFRMPerVisit
                                End If
                                    CB2 = Format(Lmt, "#,##0.00")
                            ElseIf BNFSpe!BNFCode = "103" Then
                                If IsNumeric(BNFSpe!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = BNFSpe!BNFRMPerVisit
                                End If
                                    CB3 = Format(Lmt, "#,##0.00")
                            ElseIf BNFSpe!BNFCode = "104" Then
                                If IsNumeric(BNFSpe!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = BNFSpe!BNFRMPerVisit
                                End If
                                    CB4 = Format(Lmt, "#,##0.00")
                            ElseIf BNFSpe!BNFCode = "105" Then
                                If IsNumeric(BNFSpe!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = BNFSpe!BNFRMPerVisit
                                End If
                                    CB5 = Format(Lmt, "#,##0.00")
                            ElseIf BNFSpe!BNFCode = "106" Then
                                If IsNumeric(BNFSpe!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = BNFSpe!BNFRMPerVisit
                                End If
                                    CB6 = Format(Lmt, "#,##0.00")
                            End If
                        Exit For
                    Next
                    BNFSpe.MoveNext
                    Loop
                    
                Do Until CLNBNF.EOF
                     With CLNBNF
                        For counter = 1 To 31 Step 1
                         Fgrid.Row = counter
                         Fgrid.Col = 0
                            If Fgrid.Text = !BNFCode Then
                                If !BNFCode = "301" Then
                                    If IsNumeric(!BNFRMPerVisit) = False Then
                                        CB7 = 0
                                    Else
                                        CB7 = !BNFRMPerVisit
                                    End If
                            ElseIf !BNFCode = "302" Then
                                    If IsNumeric(!BNFRMPerVisit) = False Then
                                        CB8 = 0
                                    Else
                                        CB8 = !BNFRMPerVisit
                                    End If
                            ElseIf !BNFCode = "303" Then
                                    If IsNumeric(!BNFRMPerVisit) = False Then
                                        CB9 = 0
                                    Else
                                        CB9 = !BNFRMPerVisit
                                    End If
                            ElseIf !BNFCode = "304" Then
                                    If IsNumeric(!BNFRMPerVisit) = False Then
                                        CB10 = 0
                                    Else
                                        CB10 = !BNFRMPerVisit
                                    End If
                            ElseIf !BNFCode = "305" Then
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    CB11 = 0
                                Else
                                    CB11 = !BNFRMPerVisit
                                End If
                            ElseIf !BNFCode = "306" Then
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    CB12 = 0
                                Else
                                    CB12 = !BNFRMPerVisit
                                End If
                            ElseIf !BNFCode = "307" Then
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    CB13 = 0
                                Else
                                    CB13 = !BNFRMPerVisit
                                End If
                            ElseIf !BNFCode = "308" Then
                                'If I14 > 0 Then
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    CB14 = 0
                                Else
                                    CB14 = !BNFRMPerVisit
                                End If
                            ElseIf !BNFCode = "309" Then
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    CB15 = 0
                                Else
                                    CB15 = !BNFRMPerVisit
                                End If
                            ElseIf !BNFCode = "401" Then
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    CB16 = 0
                                Else
                                    CB16 = !BNFRMPerVisit
                                End If
                            ElseIf !BNFCode = "402" Then
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    CB17 = 0
                                Else
                                    CB17 = !BNFRMPerVisit
                                End If
                            ElseIf !BNFCode = "403" Then
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    CB18 = 0
                                Else
                                    CB18 = !BNFRMPerVisit
                                End If
                            ElseIf !BNFCode = "404" Then
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    CB19 = 0
                                Else
                                    CB19 = !BNFRMPerVisit
                                End If
                            ElseIf !BNFCode = "405" Then
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    CB20 = 0
                                Else
                                    CB20 = !BNFRMPerVisit
                                End If
                            ElseIf !BNFCode = "501" Then
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    CB21 = 0
                                Else
                                    CB21 = !BNFRMPerVisit
                                End If
                            ElseIf !BNFCode = "502" Then
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    CB22 = 0
                                Else
                                    CB22 = !BNFRMPerVisit
                                End If
                            ElseIf !BNFCode = "503" Then
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    CB23 = 0
                                Else
                                    CB23 = !BNFRMPerVisit
                                End If
                            ElseIf !BNFCode = "504" Then
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    CB24 = 0
                                Else
                                    CB24 = !BNFRMPerVisit
                                End If
                            ElseIf !BNFCode = "505" Then
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    CB25 = 0
                                Else
                                    CB25 = !BNFRMPerVisit
                                End If
                             ElseIf !BNFCode = "601" Then
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    CB26 = 0
                                Else
                                    CB26 = !BNFRMPerVisit
                                End If
                            ElseIf !BNFCode = "602" Then
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    CB27 = 0
                                Else
                                    CB27 = !BNFRMPerVisit
                                End If
                            ElseIf !BNFCode = "603" Then
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    CB28 = 0
                                Else
                                    CB28 = !BNFRMPerVisit
                                End If
                            ElseIf !BNFCode = "604" Then
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    CB29 = 0
                                Else
                                    CB29 = !BNFRMPerVisit
                                End If
                            ElseIf !BNFCode = "701" Then
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    CB30 = 0
                                Else
                                    CB30 = !BNFRMPerVisit
                                End If
                            ElseIf !BNFCode = "702" Then
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    CB31 = 0
                                Else
                                    CB31 = !BNFRMPerVisit
                                End If
                                
                        End If
                        Exit For
                    End If
                Next
                .MoveNext
             End With
            Loop
            CLNBNF.Close
            Set CLNBNF = Nothing

                ElseIf CBInd = "1" Then
                    CB1 = Format((I1), "#,##0.00")
                    CB2 = Format((I2), "#,##0.00")
                    CB3 = Format((I3), "#,##0.00")
                    CB4 = Format((I4), "#,##0.00")
                    CB5 = Format((I5), "#,##0.00")
                    CB6 = Format((I6), "#,##0.00")
                    CB7 = Format((I7), "#,##0.00")
                    CB8 = Format((I8), "#,##0.00")
                    CB9 = Format((I9), "#,##0.00")
                    CB10 = Format((I10), "#,##0.00")
                    CB11 = Format((I11), "#,##0.00")
                    CB12 = Format((I12), "#,##0.00")
                    CB13 = Format((I13), "#,##0.00")
                    CB14 = Format((I14), "#,##0.00")
                    CB15 = Format((I15), "#,##0.00")
                    CB16 = Format((I16), "#,##0.00")
                    CB17 = Format((I17), "#,##0.00")
                    CB18 = Format((I18), "#,##0.00")
                    CB19 = Format((I19), "#,##0.00")
                    CB20 = Format((I20), "#,##0.00")
                    CB21 = Format((I21), "#,##0.00")
                    CB22 = Format((I22), "#,##0.00")
                    CB23 = Format((I23), "#,##0.00")
                    CB24 = Format((I24), "#,##0.00")
                    CB25 = Format((I25), "#,##0.00")
                    CB26 = Format((I26), "#,##0.00")
                    CB27 = Format((I27), "#,##0.00")
                    CB28 = Format((I28), "#,##0.00")
                    CB29 = Format((I29), "#,##0.00")
                    CB30 = Format((I30), "#,##0.00")
                    CB31 = Format((I31), "#,##0.00")

                End If
    ElseIf INVCLNForm = 3 Then
        a = 194
        If frmInvoiceAdj.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 0
                Fgrid.Col = 3
                a = a + 1
                Fgrid.Text = Format(frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(a), "#,##0.00")
            Next
            Fgrid.Row = 1
            CB1 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 2
            CB2 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 3
            CB3 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 4
            CB4 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 5
            CB5 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 6
            CB6 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 7
            CB7 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 8
            CB8 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 8
            CB8 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 9
            CB9 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 10
            CB10 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 11
            CB11 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 12
            CB12 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 13
            CB13 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 14
            CB14 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 15
            CB15 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 16
            CB16 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 17
            CB17 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 18
            CB18 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 19
            CB19 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 20
            CB20 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 21
            CB21 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 22
            CB22 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 23
            CB23 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 24
            CB24 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 25
            CB25 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 26
            CB26 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 27
            CB27 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 28
            CB28 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 29
            CB29 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 30
            CB30 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 31
            CB31 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
        End If
    ElseIf INVCLNForm = 4 Then
        a = 194
        If frmInvoiceAdjAftBordx.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 0
                Fgrid.Col = 3
                a = a + 1
                Fgrid.Text = Format(frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(a), "#,##0.00")
            Next
            Fgrid.Row = 1
            CB1 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 2
            CB2 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 3
            CB3 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 4
            CB4 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 5
            CB5 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 6
            CB6 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 7
            CB7 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 8
            CB8 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 8
            CB8 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 9
            CB9 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 10
            CB10 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 11
            CB11 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 12
            CB12 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 13
            CB13 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 14
            CB14 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 15
            CB15 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 16
            CB16 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 17
            CB17 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 18
            CB18 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 19
            CB19 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 20
            CB20 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 21
            CB21 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 22
            CB22 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 23
            CB23 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 24
            CB24 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 25
            CB25 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 26
            CB26 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 27
            CB27 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 28
            CB28 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 29
            CB29 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 30
            CB30 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 31
            CB31 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)

        End If
        
    ElseIf INVCLNForm = 5 Then
        a = 194
        If frmInvoiceAdjAfterPost.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 0
                Fgrid.Col = 3
                a = a + 1
                Fgrid.Text = Format(frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(a), "#,##0.00")
            Next
            Fgrid.Row = 1
            CB1 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 2
            CB2 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 3
            CB3 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 4
            CB4 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 5
            CB5 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 6
            CB6 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 7
            CB7 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 8
            CB8 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 8
            CB8 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 9
            CB9 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 10
            CB10 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 11
            CB11 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 12
            CB12 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 13
            CB13 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 14
            CB14 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 15
            CB15 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 16
            CB16 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 17
            CB17 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 18
            CB18 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 19
            CB19 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 20
            CB20 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 21
            CB21 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 22
            CB22 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 23
            CB23 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 24
            CB24 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 25
            CB25 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 26
            CB26 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 27
            CB27 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 28
            CB28 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 29
            CB29 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 30
            CB30 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
            Fgrid.Row = 31
            CB31 = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)

        End If
        
    End If
End Sub

Private Sub CalculateReim()
    Dim ListCount As Integer
    Dim ForLoop As Integer
    Dim curColReim As String
    Dim curColCashless As String
    
    RGPFeeBMed = 0
    RGPFeeLTMed = 0
    RGPFeeMedInj = 0
    RSPFeeBMed = 0
    RSPFeeLTMed = 0
    RSPFeeMedInj = 0
    RDental = 0
    RVision = 0
    RMaternity = 0
    RPRVCheckup = 0
    RXRay = 0
    RUltraSound = 0
    RECG = 0
    RDressing = 0
    RNebulizer = 0
    RID = 0
    RTSSTO = 0
    RMinorSurgery = 0
    RPOthers = 0
    RBloodDrawing = 0
    RBloodGlucose = 0
    RUrine = 0
    RPapSmear = 0
    RLOthers = 0
    RDTPPolio = 0
    RDTPPolioHIB = 0
    RHepatitisB = 0
    RMMR = 0
    RIOthers = 0
    CTyphoidNoDed = 0
    COthersNoDed = 0
    
    If frmInvoice.lstInvoice.ListItems.Count > 0 Then
        If frmInvoice.lstInvoice.SelectedItem.SubItems(73) = "M" Then
            RGPFeeBMed = frmInvoice.lstInvoice.SelectedItem.SubItems(15)
            RGPFeeLTMed = frmInvoice.lstInvoice.SelectedItem.SubItems(16)
            RGPFeeMedInj = frmInvoice.lstInvoice.SelectedItem.SubItems(17)
            RSPFeeBMed = frmInvoice.lstInvoice.SelectedItem.SubItems(18)
            RSPFeeLTMed = frmInvoice.lstInvoice.SelectedItem.SubItems(19)
            RSPFeeMedInj = frmInvoice.lstInvoice.SelectedItem.SubItems(20)
            RXRay = frmInvoice.lstInvoice.SelectedItem.SubItems(24)
            RUltraSound = frmInvoice.lstInvoice.SelectedItem.SubItems(25)
            RECG = frmInvoice.lstInvoice.SelectedItem.SubItems(26)
            RDressing = frmInvoice.lstInvoice.SelectedItem.SubItems(27)
            RNebulizer = frmInvoice.lstInvoice.SelectedItem.SubItems(28)
            RID = frmInvoice.lstInvoice.SelectedItem.SubItems(29)
            RTSSTO = frmInvoice.lstInvoice.SelectedItem.SubItems(30)
            RMinorSurgery = frmInvoice.lstInvoice.SelectedItem.SubItems(31)
            RPOthers = frmInvoice.lstInvoice.SelectedItem.SubItems(32)
            RBloodDrawing = frmInvoice.lstInvoice.SelectedItem.SubItems(33)
            RBloodGlucose = frmInvoice.lstInvoice.SelectedItem.SubItems(34)
            RUrine = frmInvoice.lstInvoice.SelectedItem.SubItems(35)
            RPapSmear = frmInvoice.lstInvoice.SelectedItem.SubItems(36)
            RLOthers = frmInvoice.lstInvoice.SelectedItem.SubItems(37)
            RDTPPolio = frmInvoice.lstInvoice.SelectedItem.SubItems(38)
            RDTPPolioHIB = frmInvoice.lstInvoice.SelectedItem.SubItems(39)
            RHepatitisB = frmInvoice.lstInvoice.SelectedItem.SubItems(40)
            RMMR = frmInvoice.lstInvoice.SelectedItem.SubItems(41)
            RIOthers = frmInvoice.lstInvoice.SelectedItem.SubItems(42)
            RDental = frmInvoice.lstInvoice.SelectedItem.SubItems(21)
            RMaternity = frmInvoice.lstInvoice.SelectedItem.SubItems(22)
            RPRVCheckup = frmInvoice.lstInvoice.SelectedItem.SubItems(23)
            RVision = frmInvoice.lstInvoice.SelectedItem.SubItems(102)
            CTyphoidNoDed = frmInvoice.lstInvoice.SelectedItem.SubItems(474)
            COthersNoDed = frmInvoice.lstInvoice.SelectedItem.SubItems(475)

        End If
    End If
End Sub

Private Sub CalculateCashless()
    Dim ListCount As Integer
    Dim ForLoop As Integer
    Dim curColReim As String
    Dim curColCashless As String
   
    CGPFeeBMed = 0
    CGPFeeLTMed = 0
    CGPFeeMedInj = 0
    CSPFeeBMed = 0
    CSPFeeLTMed = 0
    CSPFeeMedInj = 0
    CDental = 0
    CVision = 0
    CMaternity = 0
    CPRVCheckup = 0
    CXRay = 0
    CUltraSound = 0
    CECG = 0
    CDressing = 0
    CNebulizer = 0
    CID = 0
    CTSSTO = 0
    CMinorSurgery = 0
    CPOthers = 0
    CBloodDrawing = 0
    CBloodGlucose = 0
    CUrine = 0
    CPapSmear = 0
    CLOthers = 0
    CDTPPolio = 0
    CDTPPolioHIB = 0
    CHepatitisB = 0
    CMMR = 0
    CIOthers = 0
    CTyphoidNoDed = 0
    COthersNoDed = 0

    If frmInvoice.lstInvoice.ListItems.Count > 0 Then
        If frmInvoice.lstInvoice.SelectedItem.SubItems(73) = "C" Then
            CGPFeeBMed = frmInvoice.lstInvoice.SelectedItem.SubItems(15)
            CGPFeeLTMed = frmInvoice.lstInvoice.SelectedItem.SubItems(16)
            CGPFeeMedInj = frmInvoice.lstInvoice.SelectedItem.SubItems(17)
            CSPFeeBMed = frmInvoice.lstInvoice.SelectedItem.SubItems(18)
            CSPFeeLTMed = frmInvoice.lstInvoice.SelectedItem.SubItems(19)
            CSPFeeMedInj = frmInvoice.lstInvoice.SelectedItem.SubItems(20)
            CXRay = frmInvoice.lstInvoice.SelectedItem.SubItems(24)
            CUltraSound = frmInvoice.lstInvoice.SelectedItem.SubItems(25)
            CECG = frmInvoice.lstInvoice.SelectedItem.SubItems(26)
            CDressing = frmInvoice.lstInvoice.SelectedItem.SubItems(27)
            CNebulizer = frmInvoice.lstInvoice.SelectedItem.SubItems(28)
            CID = frmInvoice.lstInvoice.SelectedItem.SubItems(29)
            CTSSTO = frmInvoice.lstInvoice.SelectedItem.SubItems(30)
            CMinorSurgery = frmInvoice.lstInvoice.SelectedItem.SubItems(31)
            CPOthers = frmInvoice.lstInvoice.SelectedItem.SubItems(32)
            CBloodDrawing = frmInvoice.lstInvoice.SelectedItem.SubItems(33)
            CBloodGlucose = frmInvoice.lstInvoice.SelectedItem.SubItems(34)
            CUrine = frmInvoice.lstInvoice.SelectedItem.SubItems(35)
            CPapSmear = frmInvoice.lstInvoice.SelectedItem.SubItems(35)
            CLOthers = frmInvoice.lstInvoice.SelectedItem.SubItems(37)
            CDTPPolio = frmInvoice.lstInvoice.SelectedItem.SubItems(38)
            CDTPPolioHIB = frmInvoice.lstInvoice.SelectedItem.SubItems(39)
            CHepatitisB = frmInvoice.lstInvoice.SelectedItem.SubItems(40)
            CMMR = frmInvoice.lstInvoice.SelectedItem.SubItems(41)
            CIOthers = frmInvoice.lstInvoice.SelectedItem.SubItems(42)
            CDental = frmInvoice.lstInvoice.SelectedItem.SubItems(21)
            CMaternity = frmInvoice.lstInvoice.SelectedItem.SubItems(22)
            CPRVCheckup = frmInvoice.lstInvoice.SelectedItem.SubItems(23)
            CVision = frmInvoice.lstInvoice.SelectedItem.SubItems(102)
            CTyphoidNoDed = frmInvoice.lstInvoice.SelectedItem.SubItems(474)
            COthersNoDed = frmInvoice.lstInvoice.SelectedItem.SubItems(475)

        End If
    End If
End Sub

Private Sub getDatatoWS()
Dim StrSql As String
Dim CLM As New ADODB.Recordset
Dim ForLoop, ListCount As Integer

Dim A1, A2, A3, A4, A5 As Double
Dim A6, A7, A8, A9, A10 As Double
Dim A11, A12, A13, A14, A15 As Double
Dim A16, A17, A18, A19, A20, A21 As Double
Dim A22, A23, A24, A25, A26, A27, A28, A29, A30, A31 As Double

' Payto Hospital
Dim C1, C2, C3, C4, C5 As Double
Dim C6, C7, C8, C9, C10 As Double
Dim C11, C12, C13, C14, C15 As Double
Dim C16, C17, C18, C19, C20, C21   As Double
Dim C22, C23, C24, C25, C26, C27, C28, C29, C30, C31 As Double

Dim Tax As Double

' LIMIT
Dim L1, L2, L3, L4, L5 As String
Dim L6, L7, L8, L9, L10 As String
Dim L11, L12, L13, L14, L15 As String
Dim L16, L17, L18, L19, L20, L21 As String
Dim L22, L23, L24, L25, L26, L27, L28, L29, L30, L31 As String
Dim strsqlSameClm As String
Dim SameClm As New ADODB.Recordset
    
    
    ' Hospital
    A1 = 0
    'ICU1 = 0
    A2 = 0
    A3 = 0
    A4 = 0
    A5 = 0
    A6 = 0
    A7 = 0
    A8 = 0
    A9 = 0
    A10 = 0
    A11 = 0
    A12 = 0
    A13 = 0
    A14 = 0
    A15 = 0
    A16 = 0
    A17 = 0
    A18 = 0
    A19 = 0
    A20 = 0
    A21 = 0
    A22 = 0
    A23 = 0
    A24 = 0
    A25 = 0
    A26 = 0
    A27 = 0
    A28 = 0
    A29 = 0
    A30 = 0
    A31 = 0
    ' Clinic
    C1 = 0
    C2 = 0
    C3 = 0
    C4 = 0
    C5 = 0
    C6 = 0
    C7 = 0
    C8 = 0
    C9 = 0
    C10 = 0
    C11 = 0
    C12 = 0
    C13 = 0
    C14 = 0
    C15 = 0
    C16 = 0
    C17 = 0
    C18 = 0
    C19 = 0
    C20 = 0
    C21 = 0
    C22 = 0
    C23 = 0
    C24 = 0
    C25 = 0
    C26 = 0
    C27 = 0
    C28 = 0
    C29 = 0
    C30 = 0
    C31 = 0
    
        ' SPFee + Med
        A1 = CDbl(IIf(IsNumeric(I1), I1, 0))
        ' GPFee + Med + Inj
        A2 = CDbl(IIf(IsNumeric(I2), I2, 0))
        ' GPFee LTMed
        A3 = CDbl(IIf(IsNumeric(I3), I3, 0))
        ' SPFee + Med
        A4 = CDbl(IIf(IsNumeric(I4), I4, 0))
        ' SPFee + Med + Inj
        A5 = CDbl(IIf(IsNumeric(I5), I5, 0))
        ' SPFee + LTMed
        A6 = CDbl(IIf(IsNumeric(I6), I6, 0))
        ' XRay
        A7 = CDbl(IIf(IsNumeric(I7), I7, 0))
        ' UltraSound
        A8 = CDbl(IIf(IsNumeric(I8), I8, 0))
        ' ECG
        A9 = CDbl(IIf(IsNumeric(I9), I9, 0))
        ' Dressing
        A10 = CDbl(IIf(IsNumeric(I10), I10, 0))
        ' Nebulizer
        A11 = CDbl(IIf(IsNumeric(I11), I11, 0))
        ' I&D/T&S
        A12 = CDbl(IIf(IsNumeric(I12), I12, 0))
        ' STO
        A13 = CDbl(IIf(IsNumeric(I13), I13, 0))
        ' Minor Surgery
        A14 = CDbl(IIf(IsNumeric(I14), I14, 0))
        ' Proc Other
        A15 = CDbl(IIf(IsNumeric(I15), I15, 0))
        ' Blood Drawing
        A16 = CDbl(IIf(IsNumeric(I16), I16, 0))
        ' Blood Glucose
        A17 = CDbl(IIf(IsNumeric(I17), I17, 0))
        ' Urine
        A18 = CDbl(IIf(IsNumeric(I18), I18, 0))
        ' Pap Smear
        A19 = CDbl(IIf(IsNumeric(I19), I19, 0))
        ' Lab Others
        A20 = CDbl(IIf(IsNumeric(I20), I20, 0))
        ' DTP & Polio
        A21 = CDbl(IIf(IsNumeric(I21), I21, 0))
        ' DTP & Polio &HIB
        A22 = CDbl(IIf(IsNumeric(I22), I22, 0))
        ' Hepatitis B
        A23 = CDbl(IIf(IsNumeric(I23), I23, 0))
        ' MMR
        A24 = CDbl(IIf(IsNumeric(I24), I24, 0))
        ' Imm Others
        A25 = CDbl(IIf(IsNumeric(I25), I25, 0))
        ' Dental
        A26 = CDbl(IIf(IsNumeric(I26), I26, 0))
        ' Maternity
        A27 = CDbl(IIf(IsNumeric(I27), I27, 0))
        ' PRV Checkup
        A28 = CDbl(IIf(IsNumeric(I28), I28, 0))
         'Vision
        A29 = CDbl(IIf(IsNumeric(I29), I29, 0))
     
        'Typhoid no bal ded
        A30 = CDbl(IIf(IsNumeric(I30), I30, 0))
         'Others no bal ded
        A31 = CDbl(IIf(IsNumeric(I31), I31, 0))
     
     ' For paying to Clinic
        
        If INVCLNForm = 1 Or INVCLNForm = 3 Or INVCLNForm = 4 Or INVCLNForm = 5 Then
            ' SPFee + Med
            C1 = CDbl(IIf(IsNumeric(I1), I1, 0))
            ' GPFee + Med + Inj
            C2 = CDbl(IIf(IsNumeric(I2), I2, 0))
            ' GPFee LTMed
            C3 = CDbl(IIf(IsNumeric(I3), I3, 0))
            ' SPFee + Med
            C4 = CDbl(IIf(IsNumeric(I4), I4, 0))
            ' SPFee + Med + Inj
            C5 = CDbl(IIf(IsNumeric(I5), I5, 0))
            ' SPFee + LTMed
            C6 = CDbl(IIf(IsNumeric(I6), I6, 0))
            ' XRay
            C7 = CDbl(IIf(IsNumeric(I7), I7, 0))
            ' UltraSound
            C8 = CDbl(IIf(IsNumeric(I8), I8, 0))
            ' ECG
            C9 = CDbl(IIf(IsNumeric(I9), I9, 0))
            ' Dressing
            C10 = CDbl(IIf(IsNumeric(I10), I10, 0))
            ' Nebulizer
            C11 = CDbl(IIf(IsNumeric(I11), I11, 0))
            ' I&D/T&S
            C12 = CDbl(IIf(IsNumeric(I12), I12, 0))
            ' STO
            C13 = CDbl(IIf(IsNumeric(I13), I13, 0))
            ' Minor Surgery
            C14 = CDbl(IIf(IsNumeric(I14), I14, 0))
            ' Proc Other
            C15 = CDbl(IIf(IsNumeric(I15), I15, 0))
            ' Blood Drawing
            C16 = CDbl(IIf(IsNumeric(I16), I16, 0))
            ' Blood Glucose
            C17 = CDbl(IIf(IsNumeric(I17), I17, 0))
            ' Urine
            C18 = CDbl(IIf(IsNumeric(I18), I18, 0))
            ' Pap Smear
            C19 = CDbl(IIf(IsNumeric(I19), I19, 0))
            ' Lab Others
            C20 = CDbl(IIf(IsNumeric(I20), I20, 0))
            ' DTP & Polio
            C21 = CDbl(IIf(IsNumeric(I21), I21, 0))
            ' DTP & Polio &HIB
            C22 = CDbl(IIf(IsNumeric(I22), I22, 0))
            ' Hepatitis B
            C23 = CDbl(IIf(IsNumeric(I23), I23, 0))
            ' MMR
            C24 = CDbl(IIf(IsNumeric(I24), I24, 0))
            ' Imm Others
            C25 = CDbl(IIf(IsNumeric(I25), I25, 0))
            ' Dental
            C26 = CDbl(IIf(IsNumeric(I26), I26, 0))
            ' Maternity
            C27 = CDbl(IIf(IsNumeric(I27), I27, 0))
            ' PRV Checkup
            C28 = CDbl(IIf(IsNumeric(I28), I28, 0))
            ' Vision
            C29 = CDbl(IIf(IsNumeric(I29), I29, 0))
            ' Typhoid no bal ded
            C30 = CDbl(IIf(IsNumeric(I30), I30, 0))
            ' Others no bal ded
            C31 = CDbl(IIf(IsNumeric(I31), I31, 0))

        End If
    ' MBM Limit col. for Days/Months
    Fgrid.Col = 6
                            
    Fgrid.Row = 1
    L1 = I1
    Fgrid.Row = 2
    L2 = I2
    Fgrid.Row = 3
    L3 = I3
    Fgrid.Row = 4
    L4 = I4
    Fgrid.Row = 5
    L5 = I5
    Fgrid.Row = 6
    L6 = I6
    Fgrid.Row = 7
    L7 = I7
    Fgrid.Row = 8
    L8 = I8
    Fgrid.Row = 9
    L9 = I9
    Fgrid.Row = 10
    L10 = I10
    Fgrid.Row = 11
    L11 = I11
    Fgrid.Row = 12
    L12 = I12
    Fgrid.Row = 13
    L13 = I13
    Fgrid.Row = 14
    L14 = I14
    Fgrid.Row = 15
    L15 = I15
    Fgrid.Row = 16
    L16 = I16
    Fgrid.Row = 17
    L17 = I17
    Fgrid.Row = 18
    L18 = I18
    Fgrid.Row = 19
    L19 = I19
    Fgrid.Row = 20
    L20 = I20
    Fgrid.Row = 21
    L21 = I21
    Fgrid.Row = 22
    L22 = I22
    Fgrid.Row = 23
    L23 = I23
    Fgrid.Row = 24
    L24 = I24
    Fgrid.Row = 25
    L25 = I25
    Fgrid.Row = 26
    L26 = I26
    Fgrid.Row = 27
    L27 = I27
    Fgrid.Row = 28
    L28 = I28
    Fgrid.Row = 29
    L29 = I29
    Fgrid.Row = 30
    L30 = I30
    Fgrid.Row = 31
    L31 = I31

    
    'Actual Col.
    Fgrid.Col = 2
    Fgrid.Row = 1
    Fgrid.Text = IIf(A1 = 0, "-", Format(A1, "#,#.00"))
    Fgrid.Row = 2
    Fgrid.Text = IIf(A2 = 0, "-", Format(A2, "#,#.00"))
    Fgrid.Row = 3
    Fgrid.Text = IIf(A3 = 0, "-", Format(A3, "#,#.00"))
    Fgrid.Row = 4
    Fgrid.Text = IIf(A4 = 0, "-", Format(A4, " #,#.00"))
    Fgrid.Row = 5
    Fgrid.Text = IIf(A5 = 0, "-", Format(A5, "#,#.00"))
    Fgrid.Row = 6
    Fgrid.Text = IIf(A6 = 0, "-", Format(A6, "#,#.00"))
    Fgrid.Row = 7
    Fgrid.Text = IIf(A7 = 0, "-", Format(A7, "#,#.00"))
    Fgrid.Row = 8
    Fgrid.Text = IIf(A8 = 0, "-", Format(A8, "#,#.00"))
    Fgrid.Row = 9
    Fgrid.Text = IIf(A9 = 0, "-", Format(A9, "#,#.00"))
    Fgrid.Row = 10
    Fgrid.Text = IIf(A10 = 0, "-", Format(A10, "#,#.00"))
    Fgrid.Row = 11
    Fgrid.Text = IIf(A11 = 0, "-", Format(A11, "#,#.00"))
    Fgrid.Row = 12
    Fgrid.Text = IIf(A12 = 0, "-", Format(A12, "#,#.00"))
    Fgrid.Row = 13
    Fgrid.Text = IIf(A13 = 0, "-", Format(A13, "#,#.00"))
    Fgrid.Row = 14
    Fgrid.Text = IIf(A14 = 0, "-", Format(A14, "#,#.00"))
    Fgrid.Row = 15
    Fgrid.Text = IIf(A15 = 0, "-", Format(A15, "#,#.00"))
    Fgrid.Row = 16
    Fgrid.Text = IIf(A16 = 0, "-", Format(A16, "#,#.00"))
    Fgrid.Row = 17
    Fgrid.Text = IIf(A17 = 0, "-", Format(A17, "#,#.00"))
    Fgrid.Row = 18
    Fgrid.Text = IIf(A18 = 0, "-", Format(A18, "#,#.00"))
    Fgrid.Row = 19
    Fgrid.Text = IIf(A19 = 0, "-", Format(A19, "#,#.00"))
    Fgrid.Row = 20
    Fgrid.Text = IIf(A20 = 0, "-", Format(A20, "#,#.00"))
    Fgrid.Row = 21
    Fgrid.Text = IIf(A21 = 0, "-", Format(A21, "#,#.00"))
    Fgrid.Row = 22
    Fgrid.Text = IIf(A22 = 0, "-", Format(A22, "#,#.00"))
    Fgrid.Row = 23
    Fgrid.Text = IIf(A23 = 0, "-", Format(A23, "#,#.00"))
    Fgrid.Row = 24
    Fgrid.Text = IIf(A24 = 0, "-", Format(A24, "#,#.00"))
    Fgrid.Row = 25
    Fgrid.Text = IIf(A25 = 0, "-", Format(A25, "#,#.00"))
    Fgrid.Row = 26
    Fgrid.Text = IIf(A26 = 0, "-", Format(A26, "#,#.00"))
    Fgrid.Row = 27
    Fgrid.Text = IIf(A27 = 0, "-", Format(A27, "#,#.00"))
    Fgrid.Row = 28
    Fgrid.Text = IIf(A28 = 0, "-", Format(A28, "#,#.00"))
    Fgrid.Row = 29
    Fgrid.Text = IIf(A29 = 0, "-", Format(A29, "#,#.00"))
    Fgrid.Row = 30
    Fgrid.Text = IIf(A30 = 0, "-", Format(A30, "#,#.00"))
    Fgrid.Row = 31
    Fgrid.Text = IIf(A31 = 0, "-", Format(A31, "#,#.00"))
    
    'MBM approved Col'
    Fgrid.Col = 7
    Fgrid.Row = 1
    Fgrid.Text = GetMin((A1), (L1))
    App1 = Fgrid.Text
    Fgrid.Row = 2
    Fgrid.Text = GetMin((A2), (L2))
    App2 = Fgrid.Text
    Fgrid.Row = 3
    Fgrid.Text = GetMin((A3), (L3))
    App3 = Fgrid.Text
    Fgrid.Row = 4
    Fgrid.Text = GetMin((A4), (L4))
    App4 = Fgrid.Text
    Fgrid.Row = 5
    Fgrid.Text = GetMin((A5), (L5))
    App5 = Fgrid.Text
    Fgrid.Row = 6
    Fgrid.Text = GetMin((A6), (L6))
    App6 = Fgrid.Text
    Fgrid.Row = 7
    Fgrid.Text = GetMin((A7), (L7))
    App7 = Fgrid.Text
    Fgrid.Row = 8
    Fgrid.Text = GetMin((A8), (L8))
    App8 = Fgrid.Text
    Fgrid.Row = 9
    Fgrid.Text = GetMin((A9), (L9))
    App9 = Fgrid.Text
    Fgrid.Row = 10
    Fgrid.Text = GetMin((A10), (L10))
    App10 = Fgrid.Text
    Fgrid.Row = 11
    Fgrid.Text = GetMin((A11), (L11))
    App11 = Fgrid.Text
    Fgrid.Row = 12
    Fgrid.Text = GetMin((A12), (L12))
    App12 = Fgrid.Text
    Fgrid.Row = 13
    Fgrid.Text = GetMin((A13), (L13))
    App13 = Fgrid.Text
    Fgrid.Row = 14
    Fgrid.Text = GetMin((A14), (L14))
    App14 = Fgrid.Text
    Fgrid.Row = 15
    Fgrid.Text = GetMin((A15), (L15))
    App15 = Fgrid.Text
    Fgrid.Row = 16
    Fgrid.Text = GetMin((A16), (L16))
    App16 = Fgrid.Text
    Fgrid.Row = 17
    Fgrid.Text = GetMin((A17), (L17))
    App17 = Fgrid.Text
    Fgrid.Row = 18
    Fgrid.Text = GetMin((A18), (L18))
    App18 = Fgrid.Text
    Fgrid.Row = 19
    Fgrid.Text = GetMin((A19), (L19))
    App19 = Fgrid.Text
    Fgrid.Row = 20
    Fgrid.Text = GetMin((A20), (L20))
    App20 = Fgrid.Text
    Fgrid.Row = 21
    Fgrid.Text = GetMin((A21), (L21))
    App21 = Fgrid.Text
    Fgrid.Row = 22
    Fgrid.Text = GetMin((A22), (L22))
    App22 = Fgrid.Text
    Fgrid.Row = 23
    Fgrid.Text = GetMin((A23), (L23))
    App23 = Fgrid.Text
    Fgrid.Row = 24
    Fgrid.Text = GetMin((A24), (L24))
    App24 = Fgrid.Text
    Fgrid.Row = 25
    Fgrid.Text = GetMin((A25), (L25))
    APP25 = Fgrid.Text
    Fgrid.Row = 26
    Fgrid.Text = GetMin((A26), (L26))
    APP26 = Fgrid.Text
    Fgrid.Row = 27
    Fgrid.Text = GetMin((A27), (L27))
    App27 = Fgrid.Text
    Fgrid.Row = 28
    Fgrid.Text = GetMin((A28), (L28))
    APP28 = Fgrid.Text
    Fgrid.Row = 29
    Fgrid.Text = GetMin((A29), (L29))
    APP29 = Fgrid.Text
    Fgrid.Row = 30
    Fgrid.Text = GetMin((A30), (L30))
    APP30 = Fgrid.Text
    Fgrid.Row = 31
    Fgrid.Text = GetMin((A31), (L31))
    APP31 = Fgrid.Text
    
    ' approved Due to From MEM
    Fgrid.Col = 8
    Fgrid.Row = 1
        Fgrid.Text = DueToFromMBM((A1), (App1))
    Fgrid.Row = 2
        Fgrid.Text = DueToFromMBM((A2), (App2))
    
    Fgrid.Row = 3
        Fgrid.Text = DueToFromMBM((A3), (App3))
    
    Fgrid.Row = 4
        Fgrid.Text = DueToFromMBM((A4), (App4))
    
    Fgrid.Row = 5
        Fgrid.Text = DueToFromMBM((A5), (App5))
    
    Fgrid.Row = 6
        Fgrid.Text = DueToFromMBM((A6), (App6))
    
    Fgrid.Row = 7
        Fgrid.Text = DueToFromMBM((A7), (App7))
    
    Fgrid.Row = 8
        Fgrid.Text = DueToFromMBM((A8), (App8))
    
    Fgrid.Row = 9
        Fgrid.Text = DueToFromMBM((A9), (App9))
    
    Fgrid.Row = 10
        Fgrid.Text = DueToFromMBM((A10), (App10))
    
    Fgrid.Row = 11
        Fgrid.Text = DueToFromMBM((A11), (App11))
    
    Fgrid.Row = 12
        Fgrid.Text = DueToFromMBM((A12), (App12))
    
    Fgrid.Row = 13
        Fgrid.Text = DueToFromMBM((A13), (App13))
    
    Fgrid.Row = 14
        Fgrid.Text = DueToFromMBM((A14), (App14))
    
    Fgrid.Row = 15
        Fgrid.Text = DueToFromMBM((A15), (App15))
    
    Fgrid.Row = 16
        Fgrid.Text = DueToFromMBM((A16), (App16))
    
    Fgrid.Row = 17
        Fgrid.Text = DueToFromMBM((A17), (App17))
    
    Fgrid.Row = 18
        Fgrid.Text = DueToFromMBM((A18), (App18))
    
    Fgrid.Row = 19
        Fgrid.Text = DueToFromMBM((A19), (App19))
    
    Fgrid.Row = 20
        Fgrid.Text = DueToFromMBM((A20), (App20))
    
    Fgrid.Row = 21
        Fgrid.Text = DueToFromMBM((A21), (App21))
    
    Fgrid.Row = 22
        Fgrid.Text = DueToFromMBM((A22), (App22))
    
    Fgrid.Row = 23
        Fgrid.Text = DueToFromMBM((A23), (App23))
    
    Fgrid.Row = 24
        Fgrid.Text = DueToFromMBM((A24), (App24))
    
    Fgrid.Row = 25
        Fgrid.Text = DueToFromMBM((A25), (APP25))
    
    Fgrid.Row = 26
        Fgrid.Text = DueToFromMBM((A26), (APP26))
    
    Fgrid.Row = 27
        Fgrid.Text = DueToFromMBM((A27), (App27))
        
    Fgrid.Row = 28
        Fgrid.Text = DueToFromMBM((A28), (APP28))
    
    Fgrid.Row = 29
        Fgrid.Text = DueToFromMBM((A29), (APP29))
            
    Fgrid.Row = 30
        Fgrid.Text = DueToFromMBM((A30), (APP30))
    
    Fgrid.Row = 31
        Fgrid.Text = DueToFromMBM((A31), (APP31))
            
            ' CLN Limit col. for Days/Months
    Fgrid.Col = 3
    If CBInd = "1" Then
        Fgrid.Row = 1
            Fgrid.Text = C1
        Fgrid.Row = 2
            Fgrid.Text = C2
        Fgrid.Row = 3
            Fgrid.Text = C3
        Fgrid.Row = 4
            Fgrid.Text = C4
        Fgrid.Row = 5
            Fgrid.Text = C5
        Fgrid.Row = 6
            Fgrid.Text = C6
        Fgrid.Row = 7
            Fgrid.Text = C7
        Fgrid.Row = 8
            Fgrid.Text = C8
        Fgrid.Row = 9
            Fgrid.Text = C9
        Fgrid.Row = 10
            Fgrid.Text = C10
        Fgrid.Row = 11
            Fgrid.Text = C11
        Fgrid.Row = 12
            Fgrid.Text = C12
        Fgrid.Row = 13
            Fgrid.Text = C13
        Fgrid.Row = 14
            Fgrid.Text = C14
        Fgrid.Row = 15
            Fgrid.Text = C15
        Fgrid.Row = 16
            Fgrid.Text = C16
        Fgrid.Row = 17
            Fgrid.Text = C17
        Fgrid.Row = 18
            Fgrid.Text = C18
        Fgrid.Row = 19
            Fgrid.Text = C19
        Fgrid.Row = 20
            Fgrid.Text = C20
        Fgrid.Row = 21
            Fgrid.Text = C21
        Fgrid.Row = 22
            Fgrid.Text = C22
        Fgrid.Row = 23
            Fgrid.Text = C23
        Fgrid.Row = 24
            Fgrid.Text = C24
        Fgrid.Row = 25
            Fgrid.Text = C25
        Fgrid.Row = 26
            Fgrid.Text = C26
        Fgrid.Row = 27
            Fgrid.Text = C27
        Fgrid.Row = 28
            Fgrid.Text = C28
        Fgrid.Row = 29
            Fgrid.Text = C29
    
        Fgrid.Row = 30
            Fgrid.Text = C30
        Fgrid.Row = 31
            Fgrid.Text = C31
    
    ElseIf CBInd = "1" Then
        Fgrid.Row = 1
            Fgrid.Text = I1
        Fgrid.Row = 2
            Fgrid.Text = I2
        Fgrid.Row = 3
            Fgrid.Text = I3
        Fgrid.Row = 4
            Fgrid.Text = I4
        Fgrid.Row = 5
            Fgrid.Text = I5
        Fgrid.Row = 6
            Fgrid.Text = I6
        Fgrid.Row = 7
            Fgrid.Text = I7
        Fgrid.Row = 8
            Fgrid.Text = I8
        Fgrid.Row = 9
            Fgrid.Text = I9
        Fgrid.Row = 10
            Fgrid.Text = I10
        Fgrid.Row = 11
            Fgrid.Text = I11
        Fgrid.Row = 12
            Fgrid.Text = I12
        Fgrid.Row = 13
            Fgrid.Text = I13
        Fgrid.Row = 14
            Fgrid.Text = I14
        Fgrid.Row = 15
            Fgrid.Text = I15
        Fgrid.Row = 16
            Fgrid.Text = I16
        Fgrid.Row = 17
            Fgrid.Text = I17
        Fgrid.Row = 18
            Fgrid.Text = I18
        Fgrid.Row = 19
            Fgrid.Text = I19
        Fgrid.Row = 20
            Fgrid.Text = I20
        Fgrid.Row = 21
            Fgrid.Text = I21
        Fgrid.Row = 22
            Fgrid.Text = I22
        Fgrid.Row = 23
            Fgrid.Text = I23
        Fgrid.Row = 24
            Fgrid.Text = I24
        Fgrid.Row = 25
            Fgrid.Text = I25
        Fgrid.Row = 26
            Fgrid.Text = I26
        Fgrid.Row = 27
            Fgrid.Text = I27
        Fgrid.Row = 28
            Fgrid.Text = I28
        Fgrid.Row = 29
            Fgrid.Text = I29
        Fgrid.Row = 30
            Fgrid.Text = I30
        Fgrid.Row = 31
            Fgrid.Text = I31
    End If
    ' Clinic Col
    Fgrid.Col = 4
    If CBInd = "1" Then
        Fgrid.Row = 1
        Fgrid.Text = Format(C1, "#,#.00")
        Fgrid.Row = 2
        Fgrid.Text = Format(C2, "#,#.00")
        Fgrid.Row = 3
        Fgrid.Text = Format(C3, "#,#.00")
        Fgrid.Row = 4
        Fgrid.Text = Format(C4, " #,#.00")
        Fgrid.Row = 5
        Fgrid.Text = Format(C5, "#,#.00")
        Fgrid.Row = 6
        Fgrid.Text = Format(C6, "#,#.00")
        Fgrid.Row = 7
        Fgrid.Text = Format(C7, "#,#.00")
        Fgrid.Row = 8
        Fgrid.Text = Format(C8, "#,#.00")
        Fgrid.Row = 9
        Fgrid.Text = Format(C9, "#,#.00")
        Fgrid.Row = 10
        Fgrid.Text = Format(C10, "#,#.00")
        Fgrid.Row = 11
        Fgrid.Text = Format(C11, "#,#.00")
        Fgrid.Row = 12
        Fgrid.Text = Format(C12, "#,#.00")
        Fgrid.Row = 13
        Fgrid.Text = Format(C13, "#,#.00")
        Fgrid.Row = 14
        Fgrid.Text = Format(C14, "#,#.00")
        Fgrid.Row = 15
        Fgrid.Text = Format(C15, "#,#.00")
        Fgrid.Row = 16
        Fgrid.Text = Format(C16, "#,#.00")
        Fgrid.Row = 17
        Fgrid.Text = Format(C17, "#,#.00")
        Fgrid.Row = 18
        Fgrid.Text = Format(C18, "#,#.00")
        Fgrid.Row = 19
        Fgrid.Text = Format(C19, "#,#.00")
        Fgrid.Row = 20
        Fgrid.Text = Format(C20, "#,#.00")
        Fgrid.Row = 21
        Fgrid.Text = Format(C21, "#,#.00")
        Fgrid.Row = 22
        Fgrid.Text = Format(C22, "#,#.00")
        Fgrid.Row = 23
        Fgrid.Text = Format(C23, "#,#.00")
        Fgrid.Row = 24
        Fgrid.Text = Format(C24, "#,#.00")
        Fgrid.Row = 25
        Fgrid.Text = Format(C25, "#,#.00")
        Fgrid.Row = 26
        Fgrid.Text = Format(C26, "#,#.00")
        Fgrid.Row = 27
        Fgrid.Text = Format(C27, "#,#.00")
        Fgrid.Row = 28
        Fgrid.Text = Format(C28, "#,#.00")
        Fgrid.Row = 29
        Fgrid.Text = Format(C29, "#,#.00")
        Fgrid.Row = 30
        Fgrid.Text = Format(C30, "#,#.00")
        Fgrid.Row = 31
        Fgrid.Text = Format(C31, "#,#.00")
        
    ElseIf CBInd = "2" Or CBInd = "3" Then
        'Fgrid.Col = 6
        Fgrid.Row = 1
        Fgrid.Text = GetMin((C1), (I1)) 'IIf(C1 = 0, "-", Format(CB1, "#,#.00"))
        Fgrid.Row = 2
        Fgrid.Text = GetMin((C2), (I2)) 'IIf(C2 = 0, "-", Format(CB2, "#,#.00"))
        Fgrid.Row = 3
        Fgrid.Text = GetMin((C3), (I3)) 'IIf(C3 = 0, "-", Format(CB3, "#,#.00"))
        Fgrid.Row = 4
        Fgrid.Text = GetMin((C4), (I4)) 'IIf(C4 = 0, "-", Format(CB4, " #,#.00"))
        Fgrid.Row = 5
        Fgrid.Text = GetMin((C5), (I5)) 'IIf(C5 = 0, "-", Format(CB5, "#,#.00"))
        Fgrid.Row = 6
        Fgrid.Text = GetMin((C6), (I6)) 'IIf(C6 = 0, "-", Format(CB6, "#,#.00"))
        Fgrid.Row = 7
        Fgrid.Text = GetMin((C7), (I7)) 'IIf(C7 = 0, "-", Format(CB7, "#,#.00"))
        Fgrid.Row = 8
        Fgrid.Text = GetMin((C8), (I8)) 'IIf(C8 = 0, "-", Format(CB8, "#,#.00"))
        Fgrid.Row = 9
        Fgrid.Text = GetMin((C9), (I9)) 'IIf(C9 = 0, "-", Format(CB9, "#,#.00"))
        Fgrid.Row = 10
        Fgrid.Text = GetMin((C10), (I10)) 'IIf(C10 = 0, "-", Format(CB10, "#,#.00"))
        Fgrid.Row = 11
        Fgrid.Text = GetMin((C11), (I11)) 'IIf(C11 = 0, "-", Format(CB11, "#,#.00"))
        Fgrid.Row = 12
        Fgrid.Text = GetMin((C12), (I12)) 'IIf(C12 = 0, "-", Format(CB12, "#,#.00"))
        Fgrid.Row = 13
        Fgrid.Text = GetMin((C13), (I13)) 'IIf(C13 = 0, "-", Format(CB13, "#,#.00"))
        Fgrid.Row = 14
        Fgrid.Text = GetMin((C14), (I14)) 'IIf(C14 = 0, "-", Format(CB14, "#,#.00"))
        Fgrid.Row = 15
        Fgrid.Text = GetMin((C15), (I15)) 'IIf(C15 = 0, "-", Format(CB15, "#,#.00"))
        Fgrid.Row = 16
        Fgrid.Text = GetMin((C16), (I16)) 'IIf(C16 = 0, "-", Format(CB16, "#,#.00"))
        Fgrid.Row = 17
        Fgrid.Text = GetMin((C17), (I17)) 'IIf(C17 = 0, "-", Format(CB17, "#,#.00"))
        Fgrid.Row = 18
        Fgrid.Text = GetMin((C18), (I18)) 'IIf(C8 = 0, "-", Format(CB8, "#,#.00"))
        Fgrid.Row = 19
        Fgrid.Text = GetMin((C19), (I19)) 'IIf(C9 = 0, "-", Format(CB9, "#,#.00"))
        Fgrid.Row = 20
        Fgrid.Text = GetMin((C20), (I20)) 'IIf(C10 = 0, "-", Format(CB10, "#,#.00"))
        Fgrid.Row = 21
        Fgrid.Text = GetMin((C21), (I21)) 'IIf(C11 = 0, "-", Format(CB11, "#,#.00"))
        Fgrid.Row = 22
        Fgrid.Text = GetMin((C22), (I22)) 'IIf(C12 = 0, "-", Format(CB12, "#,#.00"))
        Fgrid.Row = 23
        Fgrid.Text = GetMin((C23), (I23)) 'IIf(C13 = 0, "-", Format(CB13, "#,#.00"))
        Fgrid.Row = 24
        Fgrid.Text = GetMin((C24), (I24)) 'IIf(C14 = 0, "-", Format(CB14, "#,#.00"))
        Fgrid.Row = 25
        Fgrid.Text = GetMin((C25), (I25)) 'IIf(C15 = 0, "-", Format(CB15, "#,#.00"))
        Fgrid.Row = 26
        Fgrid.Text = GetMin((C26), (I26)) 'IIf(C16 = 0, "-", Format(CB16, "#,#.00"))
        Fgrid.Row = 27
        Fgrid.Text = GetMin((C27), (I27)) 'IIf(C17 = 0, "-", Format(CB17, "#,#.00"))
        Fgrid.Row = 28
        Fgrid.Text = GetMin((C28), (I28)) 'IIf(C16 = 0, "-", Format(CB16, "#,#.00"))
        Fgrid.Row = 29
        Fgrid.Text = GetMin((C29), (I29)) 'IIf(C17 = 0, "-", Format(CB17, "#,#.00"))
        Fgrid.Row = 30
        Fgrid.Text = GetMin((C30), (I30)) 'IIf(C16 = 0, "-", Format(CB16, "#,#.00"))
        Fgrid.Row = 31
        Fgrid.Text = GetMin((C31), (I31)) 'IIf(C17 = 0, "-", Format(CB17, "#,#.00"))
        
    End If
    
    ' Not Payable to CLN
    Fgrid.Col = 5
    Fgrid.Row = 1
    If CBInd <> "1" Then
        Fgrid.Text = DueToFromCLN((C1), (CB1))
             
        Fgrid.Row = 2
        Fgrid.Text = DueToFromCLN((C2), (CB2))
        
        Fgrid.Row = 3
        Fgrid.Text = DueToFromCLN((C3), (CB3))
    
        Fgrid.Row = 4
        Fgrid.Text = DueToFromCLN((C4), (CB4))
        
        Fgrid.Row = 5
        Fgrid.Text = DueToFromCLN((C5), (CB5))
    
        Fgrid.Row = 6
        Fgrid.Text = DueToFromCLN((C6), (CB6))
        
        Fgrid.Row = 7
        Fgrid.Text = DueToFromCLN((C7), (CB7))
    
        Fgrid.Row = 8
        Fgrid.Text = DueToFromCLN((C8), (CB8))
        
        Fgrid.Row = 9
        Fgrid.Text = DueToFromCLN((C9), (CB9))
    
        Fgrid.Row = 10
        Fgrid.Text = DueToFromCLN((C10), (CB10))
        
        Fgrid.Row = 11
        Fgrid.Text = DueToFromCLN((C11), (CB11))
    
        Fgrid.Row = 12
        Fgrid.Text = DueToFromCLN((C12), (CB12))
        
        Fgrid.Row = 13
        Fgrid.Text = DueToFromCLN((C13), (CB13))
    
        Fgrid.Row = 14
        Fgrid.Text = DueToFromCLN((C14), (CB14))
        
        Fgrid.Row = 15
        Fgrid.Text = DueToFromCLN((C15), (CB15))
    
        Fgrid.Row = 16
        Fgrid.Text = DueToFromCLN((C16), (CB16))
        
        Fgrid.Row = 17
        Fgrid.Text = DueToFromCLN((C17), (CB17))
    
        Fgrid.Row = 18
        Fgrid.Text = DueToFromCLN((C18), (CB18))
        
        Fgrid.Row = 19
        Fgrid.Text = DueToFromCLN((C19), (CB19))
    
        Fgrid.Row = 20
        Fgrid.Text = DueToFromCLN((C20), (CB20))
        
        Fgrid.Row = 21
        Fgrid.Text = DueToFromCLN((C21), (CB21))
    
        Fgrid.Row = 22
        Fgrid.Text = DueToFromCLN((C22), (CB22))
        
        Fgrid.Row = 23
        Fgrid.Text = DueToFromCLN((C23), (CB23))
    
        Fgrid.Row = 24
        Fgrid.Text = DueToFromCLN((C24), (CB24))
        
        Fgrid.Row = 25
        Fgrid.Text = DueToFromCLN((C25), (CB25))
    
        Fgrid.Row = 26
        Fgrid.Text = DueToFromCLN((C26), (CB26))
        
        Fgrid.Row = 27
        Fgrid.Text = DueToFromCLN((C27), (CB27))
    
        Fgrid.Row = 28
        Fgrid.Text = DueToFromCLN((C28), (CB28))
        
        Fgrid.Row = 29
        Fgrid.Text = DueToFromCLN((C29), (CB29))
        
        Fgrid.Row = 30
        Fgrid.Text = DueToFromCLN((C30), (CB30))
        
        Fgrid.Row = 31
        Fgrid.Text = DueToFromCLN((C31), (CB31))
        
    End If
        
    Fgrid.Row = 35
    Fgrid.Col = 2
    Fgrid = copay

    Fgrid.Col = 4
    Fgrid = Format(copay * -1, "#,#.00;(#,#.00)")


    'must put back
    'If ANNInd = True Then
    '    Fgrid = "FULL"
    'Else
    Fgrid.Row = 36
    Fgrid.Col = 2

    Fgrid = Format(VisitLimit, "#,#.00;(#,#.00)")
    'End If
    
    Fgrid.Row = 37
    Fgrid.Col = 2
    Fgrid.CellFontBold = True
    Fgrid.CellForeColor = vbBlue
    'must put back
    'If ANNInd = True Then
    '    Fgrid = "FULL"
    'Else
        Fgrid = Format(BalLimit, "#,#.00;(#,#.00)")
    'End If
    
    ' Total for Actual
    Fgrid.Col = 2
    Fgrid.CellFontBold = True
    Fgrid.CellForeColor = vbBlue
    
    Fgrid = getTotal(2)
    
    TotalActual = Fgrid
    ' Total for MBM Approved
    Fgrid.Col = 7
    Fgrid.CellFontBold = True
    Fgrid.CellForeColor = vbBlue
    
    Fgrid = getTotal(7)
    TotalApprove = Fgrid
    
    ' Total for Clinic
    Fgrid.Col = 4
    Fgrid.CellFontBold = True
    Fgrid.CellForeColor = vbBlue
    
    Fgrid = getTotal(4)
    'must put back
    totalCLN = Fgrid
    
    Fgrid = getTotal(5)
    If IsNumeric(Fgrid) = True Then
        TotalCLNExc = Fgrid
    Else
        TotalCLNExc = 0
    End If
    
    'must put back
    'If IsNumeric(totalApprove) = False Then
    '    totalApprove = 0
    'End If
    'must put back
    'Fgrid = Format(CoPayApp * -1, "#,#.00;(#,#.00)")

    'must put back
    'Fgrid = Format(COPay * -1, "#,#.00;(#,#.00)")
    'COPay = Fgrid
    
    ' Total for member
    Fgrid.Row = 39
    Fgrid.Col = 8
    Fgrid.CellFontBold = True
    Fgrid.CellForeColor = vbBlue
    'must put back
    totalMem = getTotal(8)
    If IsNumeric(totalMem) = True And totalMem <> 0 Then
        Fgrid = totalMem
    Else
        Fgrid = "-"
    End If
       
    Fgrid.Col = 7
    Fgrid.Row = 22
    Fgrid.CellFontBold = True
    Fgrid.CellForeColor = vbBlue
    
    Call getFinalTotal
End Sub


Private Function DueToFromMBM(Act As String, App As String)
    If IsNumeric(Approved) = False Then
        Approved = 0
    End If
    If IsNumeric(PaytoHospital) = False Then
        PaytoHospital = 0
    End If
        If IsNumeric(App) = False Then
            App = 0
        End If
        If PAYTo = "C" Then
            If CDbl(Act) > CDbl(App) Then
                DueToFromMBM = Format(CDbl(App) - CDbl(Act), "#,#0.00;(#,#0.00)")
            End If
        ElseIf PAYTo = "M" Then
            DueToFromMBM = Format(App, "#,#0.00;(#,#0.00)")
        End If
    
        If DueToFromMBM = 0# Then
            DueToFromMBM = "-"
        End If
End Function
Private Sub getFinalTotal()
Dim FinalApprove As Double
Dim curRow  As Integer
Dim StrSql As String
Dim CLNRed As New ADODB.Recordset

    If IsNumeric(TotalApprove) = False Then
        TotalApprove = 0
    End If
    
    Fgrid.Row = 35
    Fgrid.Col = 4
    Fgrid = Format(COPayColl * -1, "#,#.00;(#,#.00)")
    
    'Fgrid.Col = 5
    'Fgrid = Format(CDbl(COPayColl) - CDbl(COPay), "#,#.00;(#,#.00)")
    
    Fgrid.Col = 7
    Fgrid = Format(COPayColl * -1, "#,#.00;(#,#.00)")
    
    'Fgrid.Col = 8
    'Fgrid = Format(CDbl(COPayColl) - CDbl(COPay), "#,#.00;(#,#.00)")
    
    Fgrid.Row = 39
    ' Clinic
    Fgrid.Col = 4
    Fgrid.CellFontBold = True
    Fgrid.CellForeColor = vbBlue
        
    'Clinic
    If PAYTo = "C" Then
        If IsNumeric(totalCLN) = False Then
            totalCLN = 0
        End If
        If IsNumeric(CLNRedAmt) = False Then
            CLNRedAmt = 0
        End If
        
            If Len(CLNRedStatus) Then
                StrSql = "Select CLNAmout,CLNPerc,CLNEffDate from CLNRedAmt where 0 = 0"
                CLNRed.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
                
                If CLNRedStatus = 1 Then
                    CLNRedAmt = CLNRed!CLNAmout
                ElseIf CLNRedStatus = 2 Then
                    CLNRedAmt = (CLNRed!CLNPerc * totalCLN) / 100
                End If
                CLNRed.Close
                Set CLNRed = Nothing
            End If
            

            
        totalFCLN = Format(IIf(CDbl(totalCLN) - CDbl(COPayColl) - CDbl(CLNRedAmt), CDbl(totalCLN) - CDbl(COPayColl), 0) - CDbl(CLNRedAmt), "#,#.00;(#,#.00)")
        
        If MCOProcStatus = True And GPStatus = True Then
            MCOProcFee = (MCOProcFeePerc * totalFCLN) / 100
            totalFCLN = CDbl(totalFCLN) - CDbl(MCOProcFee)
        Else
            MCOProcFee = 0
        End If
        
        
            
    End If
    
    Fgrid.Row = 38
        Fgrid = Format(CDbl(MCOProcFee) * -1, "#,#.00;(#,#.00)")
    
    Fgrid.Row = 39
    If totalFCLN <> 0 Then
        Fgrid = totalFCLN
    Else
        Fgrid = "-"
    End If
    

    'Fgrid.Col = 5
    'Fgrid.CellFontBold = True
    'Fgrid.CellForeColor = vbBlue
        

    'TotalFinalCLNExc = Format(IIf(IsNumeric((CDbl(TotalCLNExc) + CDbl(COPay)) - CDbl(COPayColl)), (CDbl(TotalCLNExc) + CDbl(COPay)) - CDbl(COPayColl), 0), "#,#.00;(#,#.00)")
    'If TotalFinalCLNExc <> 0 Then
    '    Fgrid = Format(TotalFinalCLNExc, "#,#0.00;(#,#0.00)")
    'Else
    '    Fgrid = "-"
    'End If


    'Actual
    Fgrid.Col = 2
    Fgrid.CellFontBold = True
    Fgrid.CellForeColor = vbBlue
    Fgrid = Format(TotalActual, "#,#.00;(#,#.00)")
    TotalFActual = Fgrid
    
    Fgrid.Col = 7
    Fgrid.CellFontBold = True
    Fgrid.CellForeColor = vbBlue
    If ANNInd = False Then
        FinalBalLmt = TotalApprove
    Else
        If INVCLNForm = "1" Then
            If IsNumeric(frmInvoice.lstInvoice.SelectedItem.SubItems(70)) = True Then
                FinalBalLmt = CDbl(frmInvoice.lstInvoice.SelectedItem.SubItems(70)) + CDbl(AdjAnnLmt)
            Else
                FinalBalLmt = 0
            End If
        ElseIf INVCLNForm = "2" Then
            If IsNumeric(frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(70)) = True Then
                FinalBalLmt = CDbl(frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(70)) + CDbl(AdjAnnLmt)
            Else
                FinalBalLmt = 0
            End If
            
        ElseIf INVCLNForm = "3" Then
            If IsNumeric(frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(70)) = True Then
                FinalBalLmt = CDbl(frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(70)) + CDbl(AdjAnnLmt)
            Else
                FinalBalLmt = 0
            End If
            
        ElseIf INVCLNForm = "4" Then
            If IsNumeric(frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(70)) = True Then
                FinalBalLmt = CDbl(frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(70)) + CDbl(AdjAnnLmt)
            Else
                FinalBalLmt = 0
            End If
            
        ElseIf INVCLNForm = "5" Then
            If IsNumeric(frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(70)) = True Then
                FinalBalLmt = CDbl(frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(70)) + CDbl(AdjAnnLmt)
            Else
                FinalBalLmt = 0
            End If
            
        End If
    End If
    
    If VISInd = False Or VisitLimit = "FULL" Then
        'If IsNumeric(frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(43)) Then
        '    FinalVisitLmt = TotalApprove + CDbl(frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(43))
        'Else
            FinalVisitLmt = TotalApprove
        'End If
    Else
        FinalVisitLmt = IIf(IsNumeric(VisitLimit) = True, VisitLimit, 0)
    End If


    If CDbl(FinalVisitLmt) > CDbl(FinalBalLmt) Then
        If CDbl(TotalApprove) > CDbl(FinalBalLmt) Then
            Fgrid = Format(CDbl(FinalBalLmt) - CDbl(COPayColl), "#,#.00;(#,#.00)")
        Else
            If TotalApprove <> 0 Then
                Fgrid = Format(CDbl(TotalApprove) - CDbl(COPayColl), "#,#.00;(#,#.00)")
            Else
                Fgrid = "-"
            End If
        End If
    ElseIf CDbl(FinalVisitLmt) <= CDbl(FinalBalLmt) Then
        'If CDbl(TotalApprove) >= CDbl(FinalVisitLmt) Then
         '   FinalApprove = Format(CDbl(FinalVisitLmt) - CDbl(COPayColl), "#,#.00;(#,#.00)")
        'ElseIf CDbl(TotalApprove) < CDbl(FinalVisitLmt) Then
           ' FinalApprove = Format(CDbl(TotalApprove) - CDbl(COPayColl), "#,#.00;(#,#.00)")
        'End If
         FinalApprove = Format(CDbl(TotalApprove) - CDbl(COPayColl), "#,#.00;(#,#.00)") '11072016
    End If
    
    If INVCLNForm = "2" Then
        If PostageStatus = True Then
            Fgrid = Format(CDbl(FinalApprove) - CDbl(copay), "#,#.00;(#,#.00)")
        Else
            Fgrid = Format(CDbl(FinalApprove) - CDbl(copay), "#,#.00;(#,#.00)")
        End If
    Else
        Fgrid = Format(CDbl(FinalApprove) - CDbl(copay), "#,#.00;(#,#.00)")
    End If
    totalFApprove = Fgrid
    
    Fgrid.Col = 8
    Fgrid.CellFontBold = True
    Fgrid.CellForeColor = vbBlue
    If IsNumeric(totalMem) = False Then
        totalMem = 0
    End If
            
    If IsNumeric(COPayColl) Then
        totalFMem = CDbl(totalMem) - CDbl(COPayColl)
    Else
        totalFMem = CDbl(totalMem)
    End If
    
    If IsNumeric(totalMem) = True Then
        totalMem = CDbl(totalMem)
    Else
        totalMem = 0
    End If

    
    If totalFMem = 0 Then
        Fgrid = "-"
    Else
        Fgrid = Format(totalFMem, "#,#.00;(#,#.00)")
        If bChangesMade = False Then
            totalFMemOriginal = totalFMem
        End If
    End If
    
    ' totalFMem  = Total Payable 2 Member  by Medix
    ' ready for debit or credit note
    'If IsNumeric(lblWSNo) = True And bChangesMade = False Then
        'If totalFMem < 0 Then
        '    cmdDebit.Enabled = True
        '    cmdCredit.Enabled = False
        'ElseIf totalFMem > 0 Then
        '    cmdDebit.Enabled = False
        '    cmdCredit.Enabled = True
        'Else
        '    cmdDebit.Enabled = False
        '    cmdCredit.Enabled = False
       'End If
    'End If

    'TotalReimburse = 0
    'totalExcessMemNeed2Pay = 0
    'Fgrid.Col = 7
    
    'For curRow = 1 To 17
    '    Fgrid.Row = curRow
    '    If IsNumeric(Fgrid.Text) = True Then
    '        If CDbl(Fgrid.Text) >= 0 Then
    '            TotalReimburse = TotalReimburse + CDbl(Fgrid.Text)
    '        Else
    '
    '                totalExcessMemNeed2Pay = totalExcessMemNeed2Pay + CDbl(Fgrid.Text)
    '        End If
    '    End If
    'Next
   'Call GetPaymentCollection
End Sub

Private Function getTotal(CurCol As Integer)
    Dim X As Integer
    
    getTotal = 0
    Fgrid.Col = CurCol
    For X = 1 To 31
        Fgrid.Row = X
        If IsNumeric(Fgrid.Text) = True Then
            getTotal = getTotal + CDbl(Fgrid.Text)
        End If
    Next
        
    If getTotal <> 0 Then
        getTotal = Format(getTotal, "#,#.00;(#,#.00)")
    Else
        getTotal = "-"
    End If
    Fgrid.Row = 33
End Function

Private Function DueToFromCLN(ClinicAmt As String, ClinicLimit As String)
    If IsNumeric(ClinicAmt) = False Then
        ClinicAmt = 0
    End If
    If IsNumeric(ClinicLimit) = False Then
        ClinicLimit = 0
    End If
    
    If CDbl(ClinicAmt) > CDbl(ClinicLimit) Then
        DueToFromCLN = Format(CDbl(ClinicLimit) - CDbl(ClinicAmt), "#,#0.00;(#,#0.00)")
    End If

    If DueToFromCLN = 0# Then
        DueToFromCLN = "-"
    End If
End Function

Public Sub display_benefit()
    Dim strsqlCOPay As String
    Dim copay As New ADODB.Recordset
    Dim BNF As New ADODB.Recordset
    Dim StrSql As String
    Dim counter As Integer
    Dim i  As Integer
    Dim strsqlCOV As String
    Dim MBMCov As New ADODB.Recordset
    Dim INV As Integer
    Dim Lmt As Integer
    Dim a As Integer
    Dim PLN As New ADODB.Recordset
    Dim strsqlPLN As String
    Dim PLNBnfStatus As Boolean
    
    Fgrid.Col = 6
    For i = 2 To 31
        Fgrid.Row = i
        Fgrid.Text = "-"
    Next
    
    If INVCLNForm = 1 Or INVCLNForm = 2 Then
        strsqlPLN = "Select BNFStatus , PLNCode from PLN  where PLnCode = '" & Trim(PLNCode) & "'"
        PLN.Open strsqlPLN, MediConnCISDB, adOpenKeyset, adLockOptimistic
            If PLN.RecordCount Then
                If PLN!BNFStatus = "Y" Then
                    PLNBnfStatus = True
                Else
                    PLNBnfStatus = False
                End If
            End If
        PLN.Close
        Set PLN = Nothing
        
        If PLNBnfStatus = True Then
            StrSql = "Select BNFCode , BNFNoOFDays, BNFNoOfAnn, BNFRMPerVisit , BNFPercentage, BNFEffDate,BNFCalculation from BNFNew  " & _
            " where PLnCode = '" & PLNCode & "' and INSCode = '" & Trim(INSCode) & "' and HLTCode = '" & Trim(HLTCode) & "'"
            StrSql = StrSql & " order by BNFCOde "
            BNF.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            
            Do Until BNF.EOF
                 With BNF
                    For counter = 1 To 33 Step 1
                     Fgrid.Row = counter
                     Fgrid.Col = 0
                        If Fgrid.Text = !BNFCode Then
                            ' LImit col
                            Fgrid.Col = 6
                            If !BNFCode = "101" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I1
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I1, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            ElseIf !BNFCode = "102" Then
                                'If I2 > 0 Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I2
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I2, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            ElseIf !BNFCode = "103" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I3
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I3, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            ElseIf !BNFCode = "104" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I4
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I4, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            ElseIf !BNFCode = "105" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I5
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I5, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            ElseIf !BNFCode = "106" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I6
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I6, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            ElseIf !BNFCode = "301" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I7
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I7, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            ElseIf !BNFCode = "302" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I8
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I8, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            ElseIf !BNFCode = "303" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I9
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I9, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            ElseIf !BNFCode = "304" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I10
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I10, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                                    
                            ElseIf !BNFCode = "305" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I11
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I11, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            ElseIf !BNFCode = "306" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I12
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I12, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            ElseIf !BNFCode = "307" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I13
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
    
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I13, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                                    
                            ElseIf !BNFCode = "308" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I14
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I14, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                                    
                            ElseIf !BNFCode = "309" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I15
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I15, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            ElseIf !BNFCode = "401" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I16
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I16, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            ElseIf !BNFCode = "402" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I17
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I17, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            ElseIf !BNFCode = "403" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I18
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I18, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                                    
                            ElseIf !BNFCode = "404" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I19
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I19, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            ElseIf !BNFCode = "405" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I20
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I20, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            ElseIf !BNFCode = "501" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I21
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I21, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            ElseIf !BNFCode = "502" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I22
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I22, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            ElseIf !BNFCode = "503" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I23
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I23, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            ElseIf !BNFCode = "504" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I24
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I24, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            ElseIf !BNFCode = "505" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I25
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I25, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            ElseIf !BNFCode = "601" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I26
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I26, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            ElseIf !BNFCode = "602" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I27
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I27, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            ElseIf !BNFCode = "603" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I28
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I28, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            ElseIf !BNFCode = "604" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I29
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I29, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            
                            ElseIf !BNFCode = "701" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I30
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I30, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            ElseIf !BNFCode = "702" Then
                                If IsNumeric(!BNFPercentage) = False Then
                                    INV = 0
                                Else
                                    INV = (!BNFPercentage / 100) * I31
                                End If
                                If IsNumeric(!BNFRMPerVisit) = False Then
                                    Lmt = 0
                                Else
                                    Lmt = !BNFRMPerVisit
                                End If
                                If !BNFCalculation = "1" Then
                                    Fgrid.Text = Format(GetMin((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "2" Then
                                    Fgrid.Text = Format(GetMax((INV), (Lmt)), "#,##0.00")
                                ElseIf !BNFCalculation = "3" Then
                                    Fgrid.Text = Format(I31, "#,##0.00")
                                ElseIf !BNFCalculation = "4" Then
                                    Fgrid.Text = Format(Lmt, "#,##0.00")
                                End If
                            
                            ElseIf !BNFRMPerVisit = 0 Then
                                Fgrid.Text = "-"
                            End If
                        End If
                        Exit For
                    Next
                    .MoveNext
                    End With
                    Loop
                    
                ElseIf PLNBnfStatus = False Then
                    ' LImit col
                    Fgrid.Col = 6
                        Fgrid.Row = 1
                        Fgrid.Text = Format(I1, "#,##0.00")
                        Fgrid.Row = 2
                        Fgrid.Text = Format(I2, "#,##0.00")
                        Fgrid.Row = 3
                        Fgrid.Text = Format(I3, "#,##0.00")
                        Fgrid.Row = 4
                        Fgrid.Text = Format(I4, "#,##0.00")
                        Fgrid.Row = 5
                        Fgrid.Text = Format(I5, "#,##0.00")
                        Fgrid.Row = 6
                        Fgrid.Text = Format(I6, "#,##0.00")
                        Fgrid.Row = 7
                        Fgrid.Text = Format(I7, "#,##0.00")
                        Fgrid.Row = 8
                        Fgrid.Text = Format(I8, "#,##0.00")
                        Fgrid.Row = 9
                        Fgrid.Text = Format(I9, "#,##0.00")
                        Fgrid.Row = 10
                        Fgrid.Text = Format(I10, "#,##0.00")
                        Fgrid.Row = 11
                        Fgrid.Text = Format(I11, "#,##0.00")
                        Fgrid.Row = 12
                        Fgrid.Text = Format(I12, "#,##0.00")
                        Fgrid.Row = 13
                        Fgrid.Text = Format(I13, "#,##0.00")
                        Fgrid.Row = 14
                        Fgrid.Text = Format(I14, "#,##0.00")
                        Fgrid.Row = 15
                        Fgrid.Text = Format(I15, "#,##0.00")
                        Fgrid.Row = 16
                        Fgrid.Text = Format(I16, "#,##0.00")
                        Fgrid.Row = 17
                        Fgrid.Text = Format(I17, "#,##0.00")
                        Fgrid.Row = 18
                        Fgrid.Text = Format(I18, "#,##0.00")
                        Fgrid.Row = 19
                        Fgrid.Text = Format(I19, "#,##0.00")
                        Fgrid.Row = 20
                        Fgrid.Text = Format(I20, "#,##0.00")
                        Fgrid.Row = 21
                        Fgrid.Text = Format(I21, "#,##0.00")
                        Fgrid.Row = 22
                        Fgrid.Text = Format(I22, "#,##0.00")
                        Fgrid.Row = 23
                        Fgrid.Text = Format(I23, "#,##0.00")
                        Fgrid.Row = 24
                        Fgrid.Text = Format(I24, "#,##0.00")
                        Fgrid.Row = 25
                        Fgrid.Text = Format(I25, "#,##0.00")
                        Fgrid.Row = 26
                        Fgrid.Text = Format(I26, "#,##0.00")
                        Fgrid.Row = 27
                        Fgrid.Text = Format(I27, "#,##0.00")
                        Fgrid.Row = 28
                        Fgrid.Text = Format(I28, "#,##0.00")
                        Fgrid.Row = 29
                        Fgrid.Text = Format(I29, "#,##0.00")
                        Fgrid.Row = 30
                        Fgrid.Text = Format(I30, "#,##0.00")
                        Fgrid.Row = 31
                        Fgrid.Text = Format(I31, "#,##0.00")
                        
                        'to Display Calculation Type
                        'If Len(!BNFCalculation) Then
                        '    strsqlCOPay = "Select COPCode , COPDesc from COPayment  " & _
                        '    " where COPCode = '" & !BNFCalculation & "'"
                        '    COPAY.Open strsqlCOPay, MediConnCISDB, adOpenKeyset, adLockOptimistic
                            
                        '    If COPAY.RecordCount Then
                        '        Fgrid.Col = 2
                        '        Fgrid.Text = !BNFCalculation & "-" & COPAY!COPDesc
                        '    End If
                         '   COPAY.Close
                         '   Set COPAY = Nothing
                        'End If
                    End If
    
    If PLNBnfStatus = True Then
        BNF.Close
        Set BNF = Nothing
    End If
    
    ElseIf INVCLNForm = 3 Then
        a = 296
        If frmInvoiceAdj.lstInvoice.ListItems.Count > 0 Then
            'DO UNTIL frmInvoiceAdj.lstInvoice.E
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 0
                Fgrid.Col = 6
                a = a + 1
                Fgrid.Text = Format(frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(a), "#,##0.00")
            Next
        End If
        
    ElseIf INVCLNForm = 4 Then
        a = 296
        If frmInvoiceAdjAftBordx.lstInvoice.ListItems.Count > 0 Then
            'DO UNTIL frmInvoiceAdj.lstInvoice.E
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 0
                Fgrid.Col = 6
                a = a + 1
                Fgrid.Text = Format(frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(a), "#,##0.00")
            Next
        End If
        
    ElseIf INVCLNForm = 5 Then
        a = 296
        If frmInvoiceAdjAfterPost.lstInvoice.ListItems.Count > 0 Then
            'DO UNTIL frmInvoiceAdj.lstInvoice.E
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 0
                Fgrid.Col = 6
                a = a + 1
                Fgrid.Text = Format(frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(a), "#,##0.00")
            Next
        End If
    End If
End Sub

Private Function GetMax(X As String, Y As String)
    If IsNumeric(Y) = True Then
        If CDbl(X) > CDbl(Y) Then
            GetMax = IIf(CDbl(X) = 0, "-", Format(Y, "#,#.00"))
        Else
            GetMax = IIf(CDbl(Y) = 0, "-", Format(Y, "#,#.00"))
        End If
    Else
        If ChkStr(Y) = "FULL" Then
            GetMax = Format(X, "#,#.00")
        Else
            GetMax = 0
        End If
        
    End If
    If GetMax = 0 Then
        GetMax = "-"
    End If
End Function

Private Function GetMin(X As String, Y As String)
    If IsNumeric(Y) = True Then
        If CDbl(X) > CDbl(Y) Then
            GetMin = IIf(CDbl(Y) = 0, "-", Format(Y, "#,#.00"))
        Else
            GetMin = IIf(CDbl(X) = 0, "-", Format(X, "#,#.00"))
        End If
    Else
        If ChkStr(Y) = "FULL" Then
            GetMin = Format(X, "#,#.00")
        Else
            GetMin = 0
        End If
        
    End If
    If GetMin = 0 Then
        GetMin = "-"
    End If
End Function

Private Sub SaveCLMActual()
Dim i As Integer
Dim counter As Integer
Dim lstRec As ListItems
    i = 167
    If INVCLNForm = 1 Then
        If frmInvoice.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 2
                i = CDbl(i) + 1
            If InvoiceClick = 1 Then
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoice.lstRec.SubItems(i) = 0
                Else
                    frmInvoice.lstRec.SubItems(i) = Fgrid.Text
                End If
            Else
                If IsNumeric(frmInvoice.lstInvoice.SelectedItem.SubItems(i)) = False Then
                    frmInvoice.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoice.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            End If
            Next
        End If
    ElseIf INVCLNForm = 2 Then
        If frmInvoiceMBM.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 2
                i = CDbl(i) + 1
            If InvoiceClick = 1 Then
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceMBM.lstRec.SubItems(i) = 0
                Else
                    frmInvoiceMBM.lstRec.SubItems(i) = Fgrid.Text
                End If
            Else
                If IsNumeric(frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(i)) = False Then
                    frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            End If
            Next
        End If
    ElseIf INVCLNForm = 3 Then
        If frmInvoiceAdj.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 2
                i = CDbl(i) + 1
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            Next
        End If
        
    ElseIf INVCLNForm = 4 Then
        If frmInvoiceAdjAftBordx.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 2
                i = CDbl(i) + 1
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            Next
        End If
        
    ElseIf INVCLNForm = 5 Then
        If frmInvoiceAdjAfterPost.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 2
                i = CDbl(i) + 1
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            Next
        End If
        
    End If
End Sub

Private Sub SaveCLMContRate()
Dim i As Integer
Dim counter As Integer
Dim lstRec As ListItem
i = 210
    If INVCLNForm = 1 Then
        If frmInvoice.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 2
                i = i + 1
            If InvoiceClick = 1 Then
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoice.lstRec.SubItems(i) = 0
                Else
                    frmInvoice.lstRec.SubItems(i) = Fgrid.Text
                End If
            Else
                If IsNumeric(frmInvoice.lstInvoice.SelectedItem.SubItems(i)) = False Then
                    frmInvoice.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoice.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            End If
            Next
        End If
    ElseIf INVCLNForm = 2 Then
        If frmInvoiceMBM.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 2
                i = i + 1
            If InvoiceClick = 1 Then
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceMBM.lstRec.SubItems(i) = 0
                Else
                    frmInvoiceMBM.lstRec.SubItems(i) = Fgrid.Text
                End If
            Else
                If IsNumeric(frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(i)) = False Then
                    frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            End If
                
            Next
        End If
    ElseIf INVCLNForm = 3 Then
        If frmInvoiceAdj.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 2
                i = i + 1
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            Next
        End If
        
    ElseIf INVCLNForm = 4 Then
        If frmInvoiceAdjAftBordx.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 2
                i = i + 1
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            Next
        End If
    ElseIf INVCLNForm = 5 Then
        If frmInvoiceAdjAfterPost.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 2
                i = i + 1
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            Next
        End If
        
    End If
End Sub

Private Sub SavePayCLN()
Dim i As Integer
Dim counter As Integer
Dim lstRec As ListItem
i = 425
    
    If INVCLNForm = 1 Then
        If frmInvoice.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 2
                i = i + 1
            If InvoiceClick = 1 Then
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoice.lstRec.SubItems(i) = 0
                Else
                    frmInvoice.lstRec.SubItems(i) = Fgrid.Text
                End If
            Else
                If IsNumeric(frmInvoice.lstInvoice.SelectedItem.SubItems(i)) = False Then
                    frmInvoice.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoice.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            End If
                
            Next
        End If
    ElseIf INVCLNForm = 2 Then
        If frmInvoiceMBM.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 2
                i = i + 1
            If InvoiceClick = 1 Then
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceMBM.lstRec.SubItems(i) = 0
                Else
                    frmInvoiceMBM.lstRec.SubItems(i) = Fgrid.Text
                End If
            Else
                If IsNumeric(frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(i)) = False Then
                    frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            End If
                
            Next
        End If
    ElseIf INVCLNForm = 3 Then
        If frmInvoiceAdj.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 2
                i = i + 1
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            Next
        End If
    ElseIf INVCLNForm = 4 Then
        If frmInvoiceAdjAftBordx.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 2
                i = i + 1
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            Next
        End If
    ElseIf INVCLNForm = 5 Then
        If frmInvoiceAdjAfterPost.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 2
                i = i + 1
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            Next
        End If
    
    End If
End Sub

Private Sub SaveNotPayCLN()
Dim i As Integer
Dim counter As Integer
Dim lstRec As ListItem
i = 382
    If INVCLNForm = 1 Then
        If frmInvoice.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 5
                i = i + 1
            If InvoiceClick = 1 Then
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoice.lstRec.SubItems(i) = 0
                Else
                    frmInvoice.lstRec.SubItems(i) = Fgrid.Text
                End If
            Else
                If IsNumeric(frmInvoice.lstInvoice.SelectedItem.SubItems(i)) = False Then
                    frmInvoice.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoice.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            End If
                
            Next
        End If
    ElseIf INVCLNForm = 2 Then
        If frmInvoiceMBM.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 5
                i = i + 1
            If InvoiceClick = 1 Then
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceMBM.lstRec.SubItems(i) = 0
                Else
                    frmInvoiceMBM.lstRec.SubItems(i) = Fgrid.Text
                End If
            Else
                If IsNumeric(frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(i)) = False Then
                    frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            End If
                
            Next
        End If
    ElseIf INVCLNForm = 3 Then
        If frmInvoiceAdj.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 5
                i = i + 1
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            Next
        End If
    ElseIf INVCLNForm = 4 Then
        If frmInvoiceAdjAftBordx.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 5
                i = i + 1
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            Next
        End If
    
    ElseIf INVCLNForm = 5 Then
        If frmInvoiceAdjAfterPost.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 5
                i = i + 1
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            Next
        End If
        
    
    End If
End Sub

Private Sub SaveMBMBnfLmt()
Dim i As Integer
Dim counter As Integer
Dim lstRec As ListItem
i = 296

    If INVCLNForm = 1 Then
        If frmInvoice.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 2
                i = i + 1
            If InvoiceClick = 1 Then
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoice.lstRec.SubItems(i) = 0
                Else
                    frmInvoice.lstRec.SubItems(i) = Fgrid.Text
                End If
            Else
                If IsNumeric(frmInvoice.lstInvoice.SelectedItem.SubItems(i)) = False Then
                    frmInvoice.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoice.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            End If
            Next
        End If
    ElseIf INVCLNForm = 2 Then
        If frmInvoiceMBM.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 2
                i = i + 1
            If InvoiceClick = 1 Then
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceMBM.lstRec.SubItems(i) = 0
                Else
                    frmInvoiceMBM.lstRec.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            Else
                If IsNumeric(frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(i)) = False Then
                    frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            End If
                
            Next
        End If
    ElseIf INVCLNForm = 3 Then
        If frmInvoiceAdj.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 2
                i = i + 1
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            Next
        End If
        
    ElseIf INVCLNForm = 4 Then
        If frmInvoiceAdjAftBordx.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 2
                i = i + 1
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            Next
        End If
        
    ElseIf INVCLNForm = 5 Then
        If frmInvoiceAdjAfterPost.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 2
                i = i + 1
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            Next
        End If
        
    End If
End Sub

Private Sub SaveMBMApp()
Dim i As Integer
Dim counter As Integer
Dim lstRec As ListItem
    
    i = 253
    If INVCLNForm = 1 Then
        If frmInvoice.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 2
                i = i + 1
            If InvoiceClick = 1 Then
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoice.lstRec.SubItems(i) = 0
                Else
                    frmInvoice.lstRec.SubItems(i) = Fgrid.Text
                End If
            Else
                If IsNumeric(frmInvoice.lstInvoice.SelectedItem.SubItems(i)) = False Then
                    frmInvoice.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoice.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            End If
    
            Next
        End If
    ElseIf INVCLNForm = 2 Then
        If frmInvoiceMBM.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 2
                i = i + 1
            If InvoiceClick = 1 Then
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceMBM.lstRec.SubItems(i) = 0
                Else
                    frmInvoiceMBM.lstRec.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            Else
                If IsNumeric(frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(i)) = False Then
                    frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            End If
    
            Next
        End If
    ElseIf INVCLNForm = 3 Then
        If frmInvoiceAdj.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 2
                i = i + 1
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            Next
        End If
        
    ElseIf INVCLNForm = 4 Then
        If frmInvoiceAdjAftBordx.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 2
                i = i + 1
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            Next
        End If
        
    ElseIf INVCLNForm = 5 Then
        If frmInvoiceAdjAfterPost.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 2
                i = i + 1
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            Next
        End If
        
    End If
End Sub

Private Sub SaveMBMExc()
Dim i As Integer
Dim counter As Integer
Dim lstRec As ListItem
i = 339
    If INVCLNForm = 1 Then
        If frmInvoice.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 8
                i = i + 1
            If InvoiceClick = 1 Then
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoice.lstRec.SubItems(i) = 0
                Else
                    frmInvoice.lstRec.SubItems(i) = Fgrid.Text
                End If
            Else
                If IsNumeric(frmInvoice.lstInvoice.SelectedItem.SubItems(i)) = False Then
                    frmInvoice.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoice.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            End If
                
            Next
        End If
    ElseIf INVCLNForm = 2 Then
        If frmInvoiceMBM.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 8
                i = i + 1
            If InvoiceClick = 1 Then
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceMBM.lstRec.SubItems(i) = 0
                Else
                    frmInvoiceMBM.lstRec.SubItems(i) = Fgrid.Text
                End If
            Else
                If IsNumeric(frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(i)) = False Then
                    frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            End If
                
            Next
        End If
    ElseIf INVCLNForm = 3 Then
        If frmInvoiceAdj.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 8
                i = i + 1
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            Next
        End If
    ElseIf INVCLNForm = 4 Then
        If frmInvoiceAdjAftBordx.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 8
                i = i + 1
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            Next
        End If
    ElseIf INVCLNForm = 5 Then
        If frmInvoiceAdjAfterPost.lstInvoice.ListItems.Count > 0 Then
            For counter = 1 To 31 Step 1
                Fgrid.Row = counter
                Fgrid.Col = 8
                i = i + 1
                If IsNumeric(Fgrid.Text) = False Then
                    frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(i) = 0
                Else
                    frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(i) = IIf(IsNumeric(Fgrid.Text) = True, Fgrid.Text, 0)
                End If
            Next
        End If
        
    End If
End Sub

Private Sub SaveClaim()
Dim lstRec As ListItem

    If IsNumeric(TotalFActual) = False Then
        TotalFActual = 0
    End If
    
    If INVCLNForm = 1 Then
        If frmInvoice.lstInvoice.ListItems.Count > 0 Then
            If InvoiceClick = 1 Then
                frmInvoice.lstRec.SubItems(150) = Date
                frmInvoice.lstRec.SubItems(151) = ChkStr(Logon)
                frmInvoice.lstRec.SubItems(152) = ChkStr(Logon)
                frmInvoice.lstRec.SubItems(153) = Date
                frmInvoice.lstRec.SubItems(154) = 0
                frmInvoice.lstRec.SubItems(155) = "C"
                frmInvoice.lstRec.SubItems(156) = ChkStr(Logon)
                frmInvoice.lstRec.SubItems(157) = Date
                frmInvoice.lstRec.SubItems(158) = IIf(IsNumeric(totalFCLN) = True, totalFCLN, 0)
                frmInvoice.lstRec.SubItems(159) = IIf(IsNumeric(TotalFinalCLNExc) = True, TotalFinalCLNExc, 0)
                frmInvoice.lstRec.SubItems(160) = IIf(IsNumeric(TotalFActual) = True, CDbl(TotalFActual) + CDbl(NONGPAmt), 0 + CDbl(NONGPAmt))
                frmInvoice.lstRec.SubItems(161) = IIf(IsNumeric(totalFApprove) = True, totalFApprove, 0)
                frmInvoice.lstRec.SubItems(162) = IIf(IsNumeric(totalFMem) = True, totalFMem, 0)
                frmInvoice.lstRec.SubItems(163) = FinalBalLmt - totalFApprove
                frmInvoice.lstRec.SubItems(164) = FinalBalLmt
                frmInvoice.lstRec.SubItems(165) = 0
                frmInvoice.lstRec.SubItems(166) = 0
                frmInvoice.lstRec.SubItems(468) = MCOProcFee
                
            Else
                frmInvoice.lstInvoice.SelectedItem.SubItems(150) = Date
                frmInvoice.lstInvoice.SelectedItem.SubItems(151) = ChkStr(Logon)
                frmInvoice.lstInvoice.SelectedItem.SubItems(152) = ChkStr(Logon)
                frmInvoice.lstInvoice.SelectedItem.SubItems(153) = Date
                frmInvoice.lstInvoice.SelectedItem.SubItems(154) = 0
                frmInvoice.lstInvoice.SelectedItem.SubItems(155) = "C"
                frmInvoice.lstInvoice.SelectedItem.SubItems(156) = ChkStr(Logon)
                frmInvoice.lstInvoice.SelectedItem.SubItems(157) = Date
                frmInvoice.lstInvoice.SelectedItem.SubItems(158) = IIf(IsNumeric(totalFCLN) = True, totalFCLN, 0)
                frmInvoice.lstInvoice.SelectedItem.SubItems(159) = IIf(IsNumeric(TotalFinalCLNExc) = True, TotalFinalCLNExc, 0)
                frmInvoice.lstInvoice.SelectedItem.SubItems(160) = IIf(IsNumeric(TotalFActual) = True, CDbl(TotalFActual) + CDbl(NONGPAmt), 0 + CDbl(NONGPAmt))
                frmInvoice.lstInvoice.SelectedItem.SubItems(161) = IIf(IsNumeric(totalFApprove) = True, totalFApprove, 0)
                frmInvoice.lstInvoice.SelectedItem.SubItems(162) = IIf(IsNumeric(totalFMem) = True, totalFMem, 0)
                frmInvoice.lstInvoice.SelectedItem.SubItems(163) = FinalBalLmt - totalFApprove
                frmInvoice.lstInvoice.SelectedItem.SubItems(164) = FinalBalLmt
                frmInvoice.lstInvoice.SelectedItem.SubItems(165) = 0
                frmInvoice.lstInvoice.SelectedItem.SubItems(166) = 0
                frmInvoice.lstInvoice.SelectedItem.SubItems(468) = MCOProcFee
            End If
        End If
    ElseIf INVCLNForm = 2 Then
        If frmInvoiceMBM.lstInvoice.ListItems.Count > 0 Then
            If InvoiceClick = 1 Then
                frmInvoiceMBM.lstRec.SubItems(150) = Date
                frmInvoiceMBM.lstRec.SubItems(151) = ChkStr(Logon)
                frmInvoiceMBM.lstRec.SubItems(152) = ChkStr(Logon)
                frmInvoiceMBM.lstRec.SubItems(153) = Date
                frmInvoiceMBM.lstRec.SubItems(154) = 0
                frmInvoiceMBM.lstRec.SubItems(155) = "C"
                frmInvoiceMBM.lstRec.SubItems(156) = ChkStr(Logon)
                frmInvoiceMBM.lstRec.SubItems(157) = Date
                frmInvoiceMBM.lstRec.SubItems(158) = IIf(IsNumeric(totalFCLN) = True, totalFCLN, 0)
                frmInvoiceMBM.lstRec.SubItems(159) = IIf(IsNumeric(TotalFinalCLNExc) = True, TotalFinalCLNExc, 0)
                frmInvoiceMBM.lstRec.SubItems(160) = IIf(IsNumeric(TotalFActual) = True, CDbl(TotalFActual) + CDbl(NONGPAmt), 0 + CDbl(NONGPAmt))
                If frmInvoiceMBM.lstRec.SubItems(476) = "" Then
                    frmInvoiceMBM.lstRec.SubItems(161) = IIf(IsNumeric(totalFApprove) = True, totalFApprove, 0)
                    frmInvoiceMBM.lstRec.SubItems(162) = IIf(IsNumeric(totalFMem) = True, totalFMem, 0)
                Else
                    If IsNumeric(frmInvoiceMBM.lstRec.SubItems(476)) And frmInvoiceMBM.lstRec.SubItems(476) > 0 Then
                        frmInvoiceMBM.lstRec.SubItems(161) = frmInvoiceMBM.lstRec.SubItems(476)
                        frmInvoiceMBM.lstRec.SubItems(162) = 0
                    Else
                        frmInvoiceMBM.lstRec.SubItems(161) = IIf(IsNumeric(totalFApprove) = True, totalFApprove, 0)
                        frmInvoiceMBM.lstRec.SubItems(162) = IIf(IsNumeric(totalFMem) = True, totalFMem, 0)
                    End If
                End If
                
                'frmInvoiceMBM.lstRec.SubItems(161) = IIf(IsNumeric(totalFApprove) = True, totalFApprove, 0)
                'frmInvoiceMBM.lstRec.SubItems(162) = IIf(IsNumeric(totalFMem) = True, totalFMem, 0)
                frmInvoiceMBM.lstRec.SubItems(163) = FinalBalLmt - totalFApprove
                frmInvoiceMBM.lstRec.SubItems(164) = FinalBalLmt
                frmInvoiceMBM.lstRec.SubItems(165) = 0
                frmInvoiceMBM.lstRec.SubItems(166) = 0
            Else
                frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(150) = Date
                frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(151) = ChkStr(Logon)
                frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(152) = ChkStr(Logon)
                frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(153) = Date
                frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(154) = 0
                frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(155) = "C"
                frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(156) = ChkStr(Logon)
                frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(157) = Date
                frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(158) = IIf(IsNumeric(totalFCLN) = True, totalFCLN, 0)
                frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(159) = IIf(IsNumeric(TotalFinalCLNExc) = True, TotalFinalCLNExc, 0)
                frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(160) = IIf(IsNumeric(TotalFActual) = True, CDbl(TotalFActual) + CDbl(NONGPAmt), 0 + CDbl(NONGPAmt))
                
                If frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(476) = "" Then
                    frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(161) = IIf(IsNumeric(totalFApprove) = True, totalFApprove, 0)
                    frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(162) = IIf(IsNumeric(totalFMem) = True, totalFMem, 0)
                Else
                    If IsNumeric(frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(476)) And frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(476) > 0 Then
                        frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(161) = frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(476)
                        frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(162) = 0
                    Else
                        frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(161) = IIf(IsNumeric(totalFApprove) = True, totalFApprove, 0)
                        frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(162) = IIf(IsNumeric(totalFMem) = True, totalFMem, 0)
                    End If
                End If
                
                frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(163) = FinalBalLmt - totalFApprove
                frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(164) = FinalBalLmt
                frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(165) = 0
                frmInvoiceMBM.lstInvoice.SelectedItem.SubItems(166) = 0
            End If
        End If
    ElseIf INVCLNForm = 3 Then
        If frmInvoiceAdj.lstInvoice.ListItems.Count > 0 Then
                frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(150) = Date
                frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(151) = ChkStr(Logon)
                frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(152) = ChkStr(Logon)
                frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(153) = Date
                frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(154) = 0
                frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(155) = "C"
                frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(156) = ChkStr(Logon)
                frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(157) = Date
                frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(158) = IIf(IsNumeric(totalFCLN) = True, totalFCLN, 0)
                frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(159) = IIf(IsNumeric(TotalFinalCLNExc) = True, TotalFinalCLNExc, 0)
                frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(160) = IIf(IsNumeric(TotalFActual) = True, CDbl(TotalFActual) + CDbl(NONGPAmt), 0 + CDbl(NONGPAmt))
                If frmInvoiceAdj.tmpPayto = "M" Then
                    If frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(475) = "" Then
                        frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(161) = IIf(IsNumeric(totalFApprove) = True, totalFApprove, 0)
                        frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(162) = IIf(IsNumeric(totalFMem) = True, totalFMem, 0)
                    Else
                    
                        If IsNumeric(frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(475)) And frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(475) > 0 Then
                            frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(161) = frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(475)
                            frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(162) = 0
                        Else
                            frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(161) = IIf(IsNumeric(totalFApprove) = True, totalFApprove, 0)
                            frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(162) = IIf(IsNumeric(totalFMem) = True, totalFMem, 0)
                        End If
                    End If
                Else
                    frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(161) = IIf(IsNumeric(totalFApprove) = True, totalFApprove, 0)
                    frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(162) = IIf(IsNumeric(totalFMem) = True, totalFMem, 0)
                End If
                frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(163) = FinalBalLmt - totalFApprove
                frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(164) = FinalBalLmt
                frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(165) = 0
                frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(166) = 0
                frmInvoiceAdj.lstInvoice.SelectedItem.SubItems(468) = MCOProcFee
        End If
    ElseIf INVCLNForm = 4 Then
        If frmInvoiceAdjAftBordx.lstInvoice.ListItems.Count > 0 Then
                frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(150) = Date
                frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(151) = ChkStr(Logon)
                frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(152) = ChkStr(Logon)
                frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(153) = Date
                frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(154) = 0
                frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(155) = "C"
                frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(156) = ChkStr(Logon)
                frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(157) = Date
                frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(158) = IIf(IsNumeric(totalFCLN) = True, totalFCLN, 0)
                frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(159) = IIf(IsNumeric(TotalFinalCLNExc) = True, TotalFinalCLNExc, 0)
                frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(160) = IIf(IsNumeric(TotalFActual) = True, CDbl(TotalFActual) + CDbl(NONGPAmt), 0 + CDbl(NONGPAmt))
                If frmInvoiceAdjAftBordx.tmpPayto = "M" Then
                    If frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(475) = "" Then
                        frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(161) = IIf(IsNumeric(totalFApprove) = True, totalFApprove, 0)
                        frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(162) = IIf(IsNumeric(totalFMem) = True, totalFMem, 0)
                    Else
                    
                        If IsNumeric(frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(475)) And frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(475) > 0 Then
                            frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(161) = frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(475)
                            frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(162) = 0
                        Else
                            frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(161) = IIf(IsNumeric(totalFApprove) = True, totalFApprove, 0)
                            frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(162) = IIf(IsNumeric(totalFMem) = True, totalFMem, 0)
                        End If
                    End If
                Else
                    frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(161) = IIf(IsNumeric(totalFApprove) = True, totalFApprove, 0)
                    frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(162) = IIf(IsNumeric(totalFMem) = True, totalFMem, 0)
                End If
                frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(163) = FinalBalLmt - totalFApprove
                frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(164) = FinalBalLmt
                frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(165) = 0
                frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(166) = 0
                frmInvoiceAdjAftBordx.lstInvoice.SelectedItem.SubItems(468) = MCOProcFee
        End If
    ElseIf INVCLNForm = 5 Then
        If frmInvoiceAdjAfterPost.lstInvoice.ListItems.Count > 0 Then
                frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(150) = Date
                frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(151) = ChkStr(Logon)
                frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(152) = ChkStr(Logon)
                frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(153) = Date
                frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(154) = 0
                frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(155) = "C"
                frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(156) = ChkStr(Logon)
                frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(157) = Date
                frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(158) = IIf(IsNumeric(totalFCLN) = True, totalFCLN, 0)
                frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(159) = IIf(IsNumeric(TotalFinalCLNExc) = True, TotalFinalCLNExc, 0)
                frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(160) = IIf(IsNumeric(TotalFActual) = True, CDbl(TotalFActual) + CDbl(NONGPAmt), 0 + CDbl(NONGPAmt))
                If frmInvoiceAdjAfterPost.tmpPayto = "M" Then
                    If frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(475) = "" Then
                        frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(161) = IIf(IsNumeric(totalFApprove) = True, totalFApprove, 0)
                        frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(162) = IIf(IsNumeric(totalFMem) = True, totalFMem, 0)
                    Else
                    
                        If IsNumeric(frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(475)) And frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(475) > 0 Then
                            frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(161) = frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(475)
                            frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(162) = 0
                        Else
                            frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(161) = IIf(IsNumeric(totalFApprove) = True, totalFApprove, 0)
                            frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(162) = IIf(IsNumeric(totalFMem) = True, totalFMem, 0)
                        End If
                    End If
                Else
                    frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(161) = IIf(IsNumeric(totalFApprove) = True, totalFApprove, 0)
                    frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(162) = IIf(IsNumeric(totalFMem) = True, totalFMem, 0)
                End If
                frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(163) = FinalBalLmt - totalFApprove
                frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(164) = FinalBalLmt
                frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(165) = 0
                frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(166) = 0
                frmInvoiceAdjAfterPost.lstInvoice.SelectedItem.SubItems(468) = MCOProcFee
        End If
        
    End If
End Sub

Private Sub GetCLNItems()
Dim StrSql As String
Dim CLN As New ADODB.Recordset

    StrSql = "Select CLNCode, CLNMCOProcfeePerc from CLN  where CLNCode = '" & Trim(CLNCode) & "'"
    CLN.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If CLN.RecordCount Then
            MCOProcFeePerc = IIf(IsNumeric(CLN!CLNMCOProcfeePerc), CLN!CLNMCOProcfeePerc, 0)
        Else
            MCOProcFeePerc = "0"
        End If
    CLN.Close
    Set CLN = Nothing

End Sub






