VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmPrint 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6870
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6870
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Tag             =   "xyf"
      Top             =   480
      Width           =   6735
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "Exit"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmPrint.frx":0000
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   240
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   6960
      m_MDI_Height    =   6075
   End
   Begin VB.Menu mnuStatus 
      Caption         =   "Popupmenu"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public QryPrn As String
Public QryPrn2 As String
Public QryPrn3 As String
Public QryPrn4 As String
Public MySource As String
Dim rsPrint As New ADODB.Recordset
Public ParamValue1 As String
Public ParamValue2 As String
Public ParamValue3 As String
Public ParamValue4 As String
Public ParamValue5 As String
Public ParamValue6 As String
Public ParamValue7 As String
Public ParamValue8 As String
Public ParamValue9 As String
Public ParamValue10 As String
'******************* CODE UPDATED: 08/04/2017 ******************
Public ParamValue11 As String
Public ParamValue12 As String
Public ParamValue13 As String
Public ParamValue14 As String
Public ParamValue15 As String
Public ParamValue16 As String
'*************** END CODE UPDATED: 08/04/2017 ******************



Private Sub cmdExit_Click()
    Unload Me
    'findBBO.Show
End Sub

Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
    If MySource = "ADJUSTMENT REQUEST" Then
        If frmOtherAdjustment.txtRequestStatus.Text = "PENDING" Then
            ans = MsgBox("Do you want to Print? This will Finalized your Request!", vbQuestion + vbYesNo, "Print")
            If ans = vbNo Then Exit Sub
            Call frmOtherAdjustment.ProcessingAR
        End If
    End If
End Sub

Private Sub Form_Load()
    
    Select Case MySource
        Case "PICKBBO"
            Dim crPICKBBO As New rptPrntBBO
            
            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crPICKBBO.DiscardSavedData
            crPICKBBO.Database.SetDataSource rsPrint
            crPICKBBO.ParameterFields(1).ClearCurrentValueAndRange
            crPICKBBO.ParameterFields(1).AddCurrentValue ParamValue1
            CRViewer1.ReportSource = crPICKBBO
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = False
            
        Case "BRANCH RECEIVING DISCREPANCY PRINT"
            Dim crBrnchRecDis As New rptRecDis
            
            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crBrnchRecDis.DiscardSavedData
            crBrnchRecDis.Database.SetDataSource rsPrint
            crBrnchRecDis.ParameterFields(1).ClearCurrentValueAndRange
            crBrnchRecDis.ParameterFields(1).AddCurrentValue ParamValue1
            crBrnchRecDis.ParameterFields(2).ClearCurrentValueAndRange
            crBrnchRecDis.ParameterFields(2).AddCurrentValue ParamValue2
            crBrnchRecDis.ParameterFields(3).ClearCurrentValueAndRange
            crBrnchRecDis.ParameterFields(3).AddCurrentValue ParamValue3
            CRViewer1.ReportSource = crBrnchRecDis
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
'********************** CODE UPDATED: 08/04/2017 ******************************
        Case "SETTLEMENT DISCREPANCY"
            Dim crWHSDiscrepancy As New rptWHSDiscrepancy
            
            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crWHSDiscrepancy.DiscardSavedData
            crWHSDiscrepancy.Database.SetDataSource rsPrint
            crWHSDiscrepancy.ParameterFields(1).ClearCurrentValueAndRange
            crWHSDiscrepancy.ParameterFields(1).AddCurrentValue ParamValue1
            crWHSDiscrepancy.ParameterFields(2).ClearCurrentValueAndRange
            crWHSDiscrepancy.ParameterFields(2).AddCurrentValue ParamValue2
            crWHSDiscrepancy.ParameterFields(3).ClearCurrentValueAndRange
            crWHSDiscrepancy.ParameterFields(3).AddCurrentValue ParamValue3
            crWHSDiscrepancy.ParameterFields(4).ClearCurrentValueAndRange
            crWHSDiscrepancy.ParameterFields(4).AddCurrentValue ParamValue4
            crWHSDiscrepancy.ParameterFields(5).ClearCurrentValueAndRange
            crWHSDiscrepancy.ParameterFields(5).AddCurrentValue ParamValue5
            crWHSDiscrepancy.ParameterFields(6).ClearCurrentValueAndRange
            crWHSDiscrepancy.ParameterFields(6).AddCurrentValue ParamValue6
            crWHSDiscrepancy.ParameterFields(7).ClearCurrentValueAndRange
            crWHSDiscrepancy.ParameterFields(7).AddCurrentValue ParamValue7
            crWHSDiscrepancy.ParameterFields(8).ClearCurrentValueAndRange
            crWHSDiscrepancy.ParameterFields(8).AddCurrentValue ParamValue8
            crWHSDiscrepancy.ParameterFields(9).ClearCurrentValueAndRange
            crWHSDiscrepancy.ParameterFields(9).AddCurrentValue ParamValue9
            crWHSDiscrepancy.ParameterFields(10).ClearCurrentValueAndRange
            crWHSDiscrepancy.ParameterFields(10).AddCurrentValue ParamValue10
            crWHSDiscrepancy.ParameterFields(11).ClearCurrentValueAndRange
            crWHSDiscrepancy.ParameterFields(11).AddCurrentValue ParamValue11
            crWHSDiscrepancy.ParameterFields(12).ClearCurrentValueAndRange
            crWHSDiscrepancy.ParameterFields(12).AddCurrentValue ParamValue12
            crWHSDiscrepancy.ParameterFields(13).ClearCurrentValueAndRange
            crWHSDiscrepancy.ParameterFields(13).AddCurrentValue ParamValue13
            crWHSDiscrepancy.ParameterFields(14).ClearCurrentValueAndRange
            crWHSDiscrepancy.ParameterFields(14).AddCurrentValue ParamValue14
            crWHSDiscrepancy.ParameterFields(15).ClearCurrentValueAndRange
            crWHSDiscrepancy.ParameterFields(15).AddCurrentValue ParamValue15

            CRViewer1.ReportSource = crWHSDiscrepancy
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
'****************** END CODE UPDATED: 08/04/2017 ******************************
'********************** CODE UPDATED: 03/20/2018 ******************************
        Case "Branch Purchase Order"
            Dim crPOSummary As New rptPOSummary
    
            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crPOSummary.DiscardSavedData
            crPOSummary.Database.SetDataSource rsPrint
            crPOSummary.ParameterFields(1).ClearCurrentValueAndRange
            crPOSummary.ParameterFields(1).AddCurrentValue ParamValue1
            crPOSummary.ParameterFields(2).ClearCurrentValueAndRange
            crPOSummary.ParameterFields(2).AddCurrentValue ParamValue2
            crPOSummary.ParameterFields(3).ClearCurrentValueAndRange
            crPOSummary.ParameterFields(3).AddCurrentValue ParamValue3

            CRViewer1.ReportSource = crPOSummary
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
'****************** END CODE UPDATED: 03/20/2018 ******************************
'****************** CODE UPDATED: 06/19/17 *****************
        Case "Parts Catalog"
            Dim crPartsCatalog As New rptPartsCatalog
            
            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn, conPartsCat, adOpenDynamic, adLockOptimistic
            crPartsCatalog.DiscardSavedData
            crPartsCatalog.Database.SetDataSource rsPrint
            crPartsCatalog.ParameterFields(1).ClearCurrentValueAndRange
            crPartsCatalog.ParameterFields(1).AddCurrentValue ParamValue1
            crPartsCatalog.ParameterFields(2).ClearCurrentValueAndRange
            crPartsCatalog.ParameterFields(2).AddCurrentValue ParamValue2
            crPartsCatalog.ParameterFields(3).ClearCurrentValueAndRange
            crPartsCatalog.ParameterFields(3).AddCurrentValue ParamValue3
            crPartsCatalog.ParameterFields(4).ClearCurrentValueAndRange
            crPartsCatalog.ParameterFields(4).AddCurrentValue ParamValue4
            crPartsCatalog.ParameterFields(5).ClearCurrentValueAndRange
            crPartsCatalog.ParameterFields(5).AddCurrentValue ParamValue5
            crPartsCatalog.ParameterFields(6).ClearCurrentValueAndRange
            crPartsCatalog.ParameterFields(6).AddCurrentValue ParamValue6
            crPartsCatalog.ParameterFields(7).ClearCurrentValueAndRange
            crPartsCatalog.ParameterFields(7).AddCurrentValue ParamValue7
            
            CRViewer1.ReportSource = crPartsCatalog
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
        Case "Parts Catalog Summary"
            Dim crPartsCatalogSumm As New partscatalogsummary
            
            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn, conPartsCat, adOpenDynamic, adLockOptimistic
            crPartsCatalogSumm.DiscardSavedData
            crPartsCatalogSumm.Database.SetDataSource rsPrint
            crPartsCatalogSumm.ParameterFields(1).ClearCurrentValueAndRange
            crPartsCatalogSumm.ParameterFields(1).AddCurrentValue ParamValue1
            
            CRViewer1.ReportSource = crPartsCatalogSumm
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
        Case "RFP"
            Dim crRFP As New rptRFPWH
            
            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            With crRFP
                .DiscardSavedData
                 DoEvents
                .Database.SetDataSource rsPrint
                .ParameterFields(1).ClearCurrentValueAndRange
                .ParameterFields(1).AddCurrentValue ParamValue1
                .ParameterFields(2).ClearCurrentValueAndRange
                .ParameterFields(2).AddCurrentValue ParamValue2
                .ParameterFields(3).ClearCurrentValueAndRange
                .ParameterFields(3).AddCurrentValue ParamValue3
                .ParameterFields(4).ClearCurrentValueAndRange
                .ParameterFields(4).AddCurrentValue ParamValue4
                .ParameterFields(5).ClearCurrentValueAndRange
                .ParameterFields(5).AddCurrentValue ParamValue5
                .ParameterFields(6).ClearCurrentValueAndRange
                .ParameterFields(6).AddCurrentValue ParamValue7
                .ParameterFields(7).ClearCurrentValueAndRange
                .ParameterFields(7).AddCurrentValue ParamValue8
                .ParameterFields(8).ClearCurrentValueAndRange
                .ParameterFields(8).AddCurrentValue ParamValue9
                .ParameterFields(9).ClearCurrentValueAndRange
                .ParameterFields(9).AddCurrentValue ParamValue10
                .ReadRecords
            End With

            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn2, LYRcon, adOpenDynamic, adLockOptimistic
            With crRFP.Subreport1.OpenSubreport
                .DiscardSavedData
                 DoEvents
                .Database.SetDataSource rsPrint
                .ReadRecords
            End With
            
            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn3, LYRcon, adOpenDynamic, adLockOptimistic
            With crRFP.Subreport2.OpenSubreport
                .DiscardSavedData
                 DoEvents
                .Database.SetDataSource rsPrint
                .ReadRecords
            End With
            
            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn4, LYRcon, adOpenDynamic, adLockOptimistic
            With crRFP.Subreport3.OpenSubreport
                .DiscardSavedData
                 DoEvents
                .Database.SetDataSource rsPrint
                .ReadRecords
            End With
            
            CRViewer1.ReportSource = crRFP
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
        Case "RFP CBD"
            Dim crRFPCBD As New rptRFPWHCBD
            
            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            With crRFPCBD
                .DiscardSavedData
                 DoEvents
                .Database.SetDataSource rsPrint
                .ParameterFields(1).ClearCurrentValueAndRange
                .ParameterFields(1).AddCurrentValue ParamValue1
                .ParameterFields(2).ClearCurrentValueAndRange
                .ParameterFields(2).AddCurrentValue ParamValue2
                .ParameterFields(3).ClearCurrentValueAndRange
                .ParameterFields(3).AddCurrentValue ParamValue3
                .ParameterFields(4).ClearCurrentValueAndRange
                .ParameterFields(4).AddCurrentValue ParamValue4
                .ParameterFields(5).ClearCurrentValueAndRange
                .ParameterFields(5).AddCurrentValue ParamValue5
                '.ParameterFields(6).ClearCurrentValueAndRange
                '.ParameterFields(6).AddCurrentValue ParamValue6
                .ReadRecords
            End With

            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn2, LYRcon, adOpenDynamic, adLockOptimistic
            With crRFPCBD.Subreport1.OpenSubreport
                .DiscardSavedData
                 DoEvents
                .Database.SetDataSource rsPrint
                .ReadRecords
            End With
            
            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn3, LYRcon, adOpenDynamic, adLockOptimistic
            With crRFPCBD.Subreport2.OpenSubreport
                .DiscardSavedData
                 DoEvents
                .Database.SetDataSource rsPrint
                .ReadRecords
            End With
            
            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn4, LYRcon, adOpenDynamic, adLockOptimistic
            With crRFPCBD.Subreport3.OpenSubreport
                .DiscardSavedData
                 DoEvents
                .Database.SetDataSource rsPrint
                .ReadRecords
            End With
            
            CRViewer1.ReportSource = crRFPCBD
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
        Case "RFP RECONCILIATION"
            Dim rptReconSumm As New rptRFPReconSummary
            
            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            rptReconSumm.DiscardSavedData
            rptReconSumm.Database.SetDataSource rsPrint
            rptReconSumm.ParameterFields(1).ClearCurrentValueAndRange
            rptReconSumm.ParameterFields(1).AddCurrentValue ParamValue1
            rptReconSumm.ParameterFields(2).ClearCurrentValueAndRange
            rptReconSumm.ParameterFields(2).AddCurrentValue ParamValue2
            rptReconSumm.ParameterFields(3).ClearCurrentValueAndRange
            rptReconSumm.ParameterFields(3).AddCurrentValue ParamValue3
            CRViewer1.ReportSource = rptReconSumm
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
    Case "ADJUSTMENT REQUEST"
            Dim crADJUSTMENTREQ As New rptAdjustmentRequest
            
            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crADJUSTMENTREQ.DiscardSavedData
            crADJUSTMENTREQ.Database.SetDataSource rsPrint
            crADJUSTMENTREQ.ParameterFields(1).ClearCurrentValueAndRange
            crADJUSTMENTREQ.ParameterFields(1).AddCurrentValue ParamValue1
            crADJUSTMENTREQ.ParameterFields(2).ClearCurrentValueAndRange
            crADJUSTMENTREQ.ParameterFields(2).AddCurrentValue ParamValue2
            crADJUSTMENTREQ.ParameterFields(3).ClearCurrentValueAndRange
            crADJUSTMENTREQ.ParameterFields(3).AddCurrentValue ParamValue3
            crADJUSTMENTREQ.ParameterFields(4).ClearCurrentValueAndRange
            crADJUSTMENTREQ.ParameterFields(4).AddCurrentValue ParamValue4
            CRViewer1.ReportSource = crADJUSTMENTREQ
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
    Case "ADJUSTMENT REQUEST SCRIPT"
            Dim crADJUSTMENTREQSCRIPT As New rptAdjustmentRequestProcessScript
            
            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crADJUSTMENTREQSCRIPT.DiscardSavedData
            crADJUSTMENTREQSCRIPT.Database.SetDataSource rsPrint
            crADJUSTMENTREQSCRIPT.ParameterFields(1).ClearCurrentValueAndRange
            crADJUSTMENTREQSCRIPT.ParameterFields(1).AddCurrentValue ParamValue1
            crADJUSTMENTREQSCRIPT.ParameterFields(2).ClearCurrentValueAndRange
            crADJUSTMENTREQSCRIPT.ParameterFields(2).AddCurrentValue ParamValue2
            crADJUSTMENTREQSCRIPT.ParameterFields(3).ClearCurrentValueAndRange
            crADJUSTMENTREQSCRIPT.ParameterFields(3).AddCurrentValue ParamValue3
            crADJUSTMENTREQSCRIPT.ParameterFields(4).ClearCurrentValueAndRange
            crADJUSTMENTREQSCRIPT.ParameterFields(4).AddCurrentValue ParamValue4
            CRViewer1.ReportSource = crADJUSTMENTREQSCRIPT
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
    Case "CANCELLATION MODULE"
            Dim crCM As New rptCancellationModule
            
            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crCM.DiscardSavedData
            crCM.Database.SetDataSource rsPrint
            crCM.ParameterFields(1).ClearCurrentValueAndRange
            crCM.ParameterFields(1).AddCurrentValue ParamValue1
            crCM.ParameterFields(2).ClearCurrentValueAndRange
            crCM.ParameterFields(2).AddCurrentValue ParamValue2
            crCM.ParameterFields(3).ClearCurrentValueAndRange
            crCM.ParameterFields(3).AddCurrentValue ParamValue3
            crCM.ParameterFields(4).ClearCurrentValueAndRange
            crCM.ParameterFields(4).AddCurrentValue ParamValue4
            CRViewer1.ReportSource = crCM
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
    End Select
'************** END CODE UPDATED: 06/19/17 *****************
    
        Me.Width = Warehouse.Width - (Warehouse.Width * 0.25)
        Me.Height = Warehouse.Height - (Warehouse.Height * 0.3)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MySource = "" Then
        findBBO.Show
    ElseIf MySource = "BRANCH RECEIVING DISCREPANCY PRINT" Then
        Exit Sub
    Else
        Exit Sub
    End If
End Sub
