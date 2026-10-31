VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Begin VB.Form frmPrn 
   Caption         =   "Print Preview"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8505
   ScaleWidth      =   11985
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Tag             =   "xyf"
      Top             =   480
      Width           =   11775
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
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   960
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   12105
      m_MDI_Height    =   9015
   End
End
Attribute VB_Name = "frmPrn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M_Proj As CRAXDRT.Application
Dim CF_JOB As CRAXDRT.Report

Dim rsPrn As New ADODB.Recordset
Public QryPrn As String
Public MySource As String
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
Public ParamValue11 As String
Public ParamValue12 As String
Public ParamValue13 As String
Public ParamValue14 As String
Public ParamValue15 As String
Public ParamValue16 As String
Public ParamValue17 As String
Public ParamValue18 As String
Public ParamValue19 As String
Public ParamValue20 As String
Public ParamValue21 As String
Public ParamValue22 As String
Public ParamValue23 As String
Public ParamValue24 As String
Public ParamValue25 As String
Public ParamValue26 As String
Public ParamValue27 As String
Public ParamValue28 As String
Public ParamValue29 As String
Public ParamValue30 As String
Public ParamValue31 As String
Public ParamValue32 As String
Public ParamValue33 As String



Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
    If MySource = "RECEIVING" Then
        If Val(Receiving.txtPrnCopy.Text) = 0 Then
            Call Receiving.UpdateItemReceived
        Else
            Receiving.txtPrnCopy.Text = Val(Receiving.txtPrnCopy.Text) + 1
            LYRcon.Execute "UPDATE invreceiving SET PRNCOPY = invreceiving.PRNCOPY + " & 1 & " Where invreceiving.RRECID = " & Val(Receiving.txtRecID.Text) & ";"
        End If
    End If
    If MySource = "OUTGOING" Then
        If Val(Outgoing.txtPrnCopy.Text) = 0 Then
            Call Outgoing.UpdateItemDelivered
        Else
            Outgoing.txtPrnCopy.Text = Val(Outgoing.txtPrnCopy.Text) + 1
            LYRcon.Execute "UPDATE invoutgoing SET PRNCOPY = invoutgoing.PRNCOPY + " & 1 & " Where invoutgoing.DRID = " & Val(Outgoing.txtDRid.Text) & ";"
        End If
    End If
    
End Sub

Private Sub Form_Load()
    
    Select Case MySource
    
        Case "TNRPERAGENT"
'            Dim crTNR As New rptTNRind
            
            crTNR.DiscardSavedData
            'crTNR.Database.SetDataSource rsPrn
            crTNR.ParameterFields(1).ClearCurrentValueAndRange
            crTNR.ParameterFields(1).AddCurrentValue ParamValue1
            crTNR.ParameterFields(2).ClearCurrentValueAndRange
            crTNR.ParameterFields(2).AddCurrentValue ParamValue2
            crTNR.ParameterFields(3).ClearCurrentValueAndRange
            crTNR.ParameterFields(3).AddCurrentValue ParamValue3
            crTNR.ParameterFields(4).ClearCurrentValueAndRange
            crTNR.ParameterFields(4).AddCurrentValue ParamValue4
            crTNR.ParameterFields(5).ClearCurrentValueAndRange
            crTNR.ParameterFields(5).AddCurrentValue ParamValue5
            crTNR.ParameterFields(6).ClearCurrentValueAndRange
            crTNR.ParameterFields(6).AddCurrentValue ParamValue6
            crTNR.ParameterFields(7).ClearCurrentValueAndRange
            crTNR.ParameterFields(7).AddCurrentValue ParamValue7
            crTNR.ParameterFields(8).ClearCurrentValueAndRange
            crTNR.ParameterFields(8).AddCurrentValue ParamValue8
            crTNR.ParameterFields(9).ClearCurrentValueAndRange
            crTNR.ParameterFields(9).AddCurrentValue ParamValue9
            crTNR.ParameterFields(10).ClearCurrentValueAndRange
            crTNR.ParameterFields(10).AddCurrentValue ParamValue10
            crTNR.ParameterFields(11).ClearCurrentValueAndRange
            crTNR.ParameterFields(11).AddCurrentValue ParamValue11
            crTNR.ParameterFields(12).ClearCurrentValueAndRange
            crTNR.ParameterFields(12).AddCurrentValue ParamValue12
            crTNR.ParameterFields(13).ClearCurrentValueAndRange
            crTNR.ParameterFields(13).AddCurrentValue ParamValue13
            crTNR.ParameterFields(14).ClearCurrentValueAndRange
            crTNR.ParameterFields(14).AddCurrentValue ParamValue14
            crTNR.ParameterFields(15).ClearCurrentValueAndRange
            crTNR.ParameterFields(15).AddCurrentValue ParamValue15
            crTNR.ParameterFields(16).ClearCurrentValueAndRange
            crTNR.ParameterFields(16).AddCurrentValue ParamValue16
            crTNR.ParameterFields(17).ClearCurrentValueAndRange
            crTNR.ParameterFields(17).AddCurrentValue ParamValue17
            crTNR.ParameterFields(18).ClearCurrentValueAndRange
            crTNR.ParameterFields(18).AddCurrentValue ParamValue18
            crTNR.ParameterFields(19).ClearCurrentValueAndRange
            crTNR.ParameterFields(19).AddCurrentValue ParamValue19
            crTNR.ParameterFields(20).ClearCurrentValueAndRange
            crTNR.ParameterFields(20).AddCurrentValue ParamValue20
            crTNR.ParameterFields(21).ClearCurrentValueAndRange
            crTNR.ParameterFields(21).AddCurrentValue ParamValue21
            crTNR.ParameterFields(22).ClearCurrentValueAndRange
            crTNR.ParameterFields(22).AddCurrentValue ParamValue22
            crTNR.ParameterFields(23).ClearCurrentValueAndRange
            crTNR.ParameterFields(23).AddCurrentValue ParamValue23
            crTNR.ParameterFields(24).ClearCurrentValueAndRange
            crTNR.ParameterFields(24).AddCurrentValue ParamValue24
            crTNR.ParameterFields(25).ClearCurrentValueAndRange
            crTNR.ParameterFields(25).AddCurrentValue ParamValue25
            CRViewer1.ReportSource = crTNR
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True

        Case "RECEIVING"
            'Dim crREC As New rptReceiving
            
            If rsPrn.State = adStateOpen Then rsPrn.Close
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crREC.DiscardSavedData
            crREC.Database.SetDataSource rsPrn
            crREC.ParameterFields(1).ClearCurrentValueAndRange
            crREC.ParameterFields(1).AddCurrentValue ParamValue1
            crREC.ParameterFields(2).ClearCurrentValueAndRange
            crREC.ParameterFields(2).AddCurrentValue ParamValue2
            crREC.ParameterFields(3).ClearCurrentValueAndRange
            crREC.ParameterFields(3).AddCurrentValue ParamValue3
            crREC.ParameterFields(4).ClearCurrentValueAndRange
            crREC.ParameterFields(4).AddCurrentValue ParamValue4
            crREC.ParameterFields(5).ClearCurrentValueAndRange
            crREC.ParameterFields(5).AddCurrentValue ParamValue5
            crREC.ParameterFields(6).ClearCurrentValueAndRange
            crREC.ParameterFields(6).AddCurrentValue ParamValue6
            crREC.ParameterFields(7).ClearCurrentValueAndRange
            crREC.ParameterFields(7).AddCurrentValue ParamValue7
            crREC.ParameterFields(8).ClearCurrentValueAndRange
            crREC.ParameterFields(8).AddCurrentValue ParamValue8
            crREC.ParameterFields(9).ClearCurrentValueAndRange
            crREC.ParameterFields(9).AddCurrentValue ParamValue9
            CRViewer1.ReportSource = crREC
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
        
        Case "DETAILED RECEIVING INV"
            'Dim crDtlRec As New RptRecDtls
            
            If rsPrn.State = adStateOpen Then rsPrn.Close
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crDtlRec.DiscardSavedData
            crDtlRec.Database.SetDataSource rsPrn
            crDtlRec.ParameterFields(1).ClearCurrentValueAndRange
            crDtlRec.ParameterFields(1).AddCurrentValue ParamValue1
            crDtlRec.ParameterFields(2).ClearCurrentValueAndRange
            crDtlRec.ParameterFields(2).AddCurrentValue ParamValue2
            crDtlRec.ParameterFields(3).ClearCurrentValueAndRange
            crDtlRec.ParameterFields(3).AddCurrentValue ParamValue3
            CRViewer1.ReportSource = crDtlRec
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
            
        Case "OUTGOING"
            'Dim crOUT As New rptOutgoing
            
            If rsPrn.State = adStateOpen Then rsPrn.Close
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crREC.DiscardSavedData
            crOUT.Database.SetDataSource rsPrn
            crOUT.ParameterFields(1).ClearCurrentValueAndRange
            crOUT.ParameterFields(1).AddCurrentValue ParamValue1
            crOUT.ParameterFields(2).ClearCurrentValueAndRange
            crOUT.ParameterFields(2).AddCurrentValue ParamValue2
            crOUT.ParameterFields(3).ClearCurrentValueAndRange
            crOUT.ParameterFields(3).AddCurrentValue ParamValue3
            crOUT.ParameterFields(4).ClearCurrentValueAndRange
            crOUT.ParameterFields(4).AddCurrentValue ParamValue4
            crOUT.ParameterFields(5).ClearCurrentValueAndRange
            crOUT.ParameterFields(5).AddCurrentValue ParamValue5
            crOUT.ParameterFields(6).ClearCurrentValueAndRange
            crOUT.ParameterFields(6).AddCurrentValue ParamValue6
            crOUT.ParameterFields(7).ClearCurrentValueAndRange
            crOUT.ParameterFields(7).AddCurrentValue ParamValue7
            crOUT.ParameterFields(8).ClearCurrentValueAndRange
            crOUT.ParameterFields(8).AddCurrentValue ParamValue8
            crOUT.ParameterFields(9).ClearCurrentValueAndRange
            crOUT.ParameterFields(9).AddCurrentValue ParamValue9
            crOUT.ParameterFields(10).ClearCurrentValueAndRange
            crOUT.ParameterFields(10).AddCurrentValue ParamValue10
            crOUT.ParameterFields(11).ClearCurrentValueAndRange
            crOUT.ParameterFields(11).AddCurrentValue ParamValue11
            CRViewer1.ReportSource = crOUT
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
        
        Case "INVENTORY"
            'Dim crInv As New RptInventory
        
            
            If rsPrn.State = adStateOpen Then rsPrn.Close
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crInv.DiscardSavedData
            crInv.Database.SetDataSource rsPrn
            crInv.ParameterFields(1).ClearCurrentValueAndRange
            crInv.ParameterFields(1).AddCurrentValue ParamValue1
            crInv.ParameterFields(2).ClearCurrentValueAndRange
            crInv.ParameterFields(2).AddCurrentValue ParamValue2
            crInv.ParameterFields(3).ClearCurrentValueAndRange
            crInv.ParameterFields(3).AddCurrentValue ParamValue3
            CRViewer1.ReportSource = crInv
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
            
        Case "DETAILED OUTGOING INV"
            'Dim crDtlOut As New RptOUTdtls
            
            If rsPrn.State = adStateOpen Then rsPrn.Close
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crDtlOut.DiscardSavedData
            crDtlOut.Database.SetDataSource rsPrn
            crDtlOut.ParameterFields(1).ClearCurrentValueAndRange
            crDtlOut.ParameterFields(1).AddCurrentValue ParamValue1
            crDtlOut.ParameterFields(2).ClearCurrentValueAndRange
            crDtlOut.ParameterFields(2).AddCurrentValue ParamValue2
            crDtlOut.ParameterFields(3).ClearCurrentValueAndRange
            crDtlOut.ParameterFields(3).AddCurrentValue ParamValue3
            CRViewer1.ReportSource = crDtlOut
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
            
        Case "Branch Purchase Order"
            Dim crDtlPO As New rptspPOdtls
            
            If rsPrn.State = adStateOpen Then rsPrn.Close
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crDtlPO.DiscardSavedData
            crDtlPO.Database.SetDataSource rsPrn
            crDtlPO.ParameterFields(1).ClearCurrentValueAndRange
            crDtlPO.ParameterFields(1).AddCurrentValue ParamValue1
            crDtlPO.ParameterFields(2).ClearCurrentValueAndRange
            crDtlPO.ParameterFields(2).AddCurrentValue ParamValue2
            crDtlPO.ParameterFields(3).ClearCurrentValueAndRange
            crDtlPO.ParameterFields(3).AddCurrentValue ParamValue3
            crDtlPO.ParameterFields(4).ClearCurrentValueAndRange
            crDtlPO.ParameterFields(4).AddCurrentValue ParamValue4
            crDtlPO.ParameterFields(5).ClearCurrentValueAndRange
            crDtlPO.ParameterFields(5).AddCurrentValue ParamValue5
            crDtlPO.ParameterFields(6).ClearCurrentValueAndRange
            crDtlPO.ParameterFields(6).AddCurrentValue ParamValue6
            crDtlPO.ParameterFields(7).ClearCurrentValueAndRange
            crDtlPO.ParameterFields(7).AddCurrentValue ParamValue7
            CRViewer1.ReportSource = crDtlPO
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
            
    End Select
    
        Me.Width = Warehouse.Width - (Warehouse.Width * 0.25)
        Me.Height = Warehouse.Height - (Warehouse.Height * 0.3)
    
End Sub
