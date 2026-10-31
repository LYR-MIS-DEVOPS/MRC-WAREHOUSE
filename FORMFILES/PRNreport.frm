VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form PRNmc 
   Caption         =   "Print Preview"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   840
   ClientWidth     =   8040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   8040
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtCriteria 
      Height          =   285
      Left            =   4080
      TabIndex        =   3
      Top             =   6960
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1560
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PRNreport.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7815
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
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
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "E&xit Report"
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
      Image           =   "PRNreport.frx":27B2
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdFilter 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Filter Records"
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
      Image           =   "PRNreport.frx":2D4C
      cBack           =   -2147483633
   End
   Begin VB.Label lblRemarks 
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   7320
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "Print"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "PRNmc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***Print Reports
Dim mcLyr As CRAXDRT.Application
Dim mcReport As CRAXDRT.Report

Dim rsPrn As New ADODB.Recordset
Public QryPrn As String
Public MySource As String
Public MyPOTYPE As String, MyRECTYPE As String
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

Private Sub cmdExit_Click()
    'Dim i As Integer, Ttl As Integer
   
    'If PRNmc.MySource = "STOCKTRANSFER" Then TranSummary.Show
        Unload Me

'    PRNmc.MySource = "STOCKTRANSFER W/ COST" Then TranSummary.Show
'    For i = 0 To Forms.Count - 1
'        Ttl = i
'    Next i
'    If Ttl < 1 Then
'        'Warehouse.tbMain.Visible = True
'    End If
End Sub

Private Sub cmdFilter_Click()
    Dim i As Integer
    
    Select Case MySource
        Case "PRICELIST"
            For i = 0 To 2
                 With SPfilter.cboLookIn(i)
                   .AddItem "PART NUMBER"
                   .AddItem "DESCRIPTION"
                   .AddItem "PRICE"
                   .AddItem "SPARE PARTS TYPE"
                   .AddItem "SUPPLIER"
                 End With
            Next i
            SPfilter.fldNames = "SPAREPARTS.PARTNUMBER*SPAREPARTS.DESCRIPTION*SPAREPARTS.STOREPRICE*SPAREPARTS.REMARKS*SPAREPARTS.SUPPLIERCODE"
            SPfilter.fldNdx = "2"
            SPfilter.fldDates = ""
            SPfilter.fldSupplier = 4
            SPfilter.MySource = "PRICELIST"
            SPfilter.Show vbModal
    End Select
End Sub

Public Sub RefreshPriceList()
    Dim crPriceList As New rptPriceList
    
    cmdFilter.Visible = True
    QryPrn = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.APPLICMODEL, SPAREPARTS.REMARKS, SPAREPARTS.SUPPLIERCODE, SPAREPARTS.DLRPRICE From SPAREPARTS " & _
             IIf(txtCriteria.Text <> "", txtCriteria.Text, "") & " ORDER BY SPAREPARTS.REMARKS, SPAREPARTS.SUPPLIERCODE, SPAREPARTS.PARTNUMBER;"
    If rsPrn.State = adStateOpen Then rsPrn.Close
    rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
    crPriceList.DiscardSavedData
    crPriceList.Database.SetDataSource rsPrn
    crPriceList.ParameterFields(1).ClearCurrentValueAndRange
    crPriceList.ParameterFields(1).AddCurrentValue lblRemarks.Caption
    crPriceList.ParameterFields(2).ClearCurrentValueAndRange
    crPriceList.ParameterFields(2).AddCurrentValue SysBranch
    'crPriceList.ParameterFields(3).ClearCurrentValueAndRange
    'crPriceList.ParameterFields(3).AddCurrentValue SysBranch
    
    CRViewer1.ReportSource = crPriceList
    CRViewer1.ViewReport
    CRViewer1.DisplayGroupTree = False
    CRViewer1.EnableExportButton = True
End Sub
Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
    Dim MyID As Integer
    
    If MySource = "AUDITINVEN" Then
        MyID = Val(AuditInvntry.txtBatchID.Text)
        '****Access
        'LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.STATUS = 'FINAL' WHERE (((AUDITINVEN.BATCHID)=" & MyId & "));"
        'mySQL
        LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.STATUS = 'FINAL' WHERE AUDITINVEN.BATCHID=" & MyID & ";"
        
        AuditInvntry.txtStatus = "FINAL"
    ElseIf MySource = "SPPO" Or MySource = "SPPOrepl" Or MySource = "SPPOreplAmt" Or MySource = "SPPOreplPnum" Or MySource = "SPPOPRICE" Then
        '***Access
        'LYRcon.Execute "UPDATE SPORDER SET SPORDER.PRINTCOPY = [SPORDER].[PRINTCOPY]+1 " & _
                       "WHERE (((SPORDER.REFERENCE)='" & SPordering.txtPONumber.Text & "'));"
        'mySQL
        If MyPOTYPE = "WAREHOUSE" Then
            LYRcon.Execute "UPDATE SPORDER SET SPORDER.PRINTCOPY = SPORDER.PRINTCOPY+1 " & _
                           "WHERE SPORDER.REFERENCE='" & SPordering.txtPONumber.Text & "';"
            
            Call SPordering.FinalSPorder
            SPordering.txtPrnCopy.Text = Val(SPordering.txtPrnCopy.Text) + 1
        Else
            LYRcon.Execute "UPDATE SPORDER SET SPORDER.PRINTCOPY = SPORDER.PRINTCOPY+1 " & _
                           "WHERE SPORDER.REFERENCE='" & SPorderingWS.txtPONumber.Text & "';"
            
            Call SPorderingWS.FinalSPorder
            SPorderingWS.txtPrnCopy.Text = Val(SPorderingWS.txtPrnCopy.Text) + 1
        End If
    ElseIf MySource = "SERVICEDAILYINCOME" Then
        SRVCincome.txtPrintCopy.Text = Val(SRVCincome.txtPrintCopy.Text) + 1
        '***Access
        'LYRcon.Execute "UPDATE SERVICEINCOME SET SERVICEINCOME.PRINTCOPY = [SERVICEINCOME].[PRINTCOPY]+1 " & _
                       "WHERE (((SERVICEINCOME.SRVCINCOMEID)=" & SRVCincome.txtSrvcID.Text & "));"
        'mySQl
        LYRcon.Execute "UPDATE SERVICEINCOME SET SERVICEINCOME.PRINTCOPY = SERVICEINCOME.PRINTCOPY+1 " & _
                       "WHERE SERVICEINCOME.SRVCINCOMEID=" & SRVCincome.txtSrvcID.Text & ";"
        
    ElseIf MySource = "SERVICEDAILYINCOMENEW" Then
        SRVCincomeNew.txtPrintCopy.Text = Val(SRVCincomeNew.txtPrintCopy.Text) + 1
        '***Access
        'LYRcon.Execute "UPDATE SERVICEINCOME SET SERVICEINCOME.PRINTCOPY = [SERVICEINCOME].[PRINTCOPY]+1 " & _
                       "WHERE (((SERVICEINCOME.SRVCINCOMEID)=" & SRVCincomeNew.txtSrvcID.Text & "));"
        'mySQl
        LYRcon.Execute "UPDATE SERVICEINCOME SET SERVICEINCOME.PRINTCOPY = SERVICEINCOME.PRINTCOPY+1 " & _
                       "WHERE SERVICEINCOME.SRVCINCOMEID=" & SRVCincomeNew.txtSrvcID.Text & ";"
    ElseIf MySource = "SPRECEIVE" Then
        If MyRECTYPE = "WAREHOUSE" Then
            Call SPincoming.FinalSPrec
        Else
            Call SPincomingWS.FinalSPrec
        End If
    ElseIf MySource = "SPBPO" Then
        LYRcon.Execute "UPDATE spordertaking SET PRINTPICK = spordertaking.PRINTPICK+1 where spordertaking.BRANCHID = " & Val(BranchPO.lblBID.Caption) & " AND spordertaking.SQSERVE = " & Val(BranchPO.lblSQSERVE) & ";"
            BranchPO.mnuconfirm.Enabled = True
    ElseIf MySource = "SPBPOACONFIRM" Then
        'If BranchPO.txtDateCon.Text = "" Then
        '    LYRcon.Execute "UPDATE spordertaking SET DATECONFIRM = '" & Format(Date, "yyyy-mm-dd") & "',FINALQT= spordertaking.FINALQT+1 " & _
        '        "where spordertaking.BRANCHID = " & Val(BranchPO.lblBID.Caption) & " AND spordertaking.SQSERVE = " & Val(BranchPO.lblSQSERVE) & ";"
        '    BranchPO.txtDateCon.Text = Format(Date, "dd/mm/yyyy")
        'End If
    ElseIf MySource = "Back Order Adj." Then
        If Val(InvntryADJ.txtPrnCopy.Text) = 0 Then
            InvntryADJ.txtPrnCopy.Text = Val(InvntryADJ.txtPrnCopy.Text) + 1
            Call InvntryADJ.UpdateSPITR
        Else
            InvntryADJ.txtPrnCopy.Text = Val(InvntryADJ.txtPrnCopy.Text) + 1
        End If
        
    End If
End Sub
Private Sub Form_Load()
    'On Error Resume Next
    
    Select Case MySource
        
        Case "SALESRETURNSUMMARY"
            Dim crSalRet As New rptSalRetSum
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crSalRet.DiscardSavedData
            crSalRet.Database.SetDataSource rsPrn
            crSalRet.ParameterFields(1).ClearCurrentValueAndRange
            crSalRet.ParameterFields(1).AddCurrentValue ParamValue1
            crSalRet.ParameterFields(2).ClearCurrentValueAndRange
            crSalRet.ParameterFields(2).AddCurrentValue ParamValue2
            crSalRet.ParameterFields(3).ClearCurrentValueAndRange
            crSalRet.ParameterFields(3).AddCurrentValue ParamValue3
            crSalRet.ParameterFields(4).ClearCurrentValueAndRange
            crSalRet.ParameterFields(4).AddCurrentValue ParamValue4
            crSalRet.ParameterFields(5).ClearCurrentValueAndRange
            crSalRet.ParameterFields(5).AddCurrentValue ParamValue5
'            crSalRet.ParameterFields(6).ClearCurrentValueAndRange
'            crSalRet.ParameterFields(6).AddCurrentValue SysBranch
'            crSalRet.ParameterFields(7).ClearCurrentValueAndRange
'            crSalRet.ParameterFields(7).AddCurrentValue SysAddress
            CRViewer1.ReportSource = crSalRet
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            'CRViewer1.EnableExportButton = False
            CRViewer1.EnableExportButton = True
            
            
        Case "AUDINVMGR" '
'            Dim crAudInvMgr As New RptAudInvMgr
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crAudInvMgr.DiscardSavedData
            crAudInvMgr.Database.SetDataSource rsPrn
            crAudInvMgr.ParameterFields(1).ClearCurrentValueAndRange
            crAudInvMgr.ParameterFields(1).AddCurrentValue ParamValue1
            crAudInvMgr.ParameterFields(2).ClearCurrentValueAndRange
            crAudInvMgr.ParameterFields(2).AddCurrentValue ParamValue2
            crAudInvMgr.ParameterFields(3).ClearCurrentValueAndRange
            crAudInvMgr.ParameterFields(3).AddCurrentValue ParamValue3
            crAudInvMgr.ParameterFields(4).ClearCurrentValueAndRange
            crAudInvMgr.ParameterFields(4).AddCurrentValue ParamValue4
            crAudInvMgr.ParameterFields(5).ClearCurrentValueAndRange
            crAudInvMgr.ParameterFields(5).AddCurrentValue SysBranch
            CRViewer1.ReportSource = crAudInvMgr
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            'CRViewer1.EnableExportButton = False
            CRViewer1.EnableExportButton = True
        Case "AUDINVMGRCOST"
            Dim crAudCost As New RptAudCost
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crAudCost.DiscardSavedData
            crAudCost.Database.SetDataSource rsPrn
            crAudCost.ParameterFields(1).ClearCurrentValueAndRange
            crAudCost.ParameterFields(1).AddCurrentValue ParamValue1
            crAudCost.ParameterFields(2).ClearCurrentValueAndRange
            crAudCost.ParameterFields(2).AddCurrentValue ParamValue2
            crAudCost.ParameterFields(3).ClearCurrentValueAndRange
            crAudCost.ParameterFields(3).AddCurrentValue ParamValue3
            crAudCost.ParameterFields(4).ClearCurrentValueAndRange
            crAudCost.ParameterFields(4).AddCurrentValue ParamValue4
            crAudCost.ParameterFields(5).ClearCurrentValueAndRange
            crAudCost.ParameterFields(5).AddCurrentValue SysBranch
            CRViewer1.ReportSource = crAudCost
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            'CRViewer1.EnableExportButton = False
            CRViewer1.EnableExportButton = True
        Case "AUDITINVEN"
            Dim crAudInv As New RptAuditInven
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crAudInv.DiscardSavedData
            crAudInv.Database.SetDataSource rsPrn
            crAudInv.ParameterFields(1).ClearCurrentValueAndRange
            crAudInv.ParameterFields(1).AddCurrentValue ParamValue1
            crAudInv.ParameterFields(2).ClearCurrentValueAndRange
            crAudInv.ParameterFields(2).AddCurrentValue ParamValue2
            crAudInv.ParameterFields(3).ClearCurrentValueAndRange
            crAudInv.ParameterFields(3).AddCurrentValue ParamValue3
            crAudInv.ParameterFields(4).ClearCurrentValueAndRange
            crAudInv.ParameterFields(4).AddCurrentValue SysBranch
            crAudInv.ParameterFields(5).ClearCurrentValueAndRange
            crAudInv.ParameterFields(5).AddCurrentValue SysAddress
            CRViewer1.ReportSource = crAudInv
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            'CRViewer1.EnableExportButton = False
            CRViewer1.EnableExportButton = True
        Case "BACKORDERPO"
            Dim crBckOrd As New rptBackOrder
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crBckOrd.DiscardSavedData
            crBckOrd.Database.SetDataSource rsPrn
            crBckOrd.ParameterFields(1).ClearCurrentValueAndRange
            crBckOrd.ParameterFields(1).AddCurrentValue ParamValue1
            crBckOrd.ParameterFields(2).ClearCurrentValueAndRange
            crBckOrd.ParameterFields(2).AddCurrentValue ParamValue2
            crBckOrd.ParameterFields(3).ClearCurrentValueAndRange
            crBckOrd.ParameterFields(3).AddCurrentValue ParamValue3
            crBckOrd.ParameterFields(4).ClearCurrentValueAndRange
            crBckOrd.ParameterFields(4).AddCurrentValue ParamValue4
            crBckOrd.ParameterFields(5).ClearCurrentValueAndRange
            crBckOrd.ParameterFields(5).AddCurrentValue ParamValue5
            crBckOrd.ParameterFields(6).ClearCurrentValueAndRange
            crBckOrd.ParameterFields(6).AddCurrentValue ParamValue6
            crBckOrd.ParameterFields(7).ClearCurrentValueAndRange
            crBckOrd.ParameterFields(7).AddCurrentValue ParamValue7
            crBckOrd.ParameterFields(8).ClearCurrentValueAndRange
            crBckOrd.ParameterFields(8).AddCurrentValue SysBranch
            crBckOrd.ParameterFields(9).ClearCurrentValueAndRange
            crBckOrd.ParameterFields(9).AddCurrentValue SysAddress
            crBckOrd.ParameterFields(10).ClearCurrentValueAndRange
            crBckOrd.ParameterFields(10).AddCurrentValue SysTel
            CRViewer1.ReportSource = crBckOrd
            CRViewer1.ViewReport
            'CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
        
        Case "SPRECEIVESUMM"
           Dim crSPrecSum As New rptRecSumm
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crSPrecSum.DiscardSavedData
            crSPrecSum.Database.SetDataSource rsPrn
            crSPrecSum.ParameterFields(1).ClearCurrentValueAndRange
            crSPrecSum.ParameterFields(1).AddCurrentValue ParamValue1
            crSPrecSum.ParameterFields(2).ClearCurrentValueAndRange
            crSPrecSum.ParameterFields(2).AddCurrentValue SysBranch
            crSPrecSum.ParameterFields(3).ClearCurrentValueAndRange
            crSPrecSum.ParameterFields(3).AddCurrentValue SysAddress
            crSPrecSum.ParameterFields(4).ClearCurrentValueAndRange
            crSPrecSum.ParameterFields(4).AddCurrentValue SysTel
            
            CRViewer1.ReportSource = crSPrecSum
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
        
        Case "MONTHLYSALES"
'            Dim crMonthSal As New rptMonthSales
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crMonthSal.DiscardSavedData
            crMonthSal.Database.SetDataSource rsPrn
            crMonthSal.ParameterFields(1).ClearCurrentValueAndRange
            crMonthSal.ParameterFields(1).AddCurrentValue ParamValue1
            crMonthSal.ParameterFields(2).ClearCurrentValueAndRange
            crMonthSal.ParameterFields(2).AddCurrentValue ParamValue2
            crMonthSal.ParameterFields(3).ClearCurrentValueAndRange
            crMonthSal.ParameterFields(3).AddCurrentValue ParamValue3
            crMonthSal.ParameterFields(4).ClearCurrentValueAndRange
            crMonthSal.ParameterFields(4).AddCurrentValue ParamValue4
            crMonthSal.ParameterFields(5).ClearCurrentValueAndRange
            crMonthSal.ParameterFields(5).AddCurrentValue ParamValue5
            crMonthSal.ParameterFields(6).ClearCurrentValueAndRange
            crMonthSal.ParameterFields(6).AddCurrentValue SysBranch
            crMonthSal.ParameterFields(7).ClearCurrentValueAndRange
            crMonthSal.ParameterFields(7).AddCurrentValue SysAddress
            crMonthSal.ParameterFields(8).ClearCurrentValueAndRange
            crMonthSal.ParameterFields(8).AddCurrentValue SysTel
            CRViewer1.ReportSource = crMonthSal
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
        
        Case "MONTHLYCOLLECTION"
'            Dim crMonthCol As New rptMonthColl
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crMonthCol.DiscardSavedData
            crMonthCol.Database.SetDataSource rsPrn
            crMonthCol.ParameterFields(1).ClearCurrentValueAndRange
            crMonthCol.ParameterFields(1).AddCurrentValue ParamValue1
            crMonthCol.ParameterFields(2).ClearCurrentValueAndRange
            crMonthCol.ParameterFields(2).AddCurrentValue ParamValue2
            crMonthCol.ParameterFields(3).ClearCurrentValueAndRange
            crMonthCol.ParameterFields(3).AddCurrentValue ParamValue3
            crMonthCol.ParameterFields(4).ClearCurrentValueAndRange
            crMonthCol.ParameterFields(4).AddCurrentValue ParamValue4
            crMonthCol.ParameterFields(5).ClearCurrentValueAndRange
            crMonthCol.ParameterFields(5).AddCurrentValue ParamValue5
            crMonthCol.ParameterFields(6).ClearCurrentValueAndRange
            crMonthCol.ParameterFields(6).AddCurrentValue SysBranch
            crMonthCol.ParameterFields(7).ClearCurrentValueAndRange
            crMonthCol.ParameterFields(7).AddCurrentValue SysAddress
            crMonthCol.ParameterFields(8).ClearCurrentValueAndRange
            crMonthCol.ParameterFields(8).AddCurrentValue SysTel
            CRViewer1.ReportSource = crMonthCol
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
    
        Case "ORATTACHMENT"
'            Dim crAttch As New rptORattachment
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crAttch.DiscardSavedData
            crAttch.ParameterFields(1).ClearCurrentValueAndRange
            crAttch.ParameterFields(1).AddCurrentValue SysBranch
            crAttch.ParameterFields(2).ClearCurrentValueAndRange
            crAttch.ParameterFields(2).AddCurrentValue SysAddress
            crAttch.Database.SetDataSource rsPrn
            CRViewer1.ReportSource = crAttch
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
        
        Case "CRATTACHMENT"
'            Dim crAttch2 As New CRattachment
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crAttch2.DiscardSavedData
            crAttch2.ParameterFields(1).ClearCurrentValueAndRange
            crAttch2.ParameterFields(1).AddCurrentValue ParamValue3
            crAttch2.ParameterFields(2).ClearCurrentValueAndRange
            crAttch2.ParameterFields(2).AddCurrentValue SysBranch
            crAttch2.ParameterFields(3).ClearCurrentValueAndRange
            crAttch2.ParameterFields(3).AddCurrentValue SysAddress
            crAttch2.Database.SetDataSource rsPrn
            CRViewer1.ReportSource = crAttch2
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
        
        
        Case "SPINVENTORY"
 '           Dim crSPmon As New rptInvMonitor
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crSPmon.DiscardSavedData
            crSPmon.Database.SetDataSource rsPrn
            crSPmon.ParameterFields(1).ClearCurrentValueAndRange
            crSPmon.ParameterFields(1).AddCurrentValue ParamValue1
            crSPmon.ParameterFields(2).ClearCurrentValueAndRange
            crSPmon.ParameterFields(2).AddCurrentValue ParamValue2
            crSPmon.ParameterFields(3).ClearCurrentValueAndRange
            crSPmon.ParameterFields(3).AddCurrentValue ParamValue3
            crSPmon.ParameterFields(4).ClearCurrentValueAndRange
            crSPmon.ParameterFields(4).AddCurrentValue ParamValue4
            crSPmon.ParameterFields(5).ClearCurrentValueAndRange
            crSPmon.ParameterFields(5).AddCurrentValue ParamValue5
            crSPmon.ParameterFields(6).ClearCurrentValueAndRange
            crSPmon.ParameterFields(6).AddCurrentValue SysBranch
            crSPmon.ParameterFields(7).ClearCurrentValueAndRange
            crSPmon.ParameterFields(7).AddCurrentValue SysAddress
            CRViewer1.ReportSource = crSPmon
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
        
        Case "SPINVOICE"
'            Dim crSPinv As New rptInvoice
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crSPinv.DiscardSavedData
            crSPinv.Database.SetDataSource rsPrn
            crSPinv.ParameterFields(1).ClearCurrentValueAndRange
            crSPinv.ParameterFields(1).AddCurrentValue ParamValue1
            crSPinv.ParameterFields(2).ClearCurrentValueAndRange
            crSPinv.ParameterFields(2).AddCurrentValue ParamValue2
            crSPinv.ParameterFields(3).ClearCurrentValueAndRange
            crSPinv.ParameterFields(3).AddCurrentValue ParamValue3
            crSPinv.ParameterFields(4).ClearCurrentValueAndRange
            crSPinv.ParameterFields(4).AddCurrentValue ParamValue4
            crSPinv.ParameterFields(5).ClearCurrentValueAndRange
            crSPinv.ParameterFields(5).AddCurrentValue ParamValue5
            crSPinv.ParameterFields(6).ClearCurrentValueAndRange
            crSPinv.ParameterFields(6).AddCurrentValue ParamValue6
            crSPinv.ParameterFields(7).ClearCurrentValueAndRange
            crSPinv.ParameterFields(7).AddCurrentValue SysBranch
            crSPinv.ParameterFields(8).ClearCurrentValueAndRange
            crSPinv.ParameterFields(8).AddCurrentValue SysAddress
            
            CRViewer1.ReportSource = crSPinv
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
            
        Case "JOSUMMARY"
'            Dim crJob As New rptJobSumm
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crJob.DiscardSavedData
            crJob.Database.SetDataSource rsPrn
            crJob.ParameterFields(1).ClearCurrentValueAndRange
            crJob.ParameterFields(1).AddCurrentValue ParamValue1
            crJob.ParameterFields(2).ClearCurrentValueAndRange
            crJob.ParameterFields(2).AddCurrentValue ParamValue2
            crJob.ParameterFields(3).ClearCurrentValueAndRange
            crJob.ParameterFields(3).AddCurrentValue SysBranch
            crJob.ParameterFields(4).ClearCurrentValueAndRange
            crJob.ParameterFields(4).AddCurrentValue SysAddress
            CRViewer1.ReportSource = crJob
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            'CRViewer1.EnableExportButton = False
            CRViewer1.EnableExportButton = True
        Case "PRICELIST"
'            Dim crPriceList As New rptPriceList
            
            cmdFilter.Visible = True
            QryPrn = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.APPLICMODEL, SPAREPARTS.REMARKS, SPAREPARTS.SUPPLIERCODE, SPAREPARTS.DLRPRICE " & _
                     "From SPAREPARTS ORDER BY SPAREPARTS.REMARKS, SPAREPARTS.SUPPLIERCODE, SPAREPARTS.PARTNUMBER;"
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crPriceList.DiscardSavedData
            crPriceList.Database.SetDataSource rsPrn
            crPriceList.ParameterFields(1).ClearCurrentValueAndRange
            crPriceList.ParameterFields(1).AddCurrentValue lblRemarks.Caption
            crPriceList.ParameterFields(2).ClearCurrentValueAndRange
            crPriceList.ParameterFields(2).AddCurrentValue SysBranch
            CRViewer1.ReportSource = crPriceList
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            'CRViewer1.EnableExportButton = False
            CRViewer1.EnableExportButton = True
        Case "SPSALESRPT"
'            Dim crSales As New rptDailySales
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crSales.DiscardSavedData
            crSales.Database.SetDataSource rsPrn
            crSales.ParameterFields(1).ClearCurrentValueAndRange
            crSales.ParameterFields(1).AddCurrentValue ParamValue1
            crSales.ParameterFields(2).ClearCurrentValueAndRange
            crSales.ParameterFields(2).AddCurrentValue ParamValue2
            crSales.ParameterFields(3).ClearCurrentValueAndRange
            crSales.ParameterFields(3).AddCurrentValue SysBranch
            CRViewer1.ReportSource = crSales
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
        
        
        Case "SPPO"
            Dim crSPPO As New rptSPPO
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crSPPO.DiscardSavedData
            crSPPO.Database.SetDataSource rsPrn
            crSPPO.ParameterFields(1).ClearCurrentValueAndRange
            crSPPO.ParameterFields(1).AddCurrentValue ParamValue1
            crSPPO.ParameterFields(2).ClearCurrentValueAndRange
            crSPPO.ParameterFields(2).AddCurrentValue ParamValue2
            crSPPO.ParameterFields(3).ClearCurrentValueAndRange
            crSPPO.ParameterFields(3).AddCurrentValue ParamValue3
            crSPPO.ParameterFields(4).ClearCurrentValueAndRange
            crSPPO.ParameterFields(4).AddCurrentValue ParamValue4
            crSPPO.ParameterFields(5).ClearCurrentValueAndRange
            crSPPO.ParameterFields(5).AddCurrentValue ParamValue5
            crSPPO.ParameterFields(6).ClearCurrentValueAndRange
            crSPPO.ParameterFields(6).AddCurrentValue ParamValue6
            crSPPO.ParameterFields(7).ClearCurrentValueAndRange
            crSPPO.ParameterFields(7).AddCurrentValue ParamValue7
            crSPPO.ParameterFields(8).ClearCurrentValueAndRange
            crSPPO.ParameterFields(8).AddCurrentValue ParamValue8
            crSPPO.ParameterFields(9).ClearCurrentValueAndRange
            crSPPO.ParameterFields(9).AddCurrentValue ParamValue9
            CRViewer1.ReportSource = crSPPO
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            If MyPOTYPE = "WHOLESALE" Then
                If Val(SPorderingWS.txtPrnCopy.Text) = 0 Then
                    CRViewer1.EnableExportButton = False
                Else
                    CRViewer1.EnableExportButton = True
                End If
            Else
                If Val(SPordering.txtPrnCopy.Text) = 0 Then
                    CRViewer1.EnableExportButton = False
                Else
                    CRViewer1.EnableExportButton = True
                End If
            End If
        Case "SPBPO"
'            Dim crSPBPO As New rptSPBPO
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crSPBPO.DiscardSavedData
            crSPBPO.Database.SetDataSource rsPrn
            crSPBPO.ParameterFields(1).ClearCurrentValueAndRange
            crSPBPO.ParameterFields(1).AddCurrentValue ParamValue1
            crSPBPO.ParameterFields(2).ClearCurrentValueAndRange
            crSPBPO.ParameterFields(2).AddCurrentValue ParamValue2
            crSPBPO.ParameterFields(3).ClearCurrentValueAndRange
            crSPBPO.ParameterFields(3).AddCurrentValue ParamValue3
            crSPBPO.ParameterFields(4).ClearCurrentValueAndRange
            crSPBPO.ParameterFields(4).AddCurrentValue ParamValue4
            crSPBPO.ParameterFields(5).ClearCurrentValueAndRange
            crSPBPO.ParameterFields(5).AddCurrentValue ParamValue5
            crSPBPO.ParameterFields(6).ClearCurrentValueAndRange
            crSPBPO.ParameterFields(6).AddCurrentValue ParamValue6
            crSPBPO.ParameterFields(7).ClearCurrentValueAndRange
            crSPBPO.ParameterFields(7).AddCurrentValue ParamValue7
            crSPBPO.ParameterFields(8).ClearCurrentValueAndRange
            crSPBPO.ParameterFields(8).AddCurrentValue ParamValue8
            'crSPBPO.ParameterFields(9).ClearCurrentValueAndRange
            'crSPBPO.ParameterFields(9).AddCurrentValue ParamValue9
            CRViewer1.Zoom 100
            CRViewer1.ReportSource = crSPBPO
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
       Case "SPBPOACONFIRM"
'            Dim crSPBPOACONFIRM As New rptSPBPOACONFIRM
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crSPBPOACONFIRM.DiscardSavedData
            crSPBPOACONFIRM.Database.SetDataSource rsPrn
            crSPBPOACONFIRM.ParameterFields(1).ClearCurrentValueAndRange
            crSPBPOACONFIRM.ParameterFields(1).AddCurrentValue ParamValue1
            crSPBPOACONFIRM.ParameterFields(2).ClearCurrentValueAndRange
            crSPBPOACONFIRM.ParameterFields(2).AddCurrentValue ParamValue2
            crSPBPOACONFIRM.ParameterFields(3).ClearCurrentValueAndRange
            crSPBPOACONFIRM.ParameterFields(3).AddCurrentValue ParamValue3
            crSPBPOACONFIRM.ParameterFields(4).ClearCurrentValueAndRange
            crSPBPOACONFIRM.ParameterFields(4).AddCurrentValue ParamValue4
            crSPBPOACONFIRM.ParameterFields(5).ClearCurrentValueAndRange
            crSPBPOACONFIRM.ParameterFields(5).AddCurrentValue ParamValue5
            crSPBPOACONFIRM.ParameterFields(6).ClearCurrentValueAndRange
            crSPBPOACONFIRM.ParameterFields(6).AddCurrentValue ParamValue6
            crSPBPOACONFIRM.ParameterFields(7).ClearCurrentValueAndRange
            crSPBPOACONFIRM.ParameterFields(7).AddCurrentValue ParamValue7
            crSPBPOACONFIRM.ParameterFields(8).ClearCurrentValueAndRange
            crSPBPOACONFIRM.ParameterFields(8).AddCurrentValue ParamValue8
            crSPBPOACONFIRM.ParameterFields(9).ClearCurrentValueAndRange
            crSPBPOACONFIRM.ParameterFields(9).AddCurrentValue ParamValue9
            crSPBPOACONFIRM.ParameterFields(10).ClearCurrentValueAndRange
            crSPBPOACONFIRM.ParameterFields(10).AddCurrentValue ParamValue10
            'crSPBPOACONFIRM.ParameterFields(11).ClearCurrentValueAndRange
            'crSPBPOACONFIRM.ParameterFields(11).AddCurrentValue ParamValue11
            CRViewer1.Zoom 100
            CRViewer1.ReportSource = crSPBPOACONFIRM
            CRViewer1.ViewReport
            
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
            
        Case "SPPOPRICE"
            Dim crPOPrc As New rptSPPOprice
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crPOPrc.DiscardSavedData
            crPOPrc.Database.SetDataSource rsPrn
            crPOPrc.ParameterFields(1).ClearCurrentValueAndRange
            crPOPrc.ParameterFields(1).AddCurrentValue ParamValue1
            crPOPrc.ParameterFields(2).ClearCurrentValueAndRange
            crPOPrc.ParameterFields(2).AddCurrentValue ParamValue2
            crPOPrc.ParameterFields(3).ClearCurrentValueAndRange
            crPOPrc.ParameterFields(3).AddCurrentValue ParamValue3
            crPOPrc.ParameterFields(4).ClearCurrentValueAndRange
            crPOPrc.ParameterFields(4).AddCurrentValue ParamValue4
            crPOPrc.ParameterFields(5).ClearCurrentValueAndRange
            crPOPrc.ParameterFields(5).AddCurrentValue ParamValue5
            crPOPrc.ParameterFields(6).ClearCurrentValueAndRange
            crPOPrc.ParameterFields(6).AddCurrentValue ParamValue6
            crPOPrc.ParameterFields(7).ClearCurrentValueAndRange
            crPOPrc.ParameterFields(7).AddCurrentValue ParamValue7
            crPOPrc.ParameterFields(8).ClearCurrentValueAndRange
            crPOPrc.ParameterFields(8).AddCurrentValue ParamValue8
            crPOPrc.ParameterFields(9).ClearCurrentValueAndRange
            crPOPrc.ParameterFields(9).AddCurrentValue ParamValue9
            CRViewer1.ReportSource = crPOPrc
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            If MyPOTYPE = "WHOLESALE" Then
                If Val(SPorderingWS.txtPrnCopy.Text) = 0 Then
                    CRViewer1.EnableExportButton = False
                Else
                    CRViewer1.EnableExportButton = True
                End If
            Else
                If Val(SPordering.txtPrnCopy.Text) = 0 Then
                    CRViewer1.EnableExportButton = False
                Else
                    CRViewer1.EnableExportButton = True
                End If
            End If
        
        Case "SPPOrepl"
            Dim crPOrepl As New rptPOrepl
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crPOrepl.DiscardSavedData
            crPOrepl.Database.SetDataSource rsPrn
            crPOrepl.ParameterFields(1).ClearCurrentValueAndRange
            crPOrepl.ParameterFields(1).AddCurrentValue ParamValue1
            crPOrepl.ParameterFields(2).ClearCurrentValueAndRange
            crPOrepl.ParameterFields(2).AddCurrentValue ParamValue2
            crPOrepl.ParameterFields(3).ClearCurrentValueAndRange
            crPOrepl.ParameterFields(3).AddCurrentValue ParamValue3
            crPOrepl.ParameterFields(4).ClearCurrentValueAndRange
            crPOrepl.ParameterFields(4).AddCurrentValue SysBranch
            crPOrepl.ParameterFields(5).ClearCurrentValueAndRange
            crPOrepl.ParameterFields(5).AddCurrentValue SysAddress
            crPOrepl.ParameterFields(6).ClearCurrentValueAndRange
            crPOrepl.ParameterFields(6).AddCurrentValue SysTel
            crPOrepl.ParameterFields(7).ClearCurrentValueAndRange
            crPOrepl.ParameterFields(7).AddCurrentValue ParamValue7
            crPOrepl.ParameterFields(8).ClearCurrentValueAndRange
            crPOrepl.ParameterFields(8).AddCurrentValue ParamValue8
            crPOrepl.ParameterFields(9).ClearCurrentValueAndRange
            crPOrepl.ParameterFields(9).AddCurrentValue ParamValue9
            crPOrepl.ParameterFields(10).ClearCurrentValueAndRange
            crPOrepl.ParameterFields(10).AddCurrentValue ParamValue10
            crPOrepl.ParameterFields(11).ClearCurrentValueAndRange
            crPOrepl.ParameterFields(11).AddCurrentValue ParamValue11
            crPOrepl.ParameterFields(12).ClearCurrentValueAndRange
            crPOrepl.ParameterFields(12).AddCurrentValue ParamValue12
            CRViewer1.ReportSource = crPOrepl
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            If MyPOTYPE = "WHOLESALE" Then
                If Val(SPorderingWS.txtPrnCopy.Text) = 0 Then
                    CRViewer1.EnableExportButton = False
                Else
                    CRViewer1.EnableExportButton = True
                End If
            Else
                If Val(SPordering.txtPrnCopy.Text) = 0 Then
                    CRViewer1.EnableExportButton = False
                Else
                    CRViewer1.EnableExportButton = True
                End If
            End If
        
        Case "SPPOreplPnum"
            Dim crPOrPnum As New rptPoReplPnum
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crPOrPnum.DiscardSavedData
            crPOrPnum.Database.SetDataSource rsPrn
            crPOrPnum.ParameterFields(1).ClearCurrentValueAndRange
            crPOrPnum.ParameterFields(1).AddCurrentValue ParamValue1
            crPOrPnum.ParameterFields(2).ClearCurrentValueAndRange
            crPOrPnum.ParameterFields(2).AddCurrentValue ParamValue2
            crPOrPnum.ParameterFields(3).ClearCurrentValueAndRange
            crPOrPnum.ParameterFields(3).AddCurrentValue ParamValue3
            crPOrPnum.ParameterFields(4).ClearCurrentValueAndRange
            crPOrPnum.ParameterFields(4).AddCurrentValue SysBranch
            crPOrPnum.ParameterFields(5).ClearCurrentValueAndRange
            crPOrPnum.ParameterFields(5).AddCurrentValue SysAddress
            crPOrPnum.ParameterFields(6).ClearCurrentValueAndRange
            crPOrPnum.ParameterFields(6).AddCurrentValue SysTel
            crPOrPnum.ParameterFields(7).ClearCurrentValueAndRange
            crPOrPnum.ParameterFields(7).AddCurrentValue ParamValue7
            crPOrPnum.ParameterFields(8).ClearCurrentValueAndRange
            crPOrPnum.ParameterFields(8).AddCurrentValue ParamValue8
            crPOrPnum.ParameterFields(9).ClearCurrentValueAndRange
            crPOrPnum.ParameterFields(9).AddCurrentValue ParamValue9
            crPOrPnum.ParameterFields(10).ClearCurrentValueAndRange
            crPOrPnum.ParameterFields(10).AddCurrentValue ParamValue10
            crPOrPnum.ParameterFields(11).ClearCurrentValueAndRange
            crPOrPnum.ParameterFields(11).AddCurrentValue ParamValue11
'            crPOrPnum.ParameterFields(12).ClearCurrentValueAndRange
'            crPOrPnum.ParameterFields(12).AddCurrentValue ParamValue12
            CRViewer1.ReportSource = crPOrPnum
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            If MyPOTYPE = "WHOLESALE" Then
                If Val(SPorderingWS.txtPrnCopy.Text) = 0 Then
                    CRViewer1.EnableExportButton = False
                Else
                    CRViewer1.EnableExportButton = True
                End If
            Else
                If Val(SPordering.txtPrnCopy.Text) = 0 Then
                    CRViewer1.EnableExportButton = False
                Else
                    CRViewer1.EnableExportButton = True
                End If
            End If
        
        Case "SPPOreplAmt"
            Dim crPOramt As New rptPOreplAmt
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crPOramt.DiscardSavedData
            crPOramt.Database.SetDataSource rsPrn
            crPOramt.ParameterFields(1).ClearCurrentValueAndRange
            crPOramt.ParameterFields(1).AddCurrentValue ParamValue1
            crPOramt.ParameterFields(2).ClearCurrentValueAndRange
            crPOramt.ParameterFields(2).AddCurrentValue ParamValue2
            crPOramt.ParameterFields(3).ClearCurrentValueAndRange
            crPOramt.ParameterFields(3).AddCurrentValue ParamValue3
            crPOramt.ParameterFields(4).ClearCurrentValueAndRange
            crPOramt.ParameterFields(4).AddCurrentValue SysBranch
            crPOramt.ParameterFields(5).ClearCurrentValueAndRange
            crPOramt.ParameterFields(5).AddCurrentValue SysAddress
            crPOramt.ParameterFields(6).ClearCurrentValueAndRange
            crPOramt.ParameterFields(6).AddCurrentValue SysTel
            crPOramt.ParameterFields(7).ClearCurrentValueAndRange
            crPOramt.ParameterFields(7).AddCurrentValue ParamValue7
            crPOramt.ParameterFields(8).ClearCurrentValueAndRange
            crPOramt.ParameterFields(8).AddCurrentValue ParamValue8
            crPOramt.ParameterFields(9).ClearCurrentValueAndRange
            crPOramt.ParameterFields(9).AddCurrentValue ParamValue9
            crPOramt.ParameterFields(10).ClearCurrentValueAndRange
            crPOramt.ParameterFields(10).AddCurrentValue ParamValue10
            crPOramt.ParameterFields(11).ClearCurrentValueAndRange
            crPOramt.ParameterFields(11).AddCurrentValue ParamValue11
            crPOramt.ParameterFields(12).ClearCurrentValueAndRange
            crPOramt.ParameterFields(12).AddCurrentValue ParamValue12
            crPOramt.ParameterFields(13).ClearCurrentValueAndRange
            crPOramt.ParameterFields(13).AddCurrentValue ParamValue13
            CRViewer1.ReportSource = crPOramt
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            If MyPOTYPE = "WHOLESALE" Then
                If Val(SPorderingWS.txtPrnCopy.Text) = 0 Then
                    CRViewer1.EnableExportButton = False
                Else
                    CRViewer1.EnableExportButton = True
                End If
            Else
                If Val(SPordering.txtPrnCopy.Text) = 0 Then
                    CRViewer1.EnableExportButton = False
                Else
                    CRViewer1.EnableExportButton = True
                End If
            End If
        
        Case "SPRECEIVE"
            Dim crSPrec As New rptPackingList
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crSPrec.DiscardSavedData
            crSPrec.Database.SetDataSource rsPrn
            crSPrec.ParameterFields(1).ClearCurrentValueAndRange
            crSPrec.ParameterFields(1).AddCurrentValue ParamValue1
            crSPrec.ParameterFields(2).ClearCurrentValueAndRange
            crSPrec.ParameterFields(2).AddCurrentValue ParamValue2
            crSPrec.ParameterFields(3).ClearCurrentValueAndRange
            crSPrec.ParameterFields(3).AddCurrentValue ParamValue3
            crSPrec.ParameterFields(4).ClearCurrentValueAndRange
            crSPrec.ParameterFields(4).AddCurrentValue ParamValue4
            crSPrec.ParameterFields(5).ClearCurrentValueAndRange
            crSPrec.ParameterFields(5).AddCurrentValue ParamValue5
            crSPrec.ParameterFields(6).ClearCurrentValueAndRange
            crSPrec.ParameterFields(6).AddCurrentValue ParamValue6
            crSPrec.ParameterFields(7).ClearCurrentValueAndRange
            crSPrec.ParameterFields(7).AddCurrentValue ParamValue7
            crSPrec.ParameterFields(8).ClearCurrentValueAndRange
            crSPrec.ParameterFields(8).AddCurrentValue ParamValue8
            crSPrec.ParameterFields(9).ClearCurrentValueAndRange
            crSPrec.ParameterFields(9).AddCurrentValue ParamValue9
            
            crSPrec.ParameterFields(10).ClearCurrentValueAndRange
            crSPrec.ParameterFields(10).AddCurrentValue ParamValue10
            crSPrec.ParameterFields(11).ClearCurrentValueAndRange
            crSPrec.ParameterFields(11).AddCurrentValue ParamValue11
'            crSPrec.ParameterFields(7).ClearCurrentValueAndRange
'            crSPrec.ParameterFields(7).AddCurrentValue SysBranch
'            crSPrec.ParameterFields(8).ClearCurrentValueAndRange
'            crSPrec.ParameterFields(8).AddCurrentValue SysAddress
'            crSPrec.ParameterFields(9).ClearCurrentValueAndRange
'            crSPrec.ParameterFields(9).AddCurrentValue SysTel
            
            CRViewer1.ReportSource = crSPrec
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
            
        Case "DMCM [A/R]"
'            Dim crARAdj As New rptAdjustment
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crARAdj.DiscardSavedData
            crARAdj.Database.SetDataSource rsPrn
            crARAdj.ParameterFields(1).ClearCurrentValueAndRange
            crARAdj.ParameterFields(1).AddCurrentValue ParamValue1
            crARAdj.ParameterFields(2).ClearCurrentValueAndRange
            crARAdj.ParameterFields(2).AddCurrentValue ParamValue2
            crARAdj.ParameterFields(3).ClearCurrentValueAndRange
            crARAdj.ParameterFields(3).AddCurrentValue ParamValue3
            crARAdj.ParameterFields(4).ClearCurrentValueAndRange
            crARAdj.ParameterFields(4).AddCurrentValue ParamValue4
            crARAdj.ParameterFields(5).ClearCurrentValueAndRange
            crARAdj.ParameterFields(5).AddCurrentValue SysBranch
            crARAdj.ParameterFields(6).ClearCurrentValueAndRange
            crARAdj.ParameterFields(6).AddCurrentValue SysAddress
            crARAdj.ParameterFields(7).ClearCurrentValueAndRange
            crARAdj.ParameterFields(7).AddCurrentValue SysTel
            CRViewer1.ReportSource = crARAdj
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
             
        Case "MONTHLYSALESADJ"
            
'            Dim crMonthSalAdj As New rptMonthSalesAdj
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crMonthSalAdj.DiscardSavedData
            crMonthSalAdj.Database.SetDataSource rsPrn
            crMonthSalAdj.ParameterFields(1).ClearCurrentValueAndRange
            crMonthSalAdj.ParameterFields(1).AddCurrentValue ParamValue1
            crMonthSalAdj.ParameterFields(2).ClearCurrentValueAndRange
            crMonthSalAdj.ParameterFields(2).AddCurrentValue ParamValue2
            crMonthSalAdj.ParameterFields(3).ClearCurrentValueAndRange
            crMonthSalAdj.ParameterFields(3).AddCurrentValue ParamValue3
            crMonthSalAdj.ParameterFields(4).ClearCurrentValueAndRange
            crMonthSalAdj.ParameterFields(4).AddCurrentValue ParamValue4
            crMonthSalAdj.ParameterFields(5).ClearCurrentValueAndRange
            crMonthSalAdj.ParameterFields(5).AddCurrentValue ParamValue5
            crMonthSalAdj.ParameterFields(6).ClearCurrentValueAndRange
            crMonthSalAdj.ParameterFields(6).AddCurrentValue SysBranch
            crMonthSalAdj.ParameterFields(7).ClearCurrentValueAndRange
            crMonthSalAdj.ParameterFields(7).AddCurrentValue SysAddress
            crMonthSalAdj.ParameterFields(8).AddCurrentValue SysTel
            
            CRViewer1.ReportSource = crMonthSalAdj
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True

        Case "MONTHLYSALESADJ"
'            Dim crFrntLn As New rptMonthSalesAdj
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crMonthSalAdj.DiscardSavedData
            crMonthSalAdj.Database.SetDataSource rsPrn
            crMonthSalAdj.ParameterFields(1).ClearCurrentValueAndRange
            crMonthSalAdj.ParameterFields(1).AddCurrentValue ParamValue1
            crMonthSalAdj.ParameterFields(2).ClearCurrentValueAndRange
            crMonthSalAdj.ParameterFields(2).AddCurrentValue ParamValue2
            crMonthSalAdj.ParameterFields(3).ClearCurrentValueAndRange
            crMonthSalAdj.ParameterFields(3).AddCurrentValue ParamValue3
            crMonthSalAdj.ParameterFields(4).ClearCurrentValueAndRange
            crMonthSalAdj.ParameterFields(4).AddCurrentValue ParamValue4
            crMonthSalAdj.ParameterFields(5).ClearCurrentValueAndRange
            crMonthSalAdj.ParameterFields(5).AddCurrentValue ParamValue5
            crMonthSalAdj.ParameterFields(6).ClearCurrentValueAndRange
            crMonthSalAdj.ParameterFields(6).AddCurrentValue SysBranch
            crMonthSalAdj.ParameterFields(7).ClearCurrentValueAndRange
            crMonthSalAdj.ParameterFields(7).AddCurrentValue SysAddress
            crMonthSalAdj.ParameterFields(8).AddCurrentValue SysTel
            
            CRViewer1.ReportSource = crMonthSalAdj
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
        
        Case "STOCKTRANSFER"
            Dim crSPtrans As New RptSPtransfer
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crSPtrans.DiscardSavedData
            crSPtrans.Database.SetDataSource rsPrn
            crSPtrans.ParameterFields(1).ClearCurrentValueAndRange
            crSPtrans.ParameterFields(1).AddCurrentValue ParamValue1
            crSPtrans.ParameterFields(2).ClearCurrentValueAndRange
            crSPtrans.ParameterFields(2).AddCurrentValue ParamValue2
            crSPtrans.ParameterFields(3).ClearCurrentValueAndRange
            crSPtrans.ParameterFields(3).AddCurrentValue ParamValue3
            crSPtrans.ParameterFields(4).ClearCurrentValueAndRange
            crSPtrans.ParameterFields(4).AddCurrentValue ParamValue4
            crSPtrans.ParameterFields(5).ClearCurrentValueAndRange
            crSPtrans.ParameterFields(5).AddCurrentValue ParamValue5
            crSPtrans.ParameterFields(6).ClearCurrentValueAndRange
            crSPtrans.ParameterFields(6).AddCurrentValue ParamValue6
            crSPtrans.ParameterFields(7).ClearCurrentValueAndRange
            crSPtrans.ParameterFields(7).AddCurrentValue ParamValue7
            crSPtrans.ParameterFields(8).ClearCurrentValueAndRange
            crSPtrans.ParameterFields(8).AddCurrentValue ParamValue8
            crSPtrans.ParameterFields(9).ClearCurrentValueAndRange
            crSPtrans.ParameterFields(9).AddCurrentValue ParamValue9
            crSPtrans.ParameterFields(10).ClearCurrentValueAndRange
            crSPtrans.ParameterFields(10).AddCurrentValue ParamValue10
            crSPtrans.ParameterFields(11).ClearCurrentValueAndRange
            crSPtrans.ParameterFields(11).AddCurrentValue ParamValue11
            'crSPtrans.ParameterFields(12).ClearCurrentValueAndRange
            'crSPtrans.ParameterFields(12).AddCurrentValue ParamValue12
            CRViewer1.ReportSource = crSPtrans
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
            
            Case "STOCKTRANSFER PENDING"
            Dim crSPtransPending As New RptSPtransferPending

            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crSPtransPending.DiscardSavedData
            crSPtransPending.Database.SetDataSource rsPrn
            crSPtransPending.ParameterFields(1).ClearCurrentValueAndRange
            crSPtransPending.ParameterFields(1).AddCurrentValue ParamValue1
            crSPtransPending.ParameterFields(2).ClearCurrentValueAndRange
            crSPtransPending.ParameterFields(2).AddCurrentValue ParamValue2
            crSPtransPending.ParameterFields(3).ClearCurrentValueAndRange
            crSPtransPending.ParameterFields(3).AddCurrentValue ParamValue3
            crSPtransPending.ParameterFields(4).ClearCurrentValueAndRange
            crSPtransPending.ParameterFields(4).AddCurrentValue ParamValue4
            crSPtransPending.ParameterFields(5).ClearCurrentValueAndRange
            crSPtransPending.ParameterFields(5).AddCurrentValue ParamValue5
            crSPtransPending.ParameterFields(6).ClearCurrentValueAndRange
            crSPtransPending.ParameterFields(6).AddCurrentValue ParamValue6
            crSPtransPending.ParameterFields(7).ClearCurrentValueAndRange
            crSPtransPending.ParameterFields(7).AddCurrentValue ParamValue7
            crSPtransPending.ParameterFields(8).ClearCurrentValueAndRange
            crSPtransPending.ParameterFields(8).AddCurrentValue ParamValue8
            crSPtransPending.ParameterFields(9).ClearCurrentValueAndRange
            crSPtransPending.ParameterFields(9).AddCurrentValue ParamValue9
            crSPtransPending.ParameterFields(10).ClearCurrentValueAndRange
            crSPtransPending.ParameterFields(10).AddCurrentValue ParamValue10
            crSPtransPending.ParameterFields(11).ClearCurrentValueAndRange
            crSPtransPending.ParameterFields(11).AddCurrentValue ParamValue11
            'crSPtrans.ParameterFields(12).ClearCurrentValueAndRange
            'crSPtrans.ParameterFields(12).AddCurrentValue ParamValue12
            CRViewer1.ReportSource = crSPtransPending
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
'
        Case "STOCKTRANSFER W/ COST"
            Dim crStkTrnsCost As New RptSpTransCost
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crStkTrnsCost.DiscardSavedData
            crStkTrnsCost.Database.SetDataSource rsPrn
            crStkTrnsCost.ParameterFields(1).ClearCurrentValueAndRange
            crStkTrnsCost.ParameterFields(1).AddCurrentValue ParamValue1
            crStkTrnsCost.ParameterFields(2).ClearCurrentValueAndRange
            crStkTrnsCost.ParameterFields(2).AddCurrentValue ParamValue2
            crStkTrnsCost.ParameterFields(3).ClearCurrentValueAndRange
            crStkTrnsCost.ParameterFields(3).AddCurrentValue ParamValue3
            crStkTrnsCost.ParameterFields(4).ClearCurrentValueAndRange
            crStkTrnsCost.ParameterFields(4).AddCurrentValue ParamValue4
            crStkTrnsCost.ParameterFields(5).ClearCurrentValueAndRange
            crStkTrnsCost.ParameterFields(5).AddCurrentValue ParamValue53
            crStkTrnsCost.ParameterFields(6).ClearCurrentValueAndRange
            crStkTrnsCost.ParameterFields(6).AddCurrentValue ParamValue6
            crStkTrnsCost.ParameterFields(7).ClearCurrentValueAndRange
            crStkTrnsCost.ParameterFields(7).AddCurrentValue ParamValue7
            crStkTrnsCost.ParameterFields(8).ClearCurrentValueAndRange
            crStkTrnsCost.ParameterFields(8).AddCurrentValue ParamValue8
            crStkTrnsCost.ParameterFields(9).ClearCurrentValueAndRange
            crStkTrnsCost.ParameterFields(9).AddCurrentValue ParamValue9
            crStkTrnsCost.ParameterFields(10).ClearCurrentValueAndRange
            crStkTrnsCost.ParameterFields(10).AddCurrentValue ParamValue10
            crStkTrnsCost.ParameterFields(11).ClearCurrentValueAndRange
            crStkTrnsCost.ParameterFields(11).AddCurrentValue ParamValue11
            CRViewer1.ReportSource = crStkTrnsCost
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
        Case "TRANSSUMMARY"
            Dim crSummary As New RptTransferSum
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crSummary.DiscardSavedData
            crSummary.Database.SetDataSource rsPrn
            crSummary.ParameterFields(1).ClearCurrentValueAndRange
            crSummary.ParameterFields(1).AddCurrentValue ParamValue1
            crSummary.ParameterFields(2).ClearCurrentValueAndRange
            crSummary.ParameterFields(2).AddCurrentValue ParamValue2
            crSummary.ParameterFields(3).ClearCurrentValueAndRange
            crSummary.ParameterFields(3).AddCurrentValue ParamValue3
            CRViewer1.ReportSource = crSummary
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
        
        Case "AUDITPAGE1"
            Dim crPage1 As New RptAudInvPg1
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crPage1.DiscardSavedData
            crPage1.Database.SetDataSource rsPrn
            crPage1.ParameterFields(1).ClearCurrentValueAndRange
            crPage1.ParameterFields(1).AddCurrentValue ParamValue1
            crPage1.ParameterFields(2).ClearCurrentValueAndRange
            crPage1.ParameterFields(2).AddCurrentValue ParamValue2
            'crPage1.ParameterFields(3).ClearCurrentValueAndRange
            'crPage1.ParameterFields(3).AddCurrentValue ParamValue3
            CRViewer1.ReportSource = crPage1
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
            
        Case "Back Order Adj."
            Dim crBOInvAdj As New rptBOInvAdj
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crBOInvAdj.DiscardSavedData
            crBOInvAdj.Database.SetDataSource rsPrn
            crBOInvAdj.ParameterFields(1).ClearCurrentValueAndRange
            crBOInvAdj.ParameterFields(1).AddCurrentValue ParamValue1
            crBOInvAdj.ParameterFields(2).ClearCurrentValueAndRange
            crBOInvAdj.ParameterFields(2).AddCurrentValue ParamValue2
            crBOInvAdj.ParameterFields(3).ClearCurrentValueAndRange
            crBOInvAdj.ParameterFields(3).AddCurrentValue ParamValue3
            crBOInvAdj.ParameterFields(4).ClearCurrentValueAndRange
            crBOInvAdj.ParameterFields(4).AddCurrentValue ParamValue4
            crBOInvAdj.ParameterFields(5).ClearCurrentValueAndRange
            crBOInvAdj.ParameterFields(5).AddCurrentValue ParamValue5
            crBOInvAdj.ParameterFields(6).ClearCurrentValueAndRange
            crBOInvAdj.ParameterFields(6).AddCurrentValue ParamValue6
            crBOInvAdj.ParameterFields(7).ClearCurrentValueAndRange
            crBOInvAdj.ParameterFields(7).AddCurrentValue ParamValue7
            crBOInvAdj.ParameterFields(8).ClearCurrentValueAndRange
            crBOInvAdj.ParameterFields(8).AddCurrentValue ParamValue8
            crBOInvAdj.ParameterFields(9).ClearCurrentValueAndRange
            crBOInvAdj.ParameterFields(9).AddCurrentValue ParamValue9
            crBOInvAdj.ParameterFields(10).ClearCurrentValueAndRange
            crBOInvAdj.ParameterFields(10).AddCurrentValue ParamValue10
            CRViewer1.ReportSource = crBOInvAdj
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
            
        Case "WS Inventory Adjustment"
            Dim crWSInvAdj As New rptWSInvAdj
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crWSInvAdj.DiscardSavedData
            crWSInvAdj.Database.SetDataSource rsPrn
            crWSInvAdj.ParameterFields(1).ClearCurrentValueAndRange
            crWSInvAdj.ParameterFields(1).AddCurrentValue ParamValue1
            crWSInvAdj.ParameterFields(2).ClearCurrentValueAndRange
            crWSInvAdj.ParameterFields(2).AddCurrentValue ParamValue2
            crWSInvAdj.ParameterFields(3).ClearCurrentValueAndRange
            crWSInvAdj.ParameterFields(3).AddCurrentValue ParamValue3
            crWSInvAdj.ParameterFields(4).ClearCurrentValueAndRange
            crWSInvAdj.ParameterFields(4).AddCurrentValue ParamValue4
            crWSInvAdj.ParameterFields(5).ClearCurrentValueAndRange
            crWSInvAdj.ParameterFields(5).AddCurrentValue ParamValue5
            crWSInvAdj.ParameterFields(6).ClearCurrentValueAndRange
            crWSInvAdj.ParameterFields(6).AddCurrentValue ParamValue6
            crWSInvAdj.ParameterFields(7).ClearCurrentValueAndRange
            crWSInvAdj.ParameterFields(7).AddCurrentValue ParamValue7
            crWSInvAdj.ParameterFields(8).ClearCurrentValueAndRange
            crWSInvAdj.ParameterFields(8).AddCurrentValue ParamValue8
            crWSInvAdj.ParameterFields(9).ClearCurrentValueAndRange
            crWSInvAdj.ParameterFields(9).AddCurrentValue ParamValue9
            crWSInvAdj.ParameterFields(10).ClearCurrentValueAndRange
            crWSInvAdj.ParameterFields(10).AddCurrentValue ParamValue10
            CRViewer1.ReportSource = crWSInvAdj
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
        Case "RETAIL STF SUMMARY"
            Dim crRSTFSummary As New rptRetailSTFSummary
            
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crRSTFSummary.DiscardSavedData
            crRSTFSummary.Database.SetDataSource rsPrn
            crRSTFSummary.ParameterFields(1).ClearCurrentValueAndRange
            crRSTFSummary.ParameterFields(1).AddCurrentValue ParamValue1
            crRSTFSummary.ParameterFields(2).ClearCurrentValueAndRange
            crRSTFSummary.ParameterFields(2).AddCurrentValue ParamValue2
            crRSTFSummary.ParameterFields(3).ClearCurrentValueAndRange
            crRSTFSummary.ParameterFields(3).AddCurrentValue ParamValue3
            CRViewer1.ReportSource = crRSTFSummary
            CRViewer1.ViewReport
            CRViewer1.DisplayGroupTree = False
            CRViewer1.EnableExportButton = True
            
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    QryPrn = ""
    MySource = ""
    If rsPrn.State = adStateOpen Then rsPrn.Close
End Sub

Private Sub Form_Resize()
    With CRViewer1
        .Top = cmdExit.Height + 50
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - Warehouse.sBar.Height
    End With
End Sub




