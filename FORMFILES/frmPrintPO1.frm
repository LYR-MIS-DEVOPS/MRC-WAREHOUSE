VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPrintPO 
   Caption         =   "Print Preview"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   840
   ClientWidth     =   11115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   11115
   WindowState     =   2  'Maximized
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Exit"
      CapAlign        =   2
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
      cBack           =   -2147483633
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3000
      Top             =   7800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin Crystal.CrystalReport cRPT 
      Left            =   1800
      Top             =   7920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin CRVIEWERLibCtl.CRViewer crViewer 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Tag             =   "xyf"
      Top             =   600
      Width           =   11055
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
      Left            =   840
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11355
      m_MDI_Height    =   8220
   End
   Begin VB.Menu mnuprinting 
      Caption         =   ""
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmPrintPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M_Proj As CRAXDRT.Application
Dim CF_JOB As CRAXDRT.Report


'Dim Report As CRAXDRT.Report
'Public DisplayDateRange As String
'Public DisplayPlaceField
Option Explicit


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
    rsPrn.Close
End Sub

Private Sub crViewer_PrintButtonClicked(UseDefault As Boolean)
    
    If MySource = "Branch Purchase Order" Then
        If Val(BranchPO.txtPrnCopy.Text) = 0 Then
           'Call Receiving.UpdateItemReceived
            If Val(BranchPO.txtPrnCopy.Text) > 0 Then Exit Sub
            BranchPO.txtPrnCopy.Text = Val(BranchPO.txtPrnCopy.Text) + 1
            'LYRcon.Execute "UPDATE invoutgoing SET PRNCOPY = invoutgoing.PRNCOPY + " & 1 & " Where invoutgoing.DRID = " & Val(txtDRid.Text) & ";"
            LYRcon.Execute "UPDATE sporderbranch SET PRINTPICK = sporderbranch.PRINTPICK + " & 1 & ", `STATUS` = 'PICKING' Where sporderbranch.BRANCHSPOID = " & Val(BranchPO.txtBranchSPOID.Text) & " AND sporderbranch.REFERENCE = '" & BranchPO.txtReference.Text & "';"
       
        Else
           ' BranchPO.txtPrnCopy.Text = Val(BranchPO.txtPrnCopy.Text) + 1
            'LYRcon.Execute "UPDATE invreceiving SET PRNCOPY = invreceiving.PRNCOPY + " & 1 & " Where invreceiving.RRECID = " & Val(Receiving.txtRecID.Text) & ";"
            LYRcon.Execute "UPDATE sporderbranch SET PRINTPICK = sporderbranch.PRINTPICK + " & 1 & "  Where sporderbranch.BRANCHSPOID = " & Val(BranchPO.txtBranchSPOID.Text) & " AND sporderbranch.REFERENCE = '" & BranchPO.txtReference.Text & "';"
        End If
        
    ElseIf MySource = "SP PACKING" Then
            
    End If
    
End Sub

Private Sub Form_Load()
    Select Case MySource
        
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
            crViewer.ReportSource = crDtlPO
            crViewer.ViewReport
            crViewer.DisplayGroupTree = False
            crViewer.EnableExportButton = True
        
        Case "SP PACKING"
            Dim crSpTrans As New rptSpTrans
            
            If rsPrn.State = adStateOpen Then rsPrn.Close
            rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crSpTrans.DiscardSavedData
            crSpTrans.Database.SetDataSource rsPrn
            crSpTrans.ParameterFields(1).ClearCurrentValueAndRange
            crSpTrans.ParameterFields(1).AddCurrentValue ParamValue1
            crSpTrans.ParameterFields(2).ClearCurrentValueAndRange
            crSpTrans.ParameterFields(2).AddCurrentValue ParamValue2
            crSpTrans.ParameterFields(3).ClearCurrentValueAndRange
            crSpTrans.ParameterFields(3).AddCurrentValue ParamValue3
            crSpTrans.ParameterFields(4).ClearCurrentValueAndRange
            crSpTrans.ParameterFields(4).AddCurrentValue ParamValue4
            crSpTrans.ParameterFields(5).ClearCurrentValueAndRange
            crSpTrans.ParameterFields(5).AddCurrentValue ParamValue5
            crSpTrans.ParameterFields(6).ClearCurrentValueAndRange
            crSpTrans.ParameterFields(6).AddCurrentValue ParamValue6
            crSpTrans.ParameterFields(7).ClearCurrentValueAndRange
            crSpTrans.ParameterFields(7).AddCurrentValue ParamValue7
            crSpTrans.ParameterFields(8).ClearCurrentValueAndRange
            crSpTrans.ParameterFields(8).AddCurrentValue ParamValue8
            crSpTrans.ParameterFields(9).ClearCurrentValueAndRange
            crSpTrans.ParameterFields(9).AddCurrentValue ParamValue9
            crSpTrans.ParameterFields(10).ClearCurrentValueAndRange
            crSpTrans.ParameterFields(10).AddCurrentValue ParamValue10
            crSpTrans.ParameterFields(11).ClearCurrentValueAndRange
            crSpTrans.ParameterFields(11).AddCurrentValue ParamValue11
            crViewer.ReportSource = crSpTrans
            crViewer.ViewReport
            crViewer.DisplayGroupTree = False
            crViewer.EnableExportButton = True
    End Select
    
    
        Me.Width = Warehouse.Width - (Warehouse.Width * 0.25)
        Me.Height = Warehouse.Height - (Warehouse.Height * 0.3)
End Sub


