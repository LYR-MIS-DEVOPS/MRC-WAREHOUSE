VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmPrint 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6795
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6795
   WindowState     =   1  'Minimized
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
      m_MDI_Width     =   6885
      m_MDI_Height    =   6030
      OriginalBorder  =   10
      OriginalCaption =   31
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public QryPrn As String
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
    End Select
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
