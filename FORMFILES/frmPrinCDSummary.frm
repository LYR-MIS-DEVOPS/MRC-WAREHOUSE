VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPrinCDSummary 
   Caption         =   "Print Purchase Order Summary"
   ClientHeight    =   8175
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Tag             =   "XYF"
      Top             =   480
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "PRINT PAGE 1"
      TabPicture(0)   =   "frmPrinCDSummary.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "crViewerPageOne"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "PRINT PAGE 2"
      TabPicture(1)   =   "frmPrinCDSummary.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "crViewerPageTwo"
      Tab(1).ControlCount=   1
      Begin CRVIEWERLibCtl.CRViewer crViewerPageOne 
         Height          =   7005
         Left            =   120
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   360
         Width           =   11085
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
         DisplayBorder   =   0   'False
         DisplayTabs     =   -1  'True
         DisplayBackgroundEdge=   -1  'True
         SelectionFormula=   ""
         EnablePopupMenu =   -1  'True
         EnableExportButton=   0   'False
         EnableSearchExpertButton=   0   'False
         EnableHelpButton=   0   'False
      End
      Begin CRVIEWERLibCtl.CRViewer crViewerPageTwo 
         Height          =   7005
         Left            =   -74880
         TabIndex        =   3
         Tag             =   "XYF"
         Top             =   360
         Width           =   11085
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
         DisplayBorder   =   0   'False
         DisplayTabs     =   -1  'True
         DisplayBackgroundEdge=   -1  'True
         SelectionFormula=   ""
         EnablePopupMenu =   -1  'True
         EnableExportButton=   0   'False
         EnableSearchExpertButton=   0   'False
         EnableHelpButton=   0   'False
      End
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   120
      TabIndex        =   2
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
      Image           =   "frmPrinCDSummary.frx":0038
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   7680
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11775
      m_MDI_Height    =   8745
   End
End
Attribute VB_Name = "frmPrinCDSummary"
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
End Sub

Private Sub Form_Load()
   
    Select Case MySource
        Case "Purchase Order Summary Report"
            Dim crRSummPageOne As New rptRetailPOSummaryPAGE1
            Dim crRSummPageTwo As New rptRetailPOSummaryPAGE2
            
            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crRSummPageOne.DiscardSavedData
            crRSummPageOne.Database.SetDataSource rsPrint
            crRSummPageOne.ParameterFields(1).ClearCurrentValueAndRange
            crRSummPageOne.ParameterFields(1).AddCurrentValue ParamValue1
            crRSummPageOne.ParameterFields(2).ClearCurrentValueAndRange
            crRSummPageOne.ParameterFields(2).AddCurrentValue ParamValue2
            crRSummPageOne.ParameterFields(3).ClearCurrentValueAndRange
            crRSummPageOne.ParameterFields(3).AddCurrentValue ParamValue3
            crRSummPageOne.ParameterFields(4).ClearCurrentValueAndRange
            crRSummPageOne.ParameterFields(4).AddCurrentValue ParamValue4
            crViewerPageOne.ReportSource = crRSummPageOne
            crViewerPageOne.ViewReport
            crViewerPageOne.DisplayGroupTree = False
            crViewerPageOne.EnableExportButton = True
            
            If rsPrint.State = adStateOpen Then rsPrint.Close
            rsPrint.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
            crRSummPageTwo.DiscardSavedData
            crRSummPageTwo.Database.SetDataSource rsPrint
            crRSummPageTwo.ParameterFields(1).ClearCurrentValueAndRange
            crRSummPageTwo.ParameterFields(1).AddCurrentValue ParamValue1
            crRSummPageTwo.ParameterFields(2).ClearCurrentValueAndRange
            crRSummPageTwo.ParameterFields(2).AddCurrentValue ParamValue2
            crRSummPageTwo.ParameterFields(3).ClearCurrentValueAndRange
            crRSummPageTwo.ParameterFields(3).AddCurrentValue ParamValue3
            crRSummPageTwo.ParameterFields(4).ClearCurrentValueAndRange
            crRSummPageTwo.ParameterFields(4).AddCurrentValue ParamValue4
            crViewerPageTwo.ReportSource = crRSummPageTwo
            crViewerPageTwo.ViewReport
            crViewerPageTwo.DisplayGroupTree = False
            crViewerPageTwo.EnableExportButton = True
'****************** END CODE UPDATED: 03/16/2018 ******************************
    End Select
End Sub
