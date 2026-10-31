VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmReport 
   Caption         =   "Form1"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9705
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6885
      Left            =   0
      TabIndex        =   0
      Tag             =   "xy"
      Top             =   480
      Width           =   9645
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
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
      Image           =   "frmReport.frx":058A
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   6960
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   9945
      m_MDI_Height    =   8025
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Public RFPID As Long

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CRViewer1_CloseButtonClicked(UseDefault As Boolean)
    If MySource = "RFPOTHER" Then
        LYRcon.Execute "Update rfp_main SET printcopy = printcopy + 1 Where RFID = " & RFPID & " AND RTYPE = 'OTHER'; "
        frmRFPOther.PrintCopy = 1
    End If
End Sub

Private Sub Form_Load()
    If MySource = "PRICING" Then
        Dim crtBBOsum As New Pricingsmryrpt
        
        rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
        crtBBOsum.DiscardSavedData
        crtBBOsum.Database.SetDataSource rsPrn
        crtBBOsum.ParameterFields(1).ClearCurrentValueAndRange
        crtBBOsum.ParameterFields(1).AddCurrentValue ParamValue1
        crtBBOsum.ParameterFields(2).ClearCurrentValueAndRange
        crtBBOsum.ParameterFields(2).AddCurrentValue ParamValue2
        
        CRViewer1.ReportSource = crtBBOsum
        CRViewer1.ViewReport
        CRViewer1.DisplayGroupTree = False
        CRViewer1.EnableExportButton = True
        CRViewer1.Zoom 100
    ElseIf MySource = "PARTNUMBER" Then
         Dim crtPrtnmbrsum As New partnumberRpt
        
        rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
        crtPrtnmbrsum.DiscardSavedData
        crtPrtnmbrsum.Database.SetDataSource rsPrn
        crtPrtnmbrsum.ParameterFields(1).ClearCurrentValueAndRange
        crtPrtnmbrsum.ParameterFields(1).AddCurrentValue ParamValue1
        crtPrtnmbrsum.ParameterFields(2).ClearCurrentValueAndRange
        crtPrtnmbrsum.ParameterFields(2).AddCurrentValue ParamValue2
        
        CRViewer1.ReportSource = crtPrtnmbrsum
        CRViewer1.ViewReport
        CRViewer1.DisplayGroupTree = False
        CRViewer1.EnableExportButton = True
        CRViewer1.Zoom 100
    ElseIf MySource = "RFPOTHERSUM" Then
        Dim rpt3 As New rptRFPOtherSum
        
        If rsPrn.State = adStateOpen Then rsPrn.Close
        rsPrn.Open QryPrn, LYRcon, adOpenDynamic, adLockOptimistic
        rpt3.ParameterFields(1).ClearCurrentValueAndRange
        rpt3.ParameterFields(1).AddCurrentValue UserFullName
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = rpt3
        CRViewer1.ViewReport
        CRViewer1.Zoom 100
        Screen.MousePointer = vbDefault
        CRViewer1.EnableExportButton = False
    
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsPrn.State = adStateOpen Then rsPrn.Close
End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub
