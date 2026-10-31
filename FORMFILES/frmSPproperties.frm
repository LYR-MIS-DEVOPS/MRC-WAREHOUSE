VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSPproperties 
   Caption         =   "Spareparts Properties"
   ClientHeight    =   6735
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14250
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   14250
   Tag             =   "XY"
   WindowState     =   2  'Maximized
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   44
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "Add &New"
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
      Image           =   "frmSPproperties.frx":0000
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtSPNo 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Tag             =   "XY"
      Top             =   720
      Width           =   3375
   End
   Begin VB.Timer tmrSUBSP 
      Interval        =   500
      Left            =   14400
      Top             =   0
   End
   Begin VB.Timer tmrCboDisplaySP 
      Interval        =   500
      Left            =   14400
      Top             =   600
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6165
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Cancel"
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
      Image           =   "frmSPproperties.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4680
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Save"
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
      Image           =   "frmSPproperties.frx":0B34
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtPMID 
      Height          =   285
      Left            =   14400
      TabIndex        =   35
      Tag             =   "XY"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtSPID 
      Height          =   285
      Left            =   14400
      TabIndex        =   34
      Tag             =   "XY"
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer tmrCboDisplay 
      Interval        =   500
      Left            =   15120
      Top             =   600
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   14400
      TabIndex        =   33
      Text            =   "HoldBackColor"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   14400
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3100
      TabIndex        =   30
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
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
      Image           =   "frmSPproperties.frx":10CE
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame2 
      Caption         =   " SUBSTITUTE SPAREPARTS"
      Height          =   2535
      Left            =   120
      TabIndex        =   26
      Tag             =   "XY"
      Top             =   4080
      Width           =   14055
      Begin MSFlexGridLib.MSFlexGrid flxSUBSP 
         Height          =   2055
         Left            =   240
         TabIndex        =   27
         Tag             =   "XY"
         Top             =   360
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   7
         RowHeightMin    =   310
         BackColorBkg    =   -2147483633
         FormatString    =   $"frmSPproperties.frx":1668
      End
      Begin VB.CommandButton cmdAltA 
         Caption         =   "&A"
         Height          =   255
         Left            =   1800
         TabIndex        =   42
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton cmdAltD 
         Caption         =   "&D"
         Height          =   255
         Left            =   2280
         TabIndex        =   43
         Top             =   1320
         Width           =   375
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1600
      TabIndex        =   29
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Update"
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
      Image           =   "frmSPproperties.frx":172D
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Tag             =   "XY"
      Top             =   360
      Width           =   14055
      Begin VB.TextBox txtRecommended 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   13920
         Locked          =   -1  'True
         TabIndex        =   49
         Text            =   "RECOMMENDED BARCODE VALUE:"
         Top             =   360
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox cboCatalogPath 
         Height          =   315
         Left            =   7680
         TabIndex        =   48
         Tag             =   "XY"
         Top             =   1800
         Width           =   6135
      End
      Begin VB.ComboBox cboSubCat 
         Height          =   315
         Left            =   7680
         TabIndex        =   46
         Tag             =   "XY"
         Top             =   840
         Width           =   6135
      End
      Begin MSMask.MaskEdBox txtSTARTSALE 
         Height          =   255
         Left            =   1920
         TabIndex        =   40
         Tag             =   "XY"
         Top             =   3240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtENDSALE 
         Height          =   255
         Left            =   7680
         TabIndex        =   31
         Tag             =   "XY"
         Top             =   3240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   10800
         TabIndex        =   11
         Tag             =   "XY"
         Top             =   3240
         Width           =   3015
      End
      Begin VB.ComboBox cboPMMC 
         Height          =   315
         Left            =   7680
         TabIndex        =   9
         Tag             =   "XY"
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtBARCODE 
         Height          =   315
         Left            =   7680
         TabIndex        =   8
         Tag             =   "XY"
         Top             =   2280
         Width           =   6135
      End
      Begin VB.ComboBox cboClass 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Tag             =   "XY"
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txtROPoint 
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Tag             =   "XY"
         Top             =   1320
         Width           =   3375
      End
      Begin VB.CheckBox chkEndSale 
         Caption         =   "SALE DATE END:"
         Height          =   255
         Left            =   5760
         TabIndex        =   21
         Tag             =   "XY"
         Top             =   3240
         Width           =   2295
      End
      Begin VB.ComboBox cboSuppBrand 
         Height          =   315
         Left            =   1920
         TabIndex        =   6
         Tag             =   "XY"
         Top             =   2760
         Width           =   3375
      End
      Begin VB.ComboBox cboPMModel 
         Height          =   315
         Left            =   10920
         TabIndex        =   10
         Tag             =   "XY"
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtApplic 
         Height          =   315
         Left            =   7680
         TabIndex        =   12
         Tag             =   "XY"
         Top             =   2760
         Width           =   6135
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   7680
         TabIndex        =   7
         Tag             =   "XY"
         Top             =   360
         Width           =   6135
      End
      Begin VB.ComboBox cboSupp 
         Height          =   315
         Left            =   1920
         TabIndex        =   5
         Tag             =   "XY"
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox txtMAXPO 
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Tag             =   "XY"
         Top             =   1800
         Width           =   3375
      End
      Begin VB.CheckBox chkStartSale 
         Caption         =   "SALE DATE:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Tag             =   "XY"
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "PARTS CATALOG:"
         Height          =   255
         Index           =   13
         Left            =   5760
         TabIndex        =   47
         Tag             =   "XY"
         Top             =   1830
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "SUB CATEGORY"
         Height          =   255
         Index           =   12
         Left            =   5760
         TabIndex        =   45
         Tag             =   "XY"
         Top             =   870
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "STATUS:"
         Height          =   255
         Index           =   11
         Left            =   10080
         TabIndex        =   28
         Tag             =   "XY"
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "PARTS MATRIX MC"
         Height          =   255
         Index           =   6
         Left            =   5760
         TabIndex        =   25
         Tag             =   "XY"
         Top             =   1350
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "BARCODE VALUE"
         Height          =   255
         Index           =   5
         Left            =   5760
         TabIndex        =   24
         Tag             =   "XY"
         Top             =   2310
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "SUPPLIER PRODUCT CODE"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   23
         Tag             =   "XY"
         Top             =   1350
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "CATEGORY"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Tag             =   "XY"
         Top             =   870
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "PRODUCT :"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   20
         Tag             =   "XY"
         Top             =   2790
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "SUPPLIER"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   19
         Tag             =   "XY"
         Top             =   2310
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "APPLICABLE MODEL:"
         Height          =   255
         Index           =   8
         Left            =   5760
         TabIndex        =   18
         Tag             =   "XY"
         Top             =   2790
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "MC MODEL"
         Height          =   255
         Index           =   7
         Left            =   9960
         TabIndex        =   17
         Tag             =   "XY"
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "DESCRIPTION"
         Height          =   255
         Index           =   4
         Left            =   5760
         TabIndex        =   16
         Tag             =   "XY"
         Top             =   390
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "MAXIMUM PO QTY."
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Tag             =   "XY"
         Top             =   1830
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "PARTNUMBER"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Tag             =   "XY"
         Top             =   360
         Width           =   2775
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxMainPartNo 
      Height          =   1455
      Left            =   14400
      TabIndex        =   41
      Tag             =   "XY"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2566
      _Version        =   393216
      BackColor       =   16777215
      FixedRows       =   0
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      GridColorUnpopulated=   16777215
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPARTNO 
      Height          =   1935
      Left            =   14400
      TabIndex        =   39
      Tag             =   "XY"
      Top             =   3240
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   3413
      _Version        =   393216
      BackColor       =   16777215
      FixedRows       =   0
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      GridColorUnpopulated=   16777215
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtPN 
      Height          =   285
      Left            =   14400
      TabIndex        =   38
      Tag             =   "XY"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   15120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14490
      m_MDI_Height    =   7305
   End
   Begin VB.Menu mnuStatus 
      Caption         =   "Popupmenu"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add Row"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete Row"
      End
   End
   Begin VB.Menu mnuStatus2 
      Caption         =   "Popupmenu2"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuBarcode 
         Caption         =   "View Recommended Barcode"
      End
   End
End
Attribute VB_Name = "frmSPproperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPType As New ADODB.Recordset, rsSupplier As New ADODB.Recordset
Dim rsParts As New ADODB.Recordset, rsPMBRAND As New ADODB.Recordset
Dim rsSUPPBRAND As New ADODB.Recordset, rsPMMODEL As New ADODB.Recordset
Dim rsSubParts As New ADODB.Recordset, rsMain As New ADODB.Recordset
Dim rsSUBCAT As New ADODB.Recordset

Dim Qry As String, formAction As String
Dim strPartial As String, strTotal As String
Dim arrSPNO() As String, arrPMMODEL() As String
Dim Tmp() As String, ans As String, MyErr As String
Dim myMsg As String, cmbstr As String

Dim i As Long, j As Long
Dim ActRow As Integer, ActCol As Integer

Dim StatSrch As Boolean, PrtBln As Boolean
Dim PrtSrch As Boolean, ProperExit As Boolean, blnFound As Boolean

Public Sub flxChanger()
    With flxSUBSP
        .Cols = 9
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.15   'PARTS NUMBER
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.2  'DESCRIPTION
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.1 'CLASSIFICATION
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.13 'SUPPLIER
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.2 'SUPPLIER BRAND
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.2 'PARTSMATRIX MODEL
        .ColWidth(7) = 0 '(.Width - .ColWidth(0)) * 0.06 'SPID
        .ColWidth(8) = 0 '(.Width - .ColWidth(0)) * 0.06 'PMID
        
        flxPartNo.ColWidth(1) = .ColWidth(1)
        
    End With
    
    With flxPartNo
        .Top = txtPN.Height + txtPN.Top
        .Left = txtPN.Left
        .Width = txtPN.Width
    End With
    txtSPNo.SetFocus
End Sub



Public Sub RefreshRecord()

    Screen.MousePointer = vbHourglass
    DoEvents
    Qry = "call sp_mrcwarehouse.sp_SPProperties_getSpareparts()"
    If rsParts.State = adStateOpen Then rsParts.Close
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    
    With rsParts
        flxMainPartNo.Clear
        Set flxMainPartNo.DataSource = rsParts
    End With
    
    With flxMainPartNo
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = txtSPNo.Width
        For i = 3 To 21
            .ColWidth(i) = 0
        Next i
    End With
    Screen.MousePointer = Default
End Sub

Private Sub SetButtonVisibility(DataEntry As Boolean)
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->VIEW SUMM
'3 ->EXIT
'4 ->PRINT
'5 ->SAVE
'6 ->CANCEL

    'hide all button
    For i = 0 To 2
        cmdTrans(i).Visible = False
    Next i

    If DataEntry Then
        cmdTrans(3).Visible = True
        cmdTrans(4).Visible = True
    Else
        For i = 0 To 2
            cmdTrans(i).Visible = True
            cmdTrans(3).Visible = False
            cmdTrans(4).Visible = False
        Next i
    End If
End Sub

Sub Clearform()

    'Frame 1
    txtSPNo.Text = ""
    txtDesc.Text = ""
    cboClass.Text = ""
    txtBARCODE.Text = ""
    txtROPoint.Text = ""
    cboPMMC.Text = ""
    cboPMModel.Text = ""
    txtMAXPO.Text = ""
    cboSupp.Text = ""
    cboStatus.Text = ""
    cboSuppBrand.Text = ""
    txtApplic.Text = ""
    chkStartSale.Value = 0
    chkEndSale.Value = 0
    txtSTARTSALE.Text = "__/__/____"
    txtENDSALE.Text = "__/__/____"
    cboSubCat.Text = ""
    txtSPid.Text = ""
    txtpmID.Text = ""

    flxSUBSP.Rows = 1
    flxSUBSP.Rows = 2

End Sub

Sub LockData(DataEntry As Boolean)

    DataEntry = Not DataEntry
    'txtSPNo.Locked = DataEntry
    cboClass.Locked = DataEntry
    txtROPoint.Locked = DataEntry
    txtMAXPO.Locked = DataEntry
    cboSupp.Locked = DataEntry
    cboSuppBrand.Locked = DataEntry
    txtDesc.Locked = DataEntry
    txtBARCODE.Locked = DataEntry
    cboPMMC.Locked = DataEntry
    cboPMModel.Locked = DataEntry
    cboStatus.Locked = DataEntry
    txtApplic.Locked = DataEntry
    cboSubCat.Locked = DataEntry
    txtSTARTSALE.Enabled = False
    txtENDSALE.Enabled = False
    chkStartSale.Enabled = False
    chkEndSale.Enabled = False
 
    If cboClass.Locked Then
        'Frame 1
        'txtSPNo.BackColor = txtHoldGrayBackColor.BackColor
        cboClass.BackColor = txtHoldGrayBackColor.BackColor
        txtROPoint.BackColor = txtHoldGrayBackColor.BackColor
        txtMAXPO.BackColor = txtHoldGrayBackColor.BackColor
        cboSupp.BackColor = txtHoldGrayBackColor.BackColor
        cboSuppBrand.BackColor = txtHoldGrayBackColor.BackColor
        txtDesc.BackColor = txtHoldGrayBackColor.BackColor
        txtBARCODE.BackColor = txtHoldGrayBackColor.BackColor
        cboPMMC.BackColor = txtHoldGrayBackColor.BackColor
        cboPMModel.BackColor = txtHoldGrayBackColor.BackColor
        cboStatus.BackColor = txtHoldGrayBackColor.BackColor
        txtApplic.BackColor = txtHoldGrayBackColor.BackColor
        txtSTARTSALE.BackColor = txtHoldGrayBackColor.BackColor
        txtENDSALE.BackColor = txtHoldGrayBackColor.BackColor
        cboSubCat.BackColor = txtHoldGrayBackColor.BackColor
    Else
        chkStartSale.Enabled = False
        chkEndSale.Enabled = False
        'txtSPNo.BackColor = txtHoldWhiteBackColor.BackColor
        cboClass.BackColor = txtHoldWhiteBackColor.BackColor
        txtROPoint.BackColor = txtHoldWhiteBackColor.BackColor
        txtMAXPO.BackColor = txtHoldWhiteBackColor.BackColor
        cboSupp.BackColor = txtHoldWhiteBackColor.BackColor
        cboSuppBrand.BackColor = txtHoldWhiteBackColor.BackColor
        txtDesc.BackColor = txtHoldWhiteBackColor.BackColor
        txtBARCODE.BackColor = txtHoldWhiteBackColor.BackColor
        cboPMMC.BackColor = txtHoldWhiteBackColor.BackColor
        cboPMModel.BackColor = txtHoldWhiteBackColor.BackColor
        cboStatus.BackColor = txtHoldWhiteBackColor.BackColor
        txtApplic.BackColor = txtHoldWhiteBackColor.BackColor
        cboSubCat.BackColor = txtHoldWhiteBackColor.BackColor
    End If
End Sub



Private Sub cboCatalogPath_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then cboSupp.SetFocus
End Sub

Private Sub cboClass_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboClass
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
           
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboClass-Change()")
End Sub

Private Sub cboClass_Click()
Call cboClass_Change
End Sub

Private Sub cboClass_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboClass_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboClass, KeyAscii, ValidChars)
    If KeyAscii = 13 Then cboSubCat.SetFocus
End Sub

Private Sub cboPMMC_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboPMMC
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
           
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboPMMC-Change()")
End Sub

Private Sub cboPMMC_Click()
Call cboPMMC_Change
End Sub

Private Sub cboPMMC_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboPMMC_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboPMMC, KeyAscii, ValidChars)
    If KeyAscii = 13 Then cboPMModel.SetFocus
End Sub

Private Sub cboPMMC_LostFocus()
If formAction = "edit" Or formAction = "add" Then
    Call CboDisplayPMMODEL
End If
End Sub

Private Sub cboPMModel_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboPMModel
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
           
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboPMModel-Change()")
End Sub

Private Sub cboPMModel_Click()
txtpmID.Text = arrPMMODEL(cboPMModel.ListIndex, 0)
End Sub

Private Sub cboPMModel_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboPMModel_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboPMModel, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtMAXPO.SetFocus
End Sub

Private Sub cboStatus_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cboStatus_Click
End Sub

Private Sub cboStatus_LostFocus()
Call cboStatus_Click
End Sub

Private Sub cboSubCat_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboSubCat
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
           
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboSubCat-Change()")
End Sub

Private Sub cboSubCat_Click()
Call cboSubCat_Change
End Sub


Private Sub cboSubCat_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboSubCat_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboSubCat, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtROPoint.SetFocus
End Sub

Private Sub cboSupp_Click()
Call cboSupp_Change
End Sub

Private Sub cboSupp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboSupp_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then cboSuppBrand.SetFocus
End Sub

Private Sub cboSuppBrand_Click()
Call cboSuppBrand_Change
End Sub

Private Sub cboSuppBrand_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboSuppBrand_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtBARCODE.SetFocus
End Sub

Private Sub cmdAltA_Click()
    Call mnuAdd_Click
End Sub

Private Sub cmdAltD_Click()
    Call mnuDelete_Click
End Sub

Private Sub flxMainPartNo_Click()
    With flxMainPartNo
        txtSPNo.Visible = True
        txtSPNo.ZOrder
        PrtSrch = True
        
        txtSPid.Text = .TextMatrix(.row, 1)
        txtSPNo.Text = .TextMatrix(.row, 2)
        txtDesc.Text = .TextMatrix(.row, 3)
        cboSuppBrand.Text = .TextMatrix(.row, 4)
        txtApplic.Text = .TextMatrix(.row, 6)
        cboClass.Text = .TextMatrix(.row, 8)
        txtROPoint.Text = .TextMatrix(.row, 9)
        txtMAXPO.Text = .TextMatrix(.row, 10)
        txtpmID.Text = .TextMatrix(.row, 13)
        cboStatus.Text = .TextMatrix(.row, 14)
        If .TextMatrix(.row, 15) <> "" Then
            txtSTARTSALE.Text = Format(.TextMatrix(.row, 15), "MM/DD/YYYY")
            txtENDSALE.Text = Format(.TextMatrix(.row, 16), "MM/DD/YYYY")
        Else
            txtSTARTSALE.Text = "__/__/____"
            txtENDSALE.Text = "__/__/____"
        End If
        If txtSTARTSALE.Text = "__/__/____" And txtENDSALE.Text = "__/__/____" Then
            chkStartSale.Value = 0
            chkEndSale.Value = 0
        ElseIf txtSTARTSALE.Text <> "" And txtENDSALE.Text = "__/__/____" Then
            chkStartSale.Value = 1
            chkEndSale.Value = 0
        ElseIf txtSTARTSALE.Text <> "" And txtENDSALE.Text <> "" Then
            chkStartSale.Value = 1
            chkEndSale.Value = 1
        End If
        txtBARCODE.Text = .TextMatrix(.row, 17)
        cboSupp.Text = .TextMatrix(.row, 18)
        cboPMMC.Text = .TextMatrix(.row, 19)
        cboPMModel.Text = .TextMatrix(.row, 20)

        txtSPNo.SelStart = Val(txtSPNo.Tag)
        txtSPNo.SelLength = Len(txtSPNo.Text)
    End With

    Call getCatalogPath(txtSPNo.Text)
    Call getSubDetails
    flxMainPartNo.Visible = False
    txtDesc.SetFocus
End Sub


Private Sub flxMainPartNo_EnterCell()
On Error GoTo er
    With flxMainPartNo
        txtSPNo.Visible = True
        txtSPNo.ZOrder
        PrtSrch = True
        
        txtSPid.Text = .TextMatrix(.row, 1)
        txtSPNo.Text = .TextMatrix(.row, 2)
        txtDesc.Text = .TextMatrix(.row, 3)
        cboSuppBrand.Text = .TextMatrix(.row, 4)
        txtApplic.Text = .TextMatrix(.row, 6)
        cboClass.Text = .TextMatrix(.row, 8)
        txtROPoint.Text = .TextMatrix(.row, 9)
        txtMAXPO.Text = .TextMatrix(.row, 10)
        txtpmID.Text = .TextMatrix(.row, 13)
        cboStatus.Text = .TextMatrix(.row, 14)
        If .TextMatrix(.row, 15) <> "" And .TextMatrix(.row, 16) <> "" Then
            txtSTARTSALE.Text = Format(.TextMatrix(.row, 15), "MM/DD/YYYY")
            txtENDSALE.Text = Format(.TextMatrix(.row, 16), "MM/DD/YYYY")
        ElseIf .TextMatrix(.row, 15) <> "" And .TextMatrix(.row, 16) = "" Then
            txtSTARTSALE.Text = Format(.TextMatrix(.row, 15), "MM/DD/YYYY")
            txtENDSALE.Text = "__/__/____"
        Else
            txtSTARTSALE.Text = "__/__/____"
            txtENDSALE.Text = "__/__/____"
        End If
        If txtSTARTSALE.Text = "__/__/____" And txtENDSALE.Text = "__/__/____" Then
            chkStartSale.Value = 0
            chkEndSale.Value = 0
        ElseIf txtSTARTSALE.Text <> "" And txtENDSALE.Text = "__/__/____" Then
            chkStartSale.Value = 1
            chkEndSale.Value = 0
        ElseIf txtSTARTSALE.Text <> "" And txtENDSALE.Text <> "" Then
            chkStartSale.Value = 1
            chkEndSale.Value = 1
        End If
        txtBARCODE.Text = .TextMatrix(.row, 17)
        cboSupp.Text = .TextMatrix(.row, 18)
        cboPMMC.Text = .TextMatrix(.row, 19)
        cboPMModel.Text = .TextMatrix(.row, 20)
        
        txtSPNo.SelStart = Val(txtSPNo.Tag)
        txtSPNo.SelLength = Len(txtSPNo.Text)
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->FlxPartNumber-EnterCell()")
End Sub


Private Sub flxMainPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
       Tmp = Split(txtSPNo.Tag, "|")
       For i = LBound(Tmp) To UBound(Tmp)
           flxMainPartNo.TextMatrix(ActRow, i + 1) = Tmp(i)
       Next i
       flxMainPartNo.Visible = False
       txtDesc.SetFocus
       
    End If
End Sub

Private Sub flxMainPartNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    
    If KeyAscii = 13 Then
        Call flxMainPartNo_Click
    End If
    
End Sub

Private Sub flxMainPartNo_LostFocus()
Call flxMainPartNo_Click
End Sub

Private Sub flxSUBSP_Scroll()
flxPartNo.Visible = False
txtPN.Visible = False
End Sub

Private Sub mnuDelete_Click()
On Error GoTo er
If formAction = "edit" Then
    ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If ans = vbYes Then
        If flxSUBSP.Rows = 2 Then
            flxSUBSP.Rows = 1
            flxSUBSP.Rows = 2
        Else
            flxSUBSP.RemoveItem ActRow
        End If
    End If
End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->mnuDelete()")
End Sub


Private Sub txtApplic_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If cboStatus.Text <> "" Then
            flxSUBSP.SetFocus
            flxSUBSP.col = 1
        Else
            cboStatus.SetFocus
        End If
    End If
End Sub

Private Sub txtBARCODE_GotFocus()
If txtSPNo.Text <> "" Then
    txtRecommended.Visible = True
    txtRecommended.Left = txtBARCODE.Left
    txtRecommended.Top = txtBARCODE.Top + txtBARCODE.Height
    txtRecommended.Width = txtBARCODE.Width
    txtRecommended.Text = "RECOMMENDED BARCODE VALUE: " & txtSPid.Text
Else
    txtRecommended.Text = ""
    txtRecommended.Visible = False
End If
End Sub

Private Sub txtBARCODE_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtApplic.SetFocus
End Sub


Private Sub txtBARCODE_LostFocus()
    txtRecommended.Text = ""
    txtRecommended.Visible = False
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then cboClass.SetFocus
End Sub


Private Sub txtMAXPO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtMAXPO, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then cboCatalogPath.SetFocus
End Sub

Private Sub txtRecommended_DblClick()
If formAction = "edit" Then
    If txtSPid.Text <> "" Then
        txtBARCODE.Text = txtBARCODE.Text & txtSPid.Text
        txtRecommended.Visible = False
    Else
        txtRecommended.Text = ""
        txtRecommended.Visible = False
    End If
End If
End Sub

Private Sub txtRecommended_GotFocus()
If txtSPid.Text <> "" Then
    txtRecommended.Visible = True
    txtRecommended.Text = "RECOMMENDED BARCODE VALUE: " & txtSPid.Text
End If
End Sub

Private Sub txtROPoint_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtROPoint, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then cboPMMC.SetFocus
End Sub

Private Sub txtSPNo_Change()
    Dim Qry As String
    If formAction = "add" Then
        Exit Sub
    Else
        With flxMainPartNo
            If PrtSrch Then
                PrtSrch = False
                Exit Sub
            End If
            
            .Top = txtSPNo.Height + txtSPNo.Top
            .Left = txtSPNo.Left
            .Width = txtSPNo.Width
            .Visible = True
            .ZOrder
            
            If txtSPNo.Text = "" Then .Visible = False: Exit Sub
            rsParts.Find "[PARTNUMBER] LIKE '" + txtSPNo.Text + "%'", , adSearchForward, 1
            If Not rsParts.EOF Then
                If Not PrtBln Then
                    txtSPNo.Tag = Len(txtSPNo.Text)
                    PrtBln = True
                End If
               
                .TopRow = rsParts.AbsolutePosition - 1
                .row = rsParts.AbsolutePosition - 1
                .col = 0: .ColSel = 5
                txtSPNo.Text = rsParts!PartNumber
                txtDesc.Text = rsParts!Description
                
                If IsNull(rsParts!SubCat) Then
                    cboSubCat.Text = ""
                Else
                    cboSubCat.Text = rsParts!SubCat
                End If
                cboClass.Text = IIf(IsNull(rsParts!Remarks), "", rsParts!Remarks)
                txtROPoint.Text = IIf(IsNull(rsParts!REORDERPOINT), 0, rsParts!REORDERPOINT)
                txtMAXPO.Text = IIf(IsNull(rsParts!MAXALLOW), 0, rsParts!MAXALLOW)
                txtApplic.Text = IIf(IsNull(rsParts!ApplicModel), "", rsParts!ApplicModel)
                cboSupp.Text = IIf(IsNull(rsParts!SUPPLIERNAME), "", rsParts!SUPPLIERNAME)
                cboSuppBrand.Text = IIf(IsNull(rsParts!Description2), "", rsParts!Description2)
                cboPMMC.Text = IIf(IsNull(rsParts!PMBRAND), "", rsParts!PMBRAND)
                If cboPMMC.Text <> "" Then
                    Call CboDisplayPMMODEL
                End If
                cboPMModel.Text = IIf(IsNull(rsParts!PMMODEL), "", rsParts!PMMODEL)
                txtBARCODE.Text = IIf(IsNull(rsParts!BCODEVALUE), "", rsParts!BCODEVALUE)
                cboStatus.Text = IIf(IsNull(rsParts!Status), "", rsParts!Status)
                
                txtSTARTSALE.Text = Format(IIf(IsNull(rsParts!SALESTART), "__/__/____", rsParts!SALESTART), "mm/dd/yyyy")
                txtENDSALE.Text = Format(IIf(IsNull(rsParts!SALEEND), "__/__/____", rsParts!SALEEND), "mm/dd/yyyy")
                
                If txtSTARTSALE.Text = "__/__/____" And txtENDSALE.Text = "__/__/____" Then
                    chkStartSale.Value = 0
                    chkEndSale.Value = 0
                ElseIf txtSTARTSALE.Text <> "" And txtENDSALE.Text = "__/__/____" Then
                    chkStartSale.Value = 1
                    chkEndSale.Value = 0
                ElseIf txtSTARTSALE.Text <> "" And txtENDSALE.Text <> "" Then
                    chkStartSale.Value = 1
                    chkEndSale.Value = 1
                End If
                
                txtSPid.Text = IIf(IsNull(Val(rsParts!SPAREPARTID)), "", rsParts!SPAREPARTID)
                txtpmID.Text = IIf(IsNull(Val(rsParts!PMATRIXID)), "", rsParts!PMATRIXID)
                
                txtSPNo.SelStart = Val(txtSPNo.Tag)
                txtSPNo.SelLength = Len(txtSPNo.Text)
                PrtBln = False
            Else
                'txtSPNo.Text = ""
                txtDesc.Text = ""
                cboClass.Text = ""
                txtBARCODE.Text = ""
                txtROPoint.Text = ""
                cboPMMC.Text = ""
                cboPMModel.Text = ""
                txtMAXPO.Text = ""
                cboSupp.Text = ""
                cboStatus.Text = ""
                cboSuppBrand.Text = ""
               ' cboSubCat.Text = ""
                txtApplic.Text = ""
                chkStartSale.Value = 0
                chkEndSale.Value = 0
                txtSTARTSALE.Text = "__/__/____"
                txtENDSALE.Text = "__/__/____"
                txtSPid.Text = ""
                txtpmID.Text = ""
                .Visible = False
            End If
        End With
    End If
End Sub

Private Sub txtSPNo_Click()
Call txtSPNo_Change
txtSPid.Text = 0
'txtSPID.Text = arrSPNO(txtSPNo.ListIndex, 0)
Call txtSPNo_LostFocus
End Sub

Private Sub txtSPNo_GotFocus()
    If rsParts.State = adStateClosed Then tmrCboDisplaySP.Enabled = True
    If txtSPNo.Text <> "" Then
        txtSPNo.SelStart = 0
        txtSPNo.SelLength = Len(txtSPNo.Text)
    End If
End Sub

Private Sub txtSPNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        flxMainPartNo.Tag = "skiplostfocus"
        flxMainPartNo.Visible = True
        flxMainPartNo.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        txtSPNo.SetFocus
        flxMainPartNo.Visible = False
    ElseIf KeyCode = vbKeyDelete Then
        PrtSrch = True
        rsParts.MoveFirst
        txtSPNo.Text = ""
        txtDesc.Text = ""
        cboClass.Text = ""
        txtBARCODE.Text = ""
        txtROPoint.Text = ""
        cboPMMC.Text = ""
        cboPMModel.Text = ""
        txtMAXPO.Text = ""
        cboSupp.Text = ""
        cboStatus.Text = ""
        cboSuppBrand.Text = ""
        'cboSubCat.Text = ""
        txtApplic.Text = ""
        chkStartSale.Value = 0
        chkEndSale.Value = 0
        txtSTARTSALE.Text = "__/__/____"
        txtENDSALE.Text = "__/__/____"
        txtSPid.Text = ""
        txtpmID.Text = ""
    ElseIf KeyCode = vbKeyBack Then
        PrtSrch = True
        'rsParts.MoveFirst
        'flxMainPartNo.Visible = False
    End If
End Sub

Public Sub getSubDetails()
Dim rsSPdtls As New Recordset

    DoEvents
    Qry = "call sp_mrcwarehouse.sp_SPProperties_getSubSPDetails(" & Val(txtSPid.Text) & "); "
    With rsSPdtls
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxSUBSP.Rows = .RecordCount + 1
            .MoveFirst
            For i = 1 To flxSUBSP.Rows - 1
                flxSUBSP.TextMatrix(i, 1) = IIf(IsNull(!PartNumber), "", !PartNumber)
                flxSUBSP.TextMatrix(i, 2) = IIf(IsNull(!Description), "", !Description)
                flxSUBSP.TextMatrix(i, 3) = IIf(IsNull(!CLASSIFICATION), "", !CLASSIFICATION)
                flxSUBSP.TextMatrix(i, 4) = IIf(IsNull(!SUPPLIERNAME), "", !SUPPLIERNAME)
                flxSUBSP.TextMatrix(i, 5) = IIf(IsNull(!SUPPLIERBRAND), "", !SUPPLIERBRAND)
                flxSUBSP.TextMatrix(i, 6) = IIf(IsNull(!PMMODEL), "", !PMMODEL)
                flxSUBSP.TextMatrix(i, 7) = IIf(IsNull(!SPAREPARTID), "", !SPAREPARTID)
                flxSUBSP.TextMatrix(i, 8) = IIf(IsNull(!PMATRIXID), "", !PMATRIXID)
                .MoveNext
            Next i
        Else
            flxSUBSP.Rows = 1
            flxSUBSP.Rows = 2
        End If
        .Close
    End With
    
End Sub

Private Sub txtSPNo_KeyPress(KeyAscii As Integer)
Dim rsSPdtls As New Recordset
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(txtSPNo, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Call txtSPNo_LostFocus
        flxMainPartNo.Visible = False
        Call getSubDetails
        Call getCatalogPath(txtSPNo.Text)
        txtDesc.SetFocus
    End If
End Sub

Private Sub txtSPNo_LostFocus()
    If flxMainPartNo.Tag = "skiplostfocus" Then
        flxMainPartNo.Tag = ""
        Exit Sub
    End If
    flxMainPartNo.Visible = False
End Sub

Public Sub getCatalogPath(myPartNum As String)
    If ConnectPartsCatalog = True Then
        Qry = "SELECT Cast(CONCAT(pc.cat_Brand,'|',pc.cat_Model,'|',pc.cat_Version,'|',pp.category_Group,'|',pp.category_Block,'|',ps.parts_Refno" & _
              ") as char) AS `MOTORCYCLE` From spareparts AS ps INNER JOIN partscategory AS pp ON pp.category_ID = ps.category_ID " & _
              "INNER JOIN catalog AS pc ON pp.cat_ID = pc.cat_ID WHERE ps.parts_Item = '" & Trim(myPartNum) & "';"
        With rsAnyUse
            If .State = adStateOpen Then .Close
            .Open Qry, conPartsCat, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                cboCatalogPath.Clear
                cboCatalogPath.Text = IIf(IsNull(!MOTORCYCLE), "", !MOTORCYCLE)
                Do While Not .EOF
                    cboCatalogPath.AddItem IIf(IsNull(!MOTORCYCLE), "", !MOTORCYCLE)
                    .MoveNext
                Loop
            Else
                cboCatalogPath.Clear
            End If
            .Close
        End With
    Else
        MsgBox "Cannot Connect To Parts Catalog Database!", vbInformation + vbOKOnly, "Validation"
        Exit Sub
    End If
End Sub
Private Sub cboStatus_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboStatus
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
           
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboStatus-Change()")
End Sub

Private Sub cboStatus_Click()
    Call cboStatus_Change
    If cboStatus.Text = "SALES ITEM" Then
        chkStartSale.Enabled = True
        chkEndSale.Enabled = True
        'txtSTARTSALE.Enabled = True
        'txtSTARTSALE.BackColor = txtHoldWhiteBackColor.BackColor
        'txtENDSALE.Enabled = True
        'txtENDSALE.BackColor = txtHoldWhiteBackColor.BackColor
    Else
        chkStartSale.Value = 0
        chkStartSale.Enabled = False
        txtSTARTSALE.Enabled = False
        txtSTARTSALE.BackColor = txtHoldGrayBackColor.BackColor
        txtSTARTSALE.Text = "__/__/____"
        
        chkEndSale.Value = 0
        chkEndSale.Enabled = False
        txtENDSALE.Enabled = False
        txtENDSALE.BackColor = txtHoldGrayBackColor.BackColor
        txtENDSALE.Text = "__/__/____"
    End If
    
    cmbstr = cboStatus.Text
    
    For i = 0 To cboStatus.ListCount - 1
        If UCase$(cboStatus.List(i)) = UCase$(cmbstr) Then
            cboStatus.ListIndex = i
            blnFound = True
            Exit For
        End If
    Next i
    
    If Not blnFound Then
        MsgBox "Item was not found in the list.", _
        vbExclamation, _
        "Warning"
        cboStatus.SetFocus
        cboStatus.Text = ""
        Exit Sub
    End If
    
    blnFound = False
    txtApplic.SetFocus
End Sub


Private Sub cboSupp_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboSupp
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
           
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboSupp-Change()")
End Sub

Private Sub cboSuppBrand_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboSuppBrand
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
           
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboSuppBrand-Change()")
End Sub

Private Sub chkEndSale_Click()
If chkEndSale.Value = 1 Then
    If formAction = "edit" Or formAction = "add" Then
        chkEndSale.Value = 1
        chkEndSale.Enabled = True
        txtENDSALE.Enabled = True
        txtENDSALE.BackColor = txtHoldWhiteBackColor.BackColor
        txtENDSALE.Text = Format$(ServerTime, "mm/dd/yyyy")
    End If
Else
    chkEndSale.Value = 0
    chkEndSale.Enabled = False
    txtENDSALE.Enabled = False
    txtENDSALE.BackColor = txtHoldGrayBackColor.BackColor
    txtENDSALE.Text = "__/__/____"
End If
End Sub

Private Sub chkStartSale_Click()
If chkStartSale.Value = 1 Then
    If formAction = "edit" Or formAction = "add" Then
        chkStartSale.Value = 1
        chkStartSale.Enabled = True
        txtSTARTSALE.Enabled = True
        txtSTARTSALE.BackColor = txtHoldWhiteBackColor.BackColor
        txtSTARTSALE.Text = Format$(ServerTime, "mm/dd/yyyy")
    End If
Else
    chkStartSale.Value = 0
    chkStartSale.Enabled = False
    txtSTARTSALE.Enabled = False
    txtSTARTSALE.BackColor = txtHoldGrayBackColor.BackColor
    txtSTARTSALE.Text = "__/__/____"
End If
End Sub

Sub CreateTempTable()
On Error GoTo er
        LYRcon.Execute "Drop Temporary Table IF  Exists sp_mrcwarehouse.TMP_SUBSPAREPARTS;"

         DoEvents
       LYRcon.Execute " Create Temporary Table IF Not Exists sp_mrcwarehouse.TMP_SUBSPAREPARTS( " & _
        "    SPAREPARTID Int(11)        " & _
        "   ,SUBSPID Int(11)            " & _
        "   ,PMATRIXID Int(11)          " & _
        "   ,SEQNO Int(11)              " & _
        ");"
        
        Qry = "Insert into sp_mrcwarehouse.TMP_SUBSPAREPARTS(SPAREPARTID,SUBSPID,PMATRIXID,SEQNO) VALUES "
        For i = 1 To flxSUBSP.Rows - 1
            If i > 1 Then Qry = Qry & ","
            Qry = Qry & "(" & Val(txtSPid.Text) & "," & Val(flxSUBSP.TextMatrix(i, 7)) & "," & Val(flxSUBSP.TextMatrix(i, 8)) & ", '" & i & "')"
        Next i
       
       Call LYRcon.Execute(Qry)
       
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CreateTempTable()")
End Sub

Function ValidationSubSP() As String
On Error GoTo er
   
    MyErr = ""
    If flxSUBSP.TextMatrix(flxSUBSP.Rows - 1, 1) = "" Then
        If flxSUBSP.Rows = 2 Then
            flxSUBSP.Rows = 1
            flxSUBSP.Rows = 2
        Else
            flxSUBSP.RemoveItem flxSUBSP.Rows - 1
        End If
    End If
    
    If flxSUBSP.Rows > 2 Then
        If flxSUBSP.TextMatrix(flxSUBSP.Rows - 1, 1) = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on Last Row"
    End If
    ValidationSubSP = MyErr


er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->ValidationPartsMatrix()")
End Function
Private Sub cmdTrans_Click(Index As Integer)
On Error GoTo er

    If Index = 0 Then 'add
        Call LockData(True)
        Call SetButtonVisibility(True)
        formAction = "add"
        Call chkStartSale_Click
        Call chkEndSale_Click
        Call Clearform
    ElseIf Index = 1 Then 'edit
        Call LockData(True)
        Call SetButtonVisibility(True)
        formAction = "edit"
        Call chkStartSale_Click
        Call chkEndSale_Click
    ElseIf Index = 2 Then 'exit
        Call Clearform
        formAction = ""
        Unload Me
    ElseIf Index = 3 Then 'save

        myMsg = ValidationSubSP
        
        If myMsg <> "" Then
            MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
            Exit Sub
        End If
    
        ans = MsgBox("Do you want to save?", vbQuestion + vbYesNo, "Save")
        If ans = vbNo Then Exit Sub
        
        Call CreateTempTable
        If formAction = "add" Then
            If cboStatus.Text <> "SALES ITEM" Then
                DoEvents
                Qry = "call sp_mrcwarehouse.sp_SPProperties_create('" & Trim(txtSPNo.Text) & "','" & Trim(txtDesc.Text) & "','" & Trim(cboClass.Text) & "','" & Trim(txtBARCODE.Text) & "'," & Val(txtROPoint.Text) & "," & Val(txtpmID.Text) & "," & Val(txtMAXPO.Text) & ",'" & Trim(cboSupp.Text) & "','" & Trim(cboStatus.Text) & "','" & Trim(cboSuppBrand.Text) & "','" & Trim(txtApplic.Text) & "','" & "" & "','" & "" & "','" & Trim(cboSubCat.Text) & "')"
            Else
                DoEvents
                Qry = "call sp_mrcwarehouse.sp_SPProperties_create('" & Trim(txtSPNo.Text) & "','" & Trim(txtDesc.Text) & "','" & Trim(cboClass.Text) & "','" & Trim(txtBARCODE.Text) & "'," & Val(txtROPoint.Text) & "," & Val(txtpmID.Text) & "," & Val(txtMAXPO.Text) & ",'" & Trim(cboSupp.Text) & "','" & Trim(cboStatus.Text) & "','" & Trim(cboSuppBrand.Text) & "','" & Trim(txtApplic.Text) & "','" & Format(txtSTARTSALE.Text, "YYYY-MM-DD") & "','" & Format(txtENDSALE.Text, "YYYY-MM-DD") & "','" & Trim(cboSubCat.Text) & "')"
            End If
        ElseIf formAction = "edit" Then
            If cboStatus.Text <> "SALES ITEM" Then
                DoEvents
                Qry = "call sp_mrcwarehouse.sp_SPProperties_update(" & Val(txtSPid.Text) & ",'" & Trim(txtDesc.Text) & "','" & Trim(cboClass.Text) & "','" & Trim(txtBARCODE.Text) & "'," & Val(txtROPoint.Text) & "," & Val(txtpmID.Text) & "," & Val(txtMAXPO.Text) & ",'" & Trim(cboSupp.Text) & "','" & Trim(cboStatus.Text) & "','" & Trim(cboSuppBrand.Text) & "','" & Trim(txtApplic.Text) & "','" & "" & "','" & "" & "','" & Trim(cboSubCat.Text) & "')"
            Else
                DoEvents
                Qry = "call sp_mrcwarehouse.sp_SPProperties_update(" & Val(txtSPid.Text) & ",'" & Trim(txtDesc.Text) & "','" & Trim(cboClass.Text) & "','" & Trim(txtBARCODE.Text) & "'," & Val(txtROPoint.Text) & "," & Val(txtpmID.Text) & "," & Val(txtMAXPO.Text) & ",'" & Trim(cboSupp.Text) & "','" & Trim(cboStatus.Text) & "','" & Trim(cboSuppBrand.Text) & "','" & Trim(txtApplic.Text) & "','" & Format(txtSTARTSALE.Text, "YYYY-MM-DD") & "','" & Format(txtENDSALE.Text, "YYYY-MM-DD") & "','" & Trim(cboSubCat.Text) & "')"
            End If
        End If
        With rsMain
            If .State = adStateOpen Then .Close
            .Open Qry, LYRcon, adOpenKeyset, adLockPessimistic
            
            If .RecordCount > 0 Then
                If !level = "Granted" Then
                    MsgBox "" & !message & "", vbInformation + vbOKOnly, "" & !level & ""
                    If formAction = "add" Then
                        txtSPid.Text = !SPid
                    End If
                Else
                    MsgBox "" & !message & "", vbInformation + vbOKOnly, "" & !level & ""
                    Exit Sub
                End If
            End If
        End With
        Call RefreshRecord
        Call SetButtonVisibility(False)
        Call LockData(False)
        formAction = ""
    ElseIf Index = 4 Then 'cancel
    
        ans = MsgBox("Do you want to cancel?", vbQuestion + vbYesNo)
        If ans = vbNo Then Exit Sub
        Call SetButtonVisibility(False)
        Call LockData(False)
        'chkFinal.Enabled = False
        formAction = ""
        
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdTrans-Click()")
End Sub

Private Sub flxPartNo_Click()
    With flxPartNo
        txtPN.Visible = True
        txtPN.ZOrder
        PrtSrch = True
        
        txtPN.Text = flxPartNo.TextMatrix(flxPartNo.row, 1)
        
        flxSUBSP.TextMatrix(flxSUBSP.row, 2) = flxPartNo.TextMatrix(flxPartNo.row, 2)
        flxSUBSP.TextMatrix(flxSUBSP.row, 3) = flxPartNo.TextMatrix(flxPartNo.row, 3)
        flxSUBSP.TextMatrix(flxSUBSP.row, 4) = flxPartNo.TextMatrix(flxPartNo.row, 4)
        flxSUBSP.TextMatrix(flxSUBSP.row, 5) = flxPartNo.TextMatrix(flxPartNo.row, 5)
        flxSUBSP.TextMatrix(flxSUBSP.row, 7) = flxPartNo.TextMatrix(flxPartNo.row, 7)
        txtPN.SelStart = Val(txtPN.Tag)
        txtPN.SelLength = Len(txtPN.Text)
    End With
    
    If txtPN.Text <> "" Then
        If ActRow = flxSUBSP.Rows - 1 Then flxSUBSP.AddItem ""
        flxSUBSP.SetFocus
        flxSUBSP.row = flxSUBSP.Rows - 1
        flxSUBSP.col = 1
    End If
    'Call txtPartNo_LostFocus
    txtPN.Visible = False
    flxPartNo.Visible = False
    flxSUBSP.SetFocus
    'CCol = 1
End Sub

Private Sub flxPartNo_EnterCell()
On Error GoTo er
    With flxPartNo
        txtPN.Visible = True
        txtPN.ZOrder
        PrtSrch = True
        txtPN.Text = flxPartNo.TextMatrix(flxPartNo.row, 1)
        flxSUBSP.TextMatrix(flxSUBSP.row, 2) = flxPartNo.TextMatrix(flxPartNo.row, 2)
        flxSUBSP.TextMatrix(flxSUBSP.row, 3) = flxPartNo.TextMatrix(flxPartNo.row, 3)
        flxSUBSP.TextMatrix(flxSUBSP.row, 4) = flxPartNo.TextMatrix(flxPartNo.row, 4)
        flxSUBSP.TextMatrix(flxSUBSP.row, 5) = flxPartNo.TextMatrix(flxPartNo.row, 5)
        flxSUBSP.TextMatrix(flxSUBSP.row, 7) = flxPartNo.TextMatrix(flxPartNo.row, 7)
        txtPN.SelStart = Val(txtPN.Tag)
        txtPN.SelLength = Len(txtPN.Text)
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->FlxPartNumber-EnterCell()")
End Sub

Private Sub flxPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
       Tmp = Split(txtPN.Tag, "|")
       For i = LBound(Tmp) To UBound(Tmp)
           flxPartNo.TextMatrix(ActRow, i + 1) = Tmp(i)
       Next i
       txtPN.Visible = False
       flxPartNo.Visible = False
       flxSUBSP.SetFocus
       flxSUBSP.col = 1
       
    End If
End Sub

Private Sub flxPartNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    
    If KeyAscii = 13 Then
        Call flxPartNo_Click
    End If
    
End Sub

Private Sub flxPartNo_LostFocus()
Call txtPN_LostFocus
End Sub

Private Sub flxSUBSP_Click()
    ActRow = flxSUBSP.row
    ActCol = flxSUBSP.col

    If ActCol = 0 Or ActRow = 0 Then Exit Sub
    Select Case ActCol
        Case 1  'PartNumber
            If formAction <> "" Then
                If txtSPNo.Text <> "" Then
                    With txtPN
                        If flxSUBSP.TextMatrix(ActRow, 1) <> "" Then .Tag = flxSUBSP.TextMatrix(ActRow, 1)
                        .Top = Frame2.Top + flxSUBSP.CellTop + flxSUBSP.Top
                        .Left = Frame2.Left + flxSUBSP.CellLeft + flxSUBSP.Left
                        .Height = flxSUBSP.CellHeight
                        .Width = flxSUBSP.CellWidth
                        .Text = flxSUBSP.Text
                        .Visible = True
                        .SelStart = 0
                        .SelLength = Len(.Text)
                        .ZOrder
                        .SetFocus
                    End With
                End If
            Else
                Exit Sub
            End If
       End Select
End Sub

Private Sub flxSUBSP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call flxSUBSP_Click
    End If
End Sub

Private Sub flxSUBSP_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If formAction = "add" Then
            PopupMenu mnuStatus, vbPopupMenuRightButton
        ElseIf formAction = "edit" Then
            PopupMenu mnuStatus, vbPopupMenuRightButton
        ElseIf formAction = "summary" Then
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Load()
    Call LockData(False)
    '===================BUTTON POSITION===================
    If cmdTrans(0).Visible = False Then
        'Align Save Button
        cmdTrans(3).Left = cmdTrans(0).Left
        cmdTrans(3).Top = cmdTrans(0).Top
        'Align Cancel Button
        cmdTrans(4).Left = cmdTrans(1).Left
        cmdTrans(4).Top = cmdTrans(1).Top
    End If
    
   
    Warehouse.sBar.Panels(6).Text = "Updating Spare Parts"
End Sub



Public Sub CboDisplayPMMODEL()
On Error GoTo er
    'Get Data from database
    DoEvents
    Qry = "call sp_mrcwarehouse.sp_SPProperties_getPMMODEL('" & cboPMMC.Text & "')"
    If rsPMMODEL.State = adStateOpen Then rsPMMODEL.Close
    rsPMMODEL.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsPMMODEL.State = 1 Then
        Call modObjTools.addToListObjectWithArrayGivenField(cboPMModel, rsPMMODEL, arrPMMODEL, "PMATRIXID", "PMMODEL")
    End If
    If rsPMMODEL.State <> 0 Then rsPMMODEL.Close
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplayPMMODEL()")
End Sub

Private Sub mnuAdd_Click()
If formAction = "edit" Then
    If flxSUBSP.TextMatrix(flxSUBSP.Rows - 1, 7) = "" Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        Exit Sub
    Else
        flxSUBSP.AddItem ""
        flxSUBSP.SetFocus
        flxSUBSP.row = flxSUBSP.Rows - 1
        flxSUBSP.col = 1
    End If
End If
End Sub

Private Sub RXResizer1_ResizeComplete()
Call flxChanger
End Sub

Private Sub tmrCboDisplay_Timer()
On Error GoTo er
    If tmrCboDisplay.Enabled = False Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    Qry = "call sp_mrcwarehouse.sp_SPProperties_getClassification()"
    With rsPType
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            cboClass.Clear
            .MoveFirst
            Do While Not .EOF
                cboClass.AddItem !CLASSIFICATION
                .MoveNext
            Loop
        End If
        .Close
    End With
        
    cboStatus.AddItem ""
    cboStatus.AddItem "INACTIVE"
    cboStatus.AddItem "INTER BRANCH TRANSFER"
    cboStatus.AddItem "SALES ITEM"
    cboStatus.AddItem "SPECIAL ORDER"
    cboStatus.AddItem "WAREHOUSE INVENTORY ONLY"
    
    cboStatus.ListIndex = 0
    cboStatus.Text = cboStatus.List(0)
    
    Qry = "call sp_mrcwarehouse.sp_SPProperties_getSupplier()"
    With rsSupplier
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboSupp.Clear
            Do While Not .EOF
                cboSupp.AddItem !SUPPLIERNAME
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    Qry = "call sp_mrcwarehouse.sp_SPProperties_getPMBrand()"
    With rsPMBRAND
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboPMMC.Clear
            Do While Not .EOF
                cboPMMC.AddItem !BRAND
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    Qry = "call sp_mrcwarehouse.sp_SPProperties_getSuppBrand()"
    With rsSUPPBRAND
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboSuppBrand.Clear
            Do While Not .EOF
                cboSuppBrand.AddItem !Description2
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    Qry = "call sp_mrcwarehouse.sp_SPProperties_getSubCategory()"
    With rsSUBCAT
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboSubCat.Clear
            Do While Not .EOF
                cboSubCat.AddItem !SubCat
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    tmrCboDisplay.Enabled = False
    Screen.MousePointer = Default
Exit Sub
er:
    tmrCboDisplay.Enabled = False
    Screen.MousePointer = Default
End Sub

'
Private Sub tmrCboDisplaySP_Timer()
On Error GoTo er
    If tmrCboDisplaySP.Enabled = False Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    DoEvents
    Qry = "call sp_mrcwarehouse.sp_SPProperties_getSpareparts()"
    If rsParts.State = adStateOpen Then rsParts.Close
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    
    With rsParts
        flxMainPartNo.Clear
        Set flxMainPartNo.DataSource = rsParts
    End With
    
    With flxMainPartNo
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = txtSPNo.Width
        For i = 3 To 21
            .ColWidth(i) = 0
        Next i
    End With
    txtSPNo.SetFocus
    tmrCboDisplaySP.Enabled = False
    Screen.MousePointer = Default
Exit Sub
er:
    tmrCboDisplaySP.Enabled = False
    Screen.MousePointer = Default
End Sub

Private Sub tmrSUBSP_Timer()
On Error GoTo er
    DoEvents
    Qry = "Call sp_mrcwarehouse.sp_SPProperties_getSubSpareparts()"
    
    With rsSubParts
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    
        Set flxPartNo.DataSource = rsSubParts
        flxPartNo.Refresh
        With flxPartNo
            .Width = flxSUBSP.ColWidth(1)
            .ColWidth(0) = 0
            .ColWidth(1) = flxSUBSP.ColWidth(1)
            .ColWidth(2) = 0
            .ColWidth(3) = 0
            .ColWidth(4) = 0
            .ColWidth(5) = 0
            .ColWidth(6) = 0
            .ColWidth(7) = 0
            .ColWidth(8) = 0
            .ColWidth(9) = 0
            .ColWidth(10) = 0
        End With
    End With
    tmrSUBSP.Enabled = False
    
Exit Sub
er:
    tmrCboDisplaySP.Enabled = False
    Screen.MousePointer = Default
End Sub


Private Sub txtPN_Change()
On Error GoTo LocalErr
    With flxPartNo
        If PrtSrch Then
            PrtSrch = False
            Exit Sub
        End If
        .Top = txtPN.Height + txtPN.Top
        .Left = txtPN.Left
         .Width = txtPN.Width
        .Visible = True
        .ZOrder
            
        If txtPN.Text = "" Then .Visible = False: Exit Sub
        rsSubParts.Find "[PARTNUMBER] LIKE '" + txtPN.Text + "%'", , adSearchForward, 1
        If Not rsSubParts.EOF Then
            If Not PrtBln Then
                txtPN.Tag = Len(txtPN.Text)
                PrtBln = True
            End If
            
            .TopRow = rsSubParts.AbsolutePosition - 1
            .row = rsSubParts.AbsolutePosition - 1
            .col = 0: .ColSel = 1
            txtPN.Text = rsSubParts!PartNumber
            flxSUBSP.TextMatrix(flxSUBSP.row, 2) = rsSubParts!Description
            flxSUBSP.TextMatrix(flxSUBSP.row, 3) = IIf(IsNull(rsSubParts!Remarks), "", rsSubParts!Remarks)
            flxSUBSP.TextMatrix(flxSUBSP.row, 4) = IIf(IsNull(rsSubParts!SUPPLIERNAME), "", rsSubParts!SUPPLIERNAME)
            flxSUBSP.TextMatrix(flxSUBSP.row, 5) = IIf(IsNull(rsSubParts!Description2), "", rsSubParts!Description2)
            flxSUBSP.TextMatrix(flxSUBSP.row, 6) = IIf(IsNull(rsSubParts!PMMODEL), "", rsSubParts!PMMODEL)
            flxSUBSP.TextMatrix(flxSUBSP.row, 7) = rsSubParts!SPAREPARTID
            flxSUBSP.TextMatrix(flxSUBSP.row, 8) = IIf(IsNull(rsSubParts!PMATRIXID), "", rsSubParts!PMATRIXID)
            txtPN.SelStart = Val(txtPN.Tag)
            txtPN.SelLength = Len(txtPN.Text)
            PrtBln = False
        Else
    
            flxSUBSP.TextMatrix(ActRow, 1) = ""
            flxSUBSP.TextMatrix(ActRow, 2) = ""
            flxSUBSP.TextMatrix(ActRow, 3) = ""
            flxSUBSP.TextMatrix(ActRow, 4) = ""
            flxSUBSP.TextMatrix(ActRow, 5) = ""
            flxSUBSP.TextMatrix(ActRow, 6) = ""
            flxSUBSP.TextMatrix(ActRow, 7) = ""
            flxSUBSP.TextMatrix(ActRow, 8) = ""
            .Visible = False
        End If
   End With
   
   Exit Sub

LocalErr:
    If Err.Number = 30009 Then
        flxPartNo.Visible = False
        txtPN.Visible = False
        If rsSubParts.State = adStateOpen Then rsSubParts.Close
        tmrSUBSP.Enabled = True
        MsgBox "Program is reloading Part Number" & vbCrLf & "Please wait for a few seconds...", vbInformation + vbOKOnly, "Reloading"
    Else
        flxPartNo.Visible = False
        txtPN.Visible = False
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
    End If
End Sub

Private Sub txtPN_GotFocus()
    If rsSubParts.State = adStateClosed Then tmrSUBSP.Enabled = True
    With flxPartNo
        .Top = txtPN.Height + txtPN.Top
        .Left = txtPN.Left
        .Width = txtPN.Width
    End With
End Sub


Private Sub txtPN_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        txtPN.Tag = "skiplostfocus"
        flxPartNo.Visible = True
        flxPartNo.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        flxSUBSP.SetFocus
        flxSUBSP.col = 1
        txtPN.Visible = False
        flxPartNo.Visible = False
    ElseIf KeyCode = vbKeyDelete Then
        PrtSrch = True
        rsSubParts.MoveFirst
        flxSUBSP.TextMatrix(flxSUBSP.row, 7) = ""
        flxSUBSP.TextMatrix(flxSUBSP.row, 8) = ""
    ElseIf KeyCode = vbKeyBack Then
        PrtSrch = True
        rsSubParts.MoveFirst
        flxSUBSP.TextMatrix(flxSUBSP.row, 2) = ""
        flxSUBSP.TextMatrix(flxSUBSP.row, 3) = ""
        flxSUBSP.TextMatrix(flxSUBSP.row, 4) = ""
        flxSUBSP.TextMatrix(flxSUBSP.row, 5) = ""
        flxSUBSP.TextMatrix(flxSUBSP.row, 6) = ""
        flxSUBSP.TextMatrix(flxSUBSP.row, 7) = ""
        flxSUBSP.TextMatrix(flxSUBSP.row, 8) = ""
        flxPartNo.Visible = False
    End If
End Sub


Public Sub ChckDuplicatePN()

On Error GoTo er
'Verify the Part Number if already exist
        If txtPN.Text = flxSUBSP.TextMatrix(flxSUBSP.row, 1) Then
            Exit Sub
        End If
    
    With flxSUBSP
        If ActRow = flxSUBSP.Rows - 1 Then
            If txtSPNo.Text <> txtPN.Text Then
                For i = 1 To .Rows - 1
                    If flxSUBSP.TextMatrix(i, 1) = txtPN.Text Then
                        MsgBox "Part Number is already exist in row " & " " & i & "" & Err.Description, vbExclamation + vbOKOnly, "Warning"
                        txtPN.Visible = False
                        flxPartNo.Visible = False
                        .TextMatrix(.row, 1) = ""
                        .TextMatrix(.row, 2) = ""
                        .TextMatrix(.row, 3) = ""
                        .TextMatrix(.row, 4) = ""
                        .TextMatrix(.row, 5) = ""
                        .TextMatrix(.row, 6) = ""
                        .TextMatrix(.row, 7) = ""
                        .TextMatrix(.row, 8) = ""
                        txtPN.Text = ""
                        Exit Sub
                    End If
                Next i
            Else
                MsgBox "Part Number is already selected! Please Entry another Part Number " & Err.Description, vbExclamation + vbOKOnly, "Warning"
                txtPN.Visible = False
                flxPartNo.Visible = False
                .TextMatrix(.row, 1) = ""
                .TextMatrix(.row, 2) = ""
                .TextMatrix(.row, 3) = ""
                .TextMatrix(.row, 4) = ""
                .TextMatrix(.row, 5) = ""
                .TextMatrix(.row, 6) = ""
                .TextMatrix(.row, 7) = ""
                .TextMatrix(.row, 8) = ""
                txtPN.Text = ""
                Exit Sub
            End If
        End If
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CheckDuplicate()")
End Sub


Private Sub txtPN_KeyPress(KeyAscii As Integer)
On Error GoTo er
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(txtPARTNO, KeyAscii, ValidChars)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If rsSubParts.State = adStateClosed Then tmrSUBSP.Enabled = True
     
    If KeyAscii = 13 Then
        Call ChckDuplicatePN
        If flxSUBSP.TextMatrix(ActRow, 7) = "" Then   'No SPID
            flxSUBSP.SetFocus
            flxSUBSP.row = ActRow
            flxSUBSP.col = 1
        Else
            If ActRow = flxSUBSP.Rows - 1 Then flxSUBSP.AddItem ""
            flxSUBSP.SetFocus
            flxSUBSP.row = flxSUBSP.Rows - 1
            flxSUBSP.col = 1
        End If
        'flxPARTNO.Visible = False
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtPN_KeyPress()")
End Sub

Private Sub txtPN_LostFocus()

On Error GoTo er
    If txtPN.Tag = "skiplostfocus" Then
        txtPN.Tag = ""
        Exit Sub
    End If
    Call ChckDuplicatePN
    If txtPN.Text <> "" Then
        If flxSUBSP.TextMatrix(ActRow, 7) = "" Or txtPN.Text = "" Then
            MsgBox "Unrecognized Part Number!" & vbCrLf & "Please select a Part Number from list.", vbExclamation + vbOKOnly, "Warning"
            txtPN.Visible = False
            flxPartNo.Visible = False
            flxSUBSP.TextMatrix(ActRow, 1) = ""
            flxSUBSP.TextMatrix(ActRow, 2) = ""
            flxSUBSP.TextMatrix(ActRow, 3) = ""
            flxSUBSP.TextMatrix(ActRow, 4) = ""
            flxSUBSP.TextMatrix(ActRow, 5) = ""
            flxSUBSP.TextMatrix(ActRow, 6) = ""
            flxSUBSP.TextMatrix(ActRow, 7) = ""
            flxSUBSP.TextMatrix(ActRow, 8) = ""
            flxSUBSP.SetFocus
            flxSUBSP.col = 1
        Exit Sub
        End If
    End If
    
        flxSUBSP.TextMatrix(ActRow, 1) = txtPN.Text
        txtPN.Visible = False
        flxPartNo.Visible = False
        
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->PartNumber-LostFocus()")
End Sub


