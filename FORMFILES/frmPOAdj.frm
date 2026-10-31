VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPOAdj 
   Caption         =   "Purchase Order Adjustment"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14655
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   14655
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxAdjPartNo 
      Height          =   2055
      Left            =   14880
      TabIndex        =   35
      Top             =   5040
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   3625
      _Version        =   393216
      Rows            =   0
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      GridColorUnpopulated=   16777215
      GridLines       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxOrdPartNo 
      Height          =   2055
      Left            =   14880
      TabIndex        =   34
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   3625
      _Version        =   393216
      Rows            =   0
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      GridColorUnpopulated=   16777215
      GridLines       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.Timer tmrLoadPO 
      Interval        =   10
      Left            =   10560
      Top             =   120
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   14880
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
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
      Image           =   "frmPOAdj.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   14880
      TabIndex        =   19
      Top             =   480
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
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
      Image           =   "frmPOAdj.frx":059A
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtAdjSO 
      Height          =   285
      Left            =   14880
      TabIndex        =   32
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAdjPartNo 
      Height          =   285
      Left            =   14880
      TabIndex        =   31
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtOrdPartNo 
      Height          =   285
      Left            =   15720
      TabIndex        =   30
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cboADJTYPE 
      Height          =   315
      Left            =   15720
      TabIndex        =   29
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   14880
      TabIndex        =   28
      Text            =   "HoldBackColor"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   14880
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   5415
      Left            =   120
      TabIndex        =   12
      Tag             =   "XYF"
      Top             =   1560
      Width           =   14415
      Begin MSFlexGridLib.MSFlexGrid flxAdjSP 
         Height          =   5055
         Left            =   120
         TabIndex        =   13
         Tag             =   "XYF"
         Top             =   240
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   8916
         _Version        =   393216
         Rows            =   3
         Cols            =   17
         FixedRows       =   2
         BackColorBkg    =   -2147483633
         FormatString    =   "|^ADJ TYPE |<PARTNUMBER |<DESCRIPTION |^ORD QTY |<PARTNUMBER |<DESCRIPTION |^SO |^PRICE"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdAltD 
         Caption         =   "&D"
         Height          =   255
         Left            =   1080
         TabIndex        =   39
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton cmdAltA 
         Caption         =   "&A"
         Height          =   255
         Left            =   720
         TabIndex        =   38
         Top             =   1800
         Width           =   375
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4800
      TabIndex        =   5
      Top             =   0
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   661
      Caption         =   "&Print Preview"
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
      Image           =   "frmPOAdj.frx":0B34
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   120
      TabIndex        =   20
      Tag             =   "XYF"
      Top             =   6960
      Width           =   14415
      Begin VB.TextBox txtAppvdBy 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9600
         TabIndex        =   25
         Tag             =   "XYF"
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox txtPrepBy 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         TabIndex        =   23
         Tag             =   "XYF"
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox txtRemarks 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         TabIndex        =   21
         Tag             =   "XYF"
         Top             =   240
         Width           =   12975
      End
      Begin VB.Label Label7 
         Caption         =   "APPROVED BY:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8280
         TabIndex        =   26
         Tag             =   "XYF"
         Top             =   750
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "PREPARED BY:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Tag             =   "XYF"
         Top             =   750
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "REMARKS:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Tag             =   "XYF"
         Top             =   270
         Width           =   1575
      End
   End
   Begin VB.Timer tmrLoadSP 
      Interval        =   10
      Left            =   12600
      Top             =   0
   End
   Begin MSFlexGridLib.MSFlexGrid flxNewSP 
      Height          =   1215
      Left            =   14880
      TabIndex        =   17
      Tag             =   "xyf"
      Top             =   50000
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   2143
      _Version        =   393216
      Rows            =   0
      Cols            =   11
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.TextBox txtSCODE 
      Height          =   285
      Left            =   14880
      TabIndex        =   16
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   661
      Caption         =   "&Add New"
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
      Image           =   "frmPOAdj.frx":10CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Top             =   0
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   661
      Caption         =   "&Edit"
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
      Image           =   "frmPOAdj.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3240
      TabIndex        =   3
      Top             =   0
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   661
      Caption         =   "&View Summary"
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
      Image           =   "frmPOAdj.frx":1C02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6360
      TabIndex        =   4
      Top             =   0
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   661
      Caption         =   "E&xit"
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
      Image           =   "frmPOAdj.frx":219C
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Tag             =   "XYF"
      Top             =   360
      Width           =   14415
      Begin VB.ComboBox cboPOAdjStatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmPOAdj.frx":2736
         Left            =   5400
         List            =   "frmPOAdj.frx":2740
         TabIndex        =   37
         Tag             =   "XYF"
         Top             =   315
         Width           =   1455
      End
      Begin VB.TextBox txtADJNO 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   12240
         Locked          =   -1  'True
         TabIndex        =   15
         Tag             =   "XYF"
         Text            =   "<AUTO>"
         Top             =   315
         Width           =   2055
      End
      Begin VB.TextBox txtADJDATE 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   12240
         Locked          =   -1  'True
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   645
         Width           =   2055
      End
      Begin VB.TextBox txtPONumber 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   315
         Width           =   2175
      End
      Begin VB.TextBox txtSUPPNAME 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   645
         Width           =   5055
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "ADJ. STATUS:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   36
         Tag             =   "XYF"
         Top             =   345
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "ADJUSTMENT NO:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10320
         TabIndex        =   14
         Tag             =   "XYF"
         Top             =   345
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "ADJUSTMENT DATE:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10320
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "P.O NUMBER:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   350
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "SUPPLIER NAME:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   675
         Width           =   1815
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxDel 
      Height          =   615
      Left            =   14880
      TabIndex        =   33
      Top             =   50000
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   17880
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14895
      m_MDI_Height    =   8955
   End
   Begin VB.Menu mnuStatus 
      Caption         =   "Popupmenu"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuAddRow 
         Caption         =   "ADD ROW"
      End
      Begin VB.Menu mnuDeleteRow 
         Caption         =   "DELETE ROW"
      End
   End
End
Attribute VB_Name = "frmPOAdj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'recordset
Dim rsMain As New ADODB.Recordset
Dim rsLoadPO As New ADODB.Recordset
Dim rsLoadSP As New ADODB.Recordset

'string
Dim strPartial As String, strTotal As String
Dim Tmp() As String, Qry As String, ans As String
Dim formAction As String, MyErr As String
Dim myMsg As String, cmbstr As String

'integer/long
Dim i As Long, j As Long
Dim ActRow As Long, ActCol As Long

'boolean
Dim StatSrch As Boolean, blnFound As Boolean, PrtBln As Boolean, NewSP  As Boolean


Public Sub ShowRecord(MyRef As String)

    Qry = "call sp_mrcwarehouse.sp_POAdj_details('" & MyRef & "')"
    
    With rsMain
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveFirst
            flxAdjSP.Rows = 2
            flxAdjSP.Rows = 3
            flxAdjSP.Rows = .RecordCount + 2

            txtSCODE.Text = IIf(IsNull(!SUPPLIERCODE), "", !SUPPLIERCODE)

            txtPONumber.Text = IIf(IsNull(!PONO), "", !PONO)
            txtSUPPNAME.Text = IIf(IsNull(!SUPPLIERNAME), "", !SUPPLIERNAME)
            txtAdjNo.Text = IIf(IsNull(!POADJNO), "<AUTO>", !POADJNO)
            txtADJDATE.Text = IIf(IsNull(!POADJDATE), "", !POADJDATE)
            txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
            txtPrepBy.Text = IIf(IsNull(!PREPBY), "", !PREPBY)
            txtAppvdBy.Text = IIf(IsNull(!ApprovedBy), "", !ApprovedBy)
            cboPOAdjStatus.Text = IIf(IsNull(!POSTATUS), "", !POSTATUS)
            
            For i = 2 To flxAdjSP.Rows - 1
                flxAdjSP.TextMatrix(i, 1) = IIf(IsNull(!AdjType), "", !AdjType)
                flxAdjSP.TextMatrix(i, 2) = IIf(IsNull(!ORDPARTNO), "", !ORDPARTNO)
                flxAdjSP.TextMatrix(i, 3) = IIf(IsNull(!ORDDESC), "", !ORDDESC)
                flxAdjSP.TextMatrix(i, 4) = IIf(IsNull(!OrderQty), "", !OrderQty)
                flxAdjSP.TextMatrix(i, 6) = IIf(IsNull(!ADJPARTNO), "", !ADJPARTNO)
                flxAdjSP.TextMatrix(i, 7) = IIf(IsNull(!ADJDESC), "", !ADJDESC)
                flxAdjSP.TextMatrix(i, 8) = IIf(IsNull(!ADJSO), "", !ADJSO)
                flxAdjSP.TextMatrix(i, 9) = IIf(IsNull(!ADJPRICE), "", !ADJPRICE)
                flxAdjSP.TextMatrix(i, 10) = IIf(IsNull(Format(!TTLAMNT, "#,###,###.00")), "", Format(!TTLAMNT, "#,###,###.00"))
                flxAdjSP.TextMatrix(i, 11) = IIf(IsNull(!ORDERSPID), "", !ORDERSPID)
                flxAdjSP.TextMatrix(i, 12) = IIf(IsNull(!ADJSPID), "", !ADJSPID)
                .MoveNext
            Next i
        End If
        .Close
    End With
End Sub



Function ValidationPOAdj() As String
On Error GoTo er
   
    MyErr = ""
    If flxAdjSP.TextMatrix(flxAdjSP.Rows - 1, 1) = "" Then
        If flxAdjSP.Rows = 2 Then
            flxAdjSP.Rows = 2
            flxAdjSP.Rows = 3
        Else
            flxAdjSP.RemoveItem flxAdjSP.Rows - 1
        End If
    End If
    
    If flxAdjSP.Rows > 2 Then
        If flxAdjSP.TextMatrix(flxAdjSP.Rows - 1, 1) = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on Last Row" & i
        If flxAdjSP.TextMatrix(flxAdjSP.Rows - 1, 1) = "REPLACE" Then
            If flxAdjSP.TextMatrix(flxAdjSP.Rows - 1, 2) = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on Last Row"
            If flxAdjSP.TextMatrix(flxAdjSP.Rows - 1, 6) = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on Last Row"
            If flxAdjSP.TextMatrix(flxAdjSP.Rows - 1, 8) = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on Last Row"
        Else
            If flxAdjSP.TextMatrix(flxAdjSP.Rows - 1, 6) = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on Last Row"
            If flxAdjSP.TextMatrix(flxAdjSP.Rows - 1, 8) = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on Last Row"
        End If
    End If
    ValidationPOAdj = MyErr


er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->ValidationPartsMatrix()")
End Function

Sub CreateTempTable()
On Error GoTo er
        LYRcon.Execute "Drop Temporary Table IF  Exists sp_mrcwarehouse.TMP_POADJUSTMENT_DTLS;"

         DoEvents
        LYRcon.Execute " Create Temporary Table IF Not Exists sp_mrcwarehouse.TMP_POADJUSTMENT_DTLS( " & _
        "    ADJTYPE            Varchar(50)      " & _
        "   ,ORDERSPID          Int(11)          " & _
        "   ,ORDERQTY           Int(11)          " & _
        "   ,ADJSPID            Int(11)          " & _
        "   ,ADJSO              Int(11)          " & _
        "   ,ADJPRICE           Decimal(10,2)    " & _
        ");"
        
        Qry = "Insert into sp_mrcwarehouse.TMP_POADJUSTMENT_DTLS(ADJTYPE,ORDERSPID,ORDERQTY,ADJSPID,ADJSO,ADJPRICE) VALUES "
'********************** CODE UPDATED: 06/16/17 **************************
        For i = 2 To flxAdjSP.Rows - 1
            If i > 2 Then Qry = Qry & ","
            Qry = Qry & "('" & flxAdjSP.TextMatrix(i, 1) & "'," & Val(flxAdjSP.TextMatrix(i, 11)) & "," & Val(flxAdjSP.TextMatrix(i, 4)) & "," & Val(flxAdjSP.TextMatrix(i, 12)) & "," & Val(flxAdjSP.TextMatrix(i, 8)) & ", '" & Replace(flxAdjSP.TextMatrix(i, 9), ",", "") & "') "
        Next i
'****************** END CODE UPDATED: 06/16/17 **************************
       
       Call LYRcon.Execute(Qry)
       
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CreateTempTable()")
End Sub

Sub Clearform()
    flxAdjSP.Rows = 2
    flxAdjSP.Rows = 3
    txtPONumber.Text = ""
    txtSUPPNAME.Text = ""
    txtAdjNo.Text = "<AUTO>"
    txtADJDATE.Text = Format(ServerTime, "mm/dd/yyyy")
    txtRemarks.Text = ""
    txtPrepBy.Text = ""
    txtAppvdBy.Text = ""
    cboPOAdjStatus.Text = ""
    txtSCODE.Text = ""
End Sub

Sub LockData(DataEntry As Boolean)
    DataEntry = Not DataEntry

    txtPONumber.Locked = DataEntry
    txtSUPPNAME.Locked = DataEntry
    txtRemarks.Locked = DataEntry
    txtPrepBy.Locked = DataEntry
    txtAppvdBy.Locked = DataEntry
    cboPOAdjStatus.Locked = DataEntry
    
    If txtPONumber.Locked Then
        txtPONumber.BackColor = txtHoldGrayBackColor.BackColor
        txtSUPPNAME.BackColor = txtHoldGrayBackColor.BackColor
        txtRemarks.BackColor = txtHoldGrayBackColor.BackColor
        txtPrepBy.BackColor = txtHoldGrayBackColor.BackColor
        txtAppvdBy.BackColor = txtHoldGrayBackColor.BackColor
        cboPOAdjStatus.BackColor = txtHoldGrayBackColor.BackColor
    Else
        txtPONumber.BackColor = txtHoldWhiteBackColor.BackColor
        txtRemarks.BackColor = txtHoldWhiteBackColor.BackColor
        txtAppvdBy.BackColor = txtHoldWhiteBackColor.BackColor
        cboPOAdjStatus.BackColor = txtHoldWhiteBackColor.BackColor
        txtPrepBy.Locked = True
        txtSUPPNAME.Locked = True
    End If
End Sub


Private Sub SetButtonVisibility(DataEntry As Boolean)
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->VIEW SUMM
'3 ->PRINT
'4 ->EXIT
'5 ->SAVE
'6 ->CANCEL


    'hide all button
    For i = 0 To 4
        cmdTrans(i).Visible = False
    Next i

    If DataEntry Then
        cmdTrans(5).Visible = True
        cmdTrans(6).Visible = True
    Else
        For i = 0 To 4
            cmdTrans(i).Visible = True
            cmdTrans(5).Visible = False
            cmdTrans(6).Visible = False
        Next i
    End If
End Sub

Public Sub flxColChanger()

    With flxAdjSP
    
        .ColAlignment(5) = 4
        'MERGE COL
        .TextMatrix(0, 0) = " "
        .TextMatrix(1, 0) = " "
        'MERGE ROW
        .TextMatrix(0, 1) = "ORDERED SPAREPARTS DETAILS"
        .TextMatrix(0, 2) = "ORDERED SPAREPARTS DETAILS"
        .TextMatrix(0, 3) = "ORDERED SPAREPARTS DETAILS"
        .TextMatrix(0, 4) = "ORDERED SPAREPARTS DETAILS"
        .TextMatrix(0, 5) = "ADJUSTMENT SPAREPARTS DETAILS"
        .TextMatrix(0, 6) = "ADJUSTMENT SPAREPARTS DETAILS"
        .TextMatrix(0, 7) = "ADJUSTMENT SPAREPARTS DETAILS"
        .TextMatrix(0, 8) = "ADJUSTMENT SPAREPARTS DETAILS"
        .TextMatrix(0, 9) = "ADJUSTMENT SPAREPARTS DETAILS"
        .TextMatrix(0, 10) = "ADJUSTMENT SPAREPARTS DETAILS"
        
        'MERGE COL =FALSE
        .TextMatrix(1, 1) = "ADJ TYPE"
        .TextMatrix(1, 2) = "PARTNUMBER"
        .TextMatrix(1, 3) = "DESCRIPTION"
        .TextMatrix(1, 4) = "OQ"
        .TextMatrix(1, 5) = ""
        .TextMatrix(1, 6) = "PARTNUMBER"
        .TextMatrix(1, 7) = "DESCRIPTION"
        .TextMatrix(1, 8) = "SO"
        .TextMatrix(1, 9) = "PRICE"
        .TextMatrix(1, 10) = "TTL PRICE"
        
        
        'SET MERGING
        .MergeCells = 1
        .MergeRow(0) = True
        .MergeCol(0) = True
        .MergeCol(1) = True
        For i = 1 To 3
            .MergeCol(i) = False
        Next i
        
        'SET THE COLUMN COLOR GREEN
        .row = 0
        .Col = 0
        .Col = 1
    
        '.Clear
        .Cols = 17
        
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.07        'ADJ TYPE
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.15        'PARTNUMBER
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.19        'DESCRIPTION
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.05        'ORDER QTY
        .ColWidth(5) = 0                                     'Para sa Alignment sa header
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.15        'PARTNUMBER
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.18        'DESCRIPTION
        .ColWidth(8) = (.Width - .ColWidth(0)) * 0.05        'SO
        .ColWidth(9) = (.Width - .ColWidth(0)) * 0.07        'PRICE
        .ColWidth(10) = (.Width - .ColWidth(0)) * 0.07       'Total Price
        .ColWidth(11) = 0 '(.Width - .ColWidth(0)) * 0.06    'ORDER SPID
        .ColWidth(12) = 0 '(.Width - .ColWidth(0)) * 0.06    'NEW SPID
        .ColWidth(13) = 0 '(.Width - .ColWidth(0)) * 0.06    'STOREPRICE or RETAIL PRICE
        .ColWidth(14) = 0 '(.Width - .ColWidth(0)) * 0.06    'APPLIC MODEL
        .ColWidth(15) = 0 '(.Width - .ColWidth(0)) * 0.06    'REMARKS - if new or dili
        .ColWidth(16) = 0 '(.Width - .ColWidth(0)) * 0.06    'MAX DISCOUNT
        
        flxOrdPartNo.ColWidth(4) = .ColWidth(2)
        flxAdjPartNo.ColWidth(1) = .ColWidth(6)
       
    End With

    With flxOrdPartNo
        .Top = txtOrdPartNo.Height + txtOrdPartNo.Top
        .Left = txtOrdPartNo.Left
        .Width = txtOrdPartNo.Width
    End With
    
    With flxAdjPartNo
        .Top = txtAdjPartNo.Height + txtAdjPartNo.Top
        .Left = txtAdjPartNo.Left
        .Width = txtAdjPartNo.Width
    End With
    
End Sub


Private Sub cboADJTYPE_Change()
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If
    With cboADJTYPE
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
End Sub

Private Sub cboADJTYPE_Click()
    cmbstr = cboADJTYPE.Text
    For i = 0 To cboADJTYPE.ListCount - 1
        If UCase$(cboADJTYPE.List(i)) = UCase$(cmbstr) Then
            cboADJTYPE.ListIndex = i
            blnFound = True
            Exit For
        End If
    Next i
    
    If Not blnFound Then
        MsgBox "Item was not found in the list.", _
        vbExclamation, _
        "Warning"
        cboADJTYPE.Visible = False
        Exit Sub
    End If
    blnFound = False
    
    flxAdjSP.TextMatrix(ActRow, 1) = cboADJTYPE.Text
    
    DoEvents
    If flxAdjSP.TextMatrix(ActRow, 1) = "REPLACE" Or cboADJTYPE.Text = "REPLACE" Then
        flxAdjSP.SetFocus
        flxAdjSP.Col = 2
        flxAdjSP.row = ActRow
    ElseIf flxAdjSP.TextMatrix(ActRow, 1) = "NEW" Or cboADJTYPE.Text = "NEW" Then
        flxAdjSP.TextMatrix(ActRow, 2) = ""
        flxAdjSP.TextMatrix(ActRow, 3) = ""
        flxAdjSP.TextMatrix(ActRow, 4) = ""
        flxAdjSP.TextMatrix(ActRow, 11) = ""
        flxAdjSP.SetFocus
        flxAdjSP.Col = 6
        flxAdjSP.row = ActRow
    End If
    cboADJTYPE.Visible = False
End Sub

Private Sub cboADJTYPE_KeyDown(KeyCode As Integer, Shift As Integer)
 
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        StatSrch = True
    End If
End Sub

Private Sub cboADJTYPE_KeyPress(KeyAscii As Integer)
On Error GoTo er

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboADJTYPE, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Call cboADJTYPE_Click
        'cboADJTYPE.Visible = False
    End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboADJTYPE-KeyPress()")

End Sub

Private Sub cmdAltA_Click()
Call mnuAddRow_Click
End Sub

Private Sub cmdAltD_Click()
Call mnuDeleteRow_Click
End Sub


Private Sub cmdTrans_Click(Index As Integer)
On Error GoTo er

    If Index = 0 Then 'add
        Call LockData(True)
        Call SetButtonVisibility(True)
        Call Clearform
        Qry = "SELECT concat(employees.FIRSTNAME,' ',employees.LASTNAME) as PrepBy From employees Where employees.EmpId = " & UserID & ";"
        txtPrepBy.Text = GetFieldValue(Qry)
        formAction = "add"
    ElseIf Index = 1 Then 'edit
        Call LockData(True)
        Call SetButtonVisibility(True)
        formAction = "edit"
    ElseIf Index = 2 Then 'view
        FrmSummary.mySrc = "PO ADJUSTMENT"
        FrmSummary.Show
        Unload Me
    ElseIf Index = 3 Then 'print
        If txtAdjNo.Text <> "<AUTO>" Then
            Qry = "call sp_mrcwarehouse.sp_POAdj_details('" & txtAdjNo.Text & "')"
            frmPrn.QryPrn = Qry
        
            frmPrn.ParamValue1 = SysBranch
            frmPrn.ParamValue2 = Format(SysAddress, "")
            
            frmPrn.MySource = "PO ADJUSTMENT"
            frmPrn.Show vbModal
        Else
            MsgBox "No Record to Print", vbInformation + vbOKOnly, "Validation"
            Exit Sub
        End If
    ElseIf Index = 4 Then 'exit
        Call Clearform
        formAction = ""
        Unload Me
    ElseIf Index = 5 Then 'save

        myMsg = ValidationPOAdj

        If myMsg <> "" Then
            MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
            Exit Sub
        End If

        ans = MsgBox("Do you want to save?", vbQuestion + vbYesNo, "Save")
        If ans = vbNo Then Exit Sub

        If formAction = "add" Then
            Call CreateTempTable
            Qry = "call sp_mrcwarehouse.sp_POAdj_create('" & txtPONumber.Text & "','" & cboPOAdjStatus.Text & "','" & txtSUPPNAME.Text & "','" & txtRemarks.Text & "','" & txtAppvdBy.Text & "'," & Val(UserID) & ")"
            
            With rsMain
                If .State = adStateOpen Then .Close
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                
                If .RecordCount > 0 Then
                    .MoveFirst
                    
                    If !level = "Granted" Then
                        MsgBox "" & !message & "", vbInformation + vbOKOnly, "" & !level & ""
                        txtAdjNo.Text = !POADJNO
                    Else
                        MsgBox "" & !message & "", vbInformation + vbOKOnly, "" & !level & ""
                        .Close
                        Exit Sub
                    End If
                End If
                .Close
            End With
        Else
            Call CreateTempTable
            Qry = "call sp_mrcwarehouse.sp_POAdj_update('" & txtAdjNo.Text & "','" & txtPONumber.Text & "','" & cboPOAdjStatus.Text & "','" & txtSUPPNAME.Text & "','" & txtRemarks.Text & "','" & txtAppvdBy.Text & "'," & Val(UserID) & ")"
            
            With rsMain
                If .State = adStateOpen Then .Close
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                
                If .RecordCount > 0 Then
                    .MoveFirst
                    
                    If !level = "Granted" Then
                        MsgBox "" & !message & "", vbInformation + vbOKOnly, "" & !level & ""
                    Else
                        MsgBox "" & !message & "", vbInformation + vbOKOnly, "" & !level & ""
                        .Close
                        Exit Sub
                    End If
                End If
                .Close
            End With
        End If

        Call SetButtonVisibility(False)
        Call LockData(False)
        formAction = ""
    ElseIf Index = 6 Then 'cancel
        
        ans = MsgBox("Do you want to cancel?", vbQuestion + vbYesNo)
        
        If ans = vbNo Then Exit Sub
        
        Call SetButtonVisibility(False)
        Call LockData(False)
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdTrans-Click()")
End Sub



Private Sub flxAdjPartNo_Click()
    With flxAdjPartNo
        txtAdjPartNo.Visible = True
        txtAdjPartNo.ZOrder
        StatSrch = True
        
        txtAdjPartNo.Text = flxAdjPartNo.TextMatrix(flxAdjPartNo.row, 1)
        flxAdjSP.TextMatrix(flxAdjSP.row, 7) = flxAdjPartNo.TextMatrix(flxAdjPartNo.row, 2) 'rsLoadSP!Description
        flxAdjSP.TextMatrix(flxAdjSP.row, 9) = flxAdjPartNo.TextMatrix(flxAdjPartNo.row, 3) 'rsLoadSP!Cost
        flxAdjSP.TextMatrix(flxAdjSP.row, 12) = flxAdjPartNo.TextMatrix(flxAdjPartNo.row, 0) 'rsLoadSP!SPAREPARTID
        
        txtAdjPartNo.SelStart = Val(txtAdjPartNo.Tag)
        txtAdjPartNo.SelLength = Len(txtAdjPartNo.Text)
    End With

    If flxAdjSP.TextMatrix(ActRow, 8) <> "" And flxAdjSP.TextMatrix(ActRow, 2) <> "" And flxAdjSP.TextMatrix(ActRow, 1) = "REPLACE" Then
        txtOrdPartNo.Tag = "NxtRow"
        If ActRow = flxAdjSP.Rows - 1 Then
            flxAdjSP.AddItem ""
            flxAdjSP.SetFocus
            flxAdjSP.row = flxAdjSP.Rows - 1
            flxAdjSP.Col = 1
        End If
    ElseIf flxAdjSP.TextMatrix(ActRow, 3) = "" And flxAdjSP.TextMatrix(ActRow, 1) = "REPLACE" Then
        flxAdjSP.SetFocus
        flxAdjSP.Col = 2
    ElseIf flxAdjSP.TextMatrix(ActRow, 6) = "" Then
        flxAdjSP.SetFocus
        flxAdjSP.Col = 6
    ElseIf flxAdjSP.TextMatrix(ActRow, 8) = "" Then
        flxAdjSP.SetFocus
        flxAdjSP.Col = 8
    End If
    
    txtAdjPartNo.Visible = False
    flxAdjPartNo.Visible = False
    flxAdjSP.SetFocus

End Sub

Private Sub flxAdjPartNo_EnterCell()
On Error GoTo er
    With flxAdjPartNo
        txtAdjPartNo.Visible = True
        txtAdjPartNo.ZOrder
        StatSrch = True
        
        txtAdjPartNo.Text = flxAdjPartNo.TextMatrix(flxAdjPartNo.row, 1)
        flxAdjSP.TextMatrix(flxAdjSP.row, 7) = flxAdjPartNo.TextMatrix(flxAdjPartNo.row, 2) 'rsLoadSP!Description
        flxAdjSP.TextMatrix(flxAdjSP.row, 9) = flxAdjPartNo.TextMatrix(flxAdjPartNo.row, 3) 'rsLoadSP!Cost
        flxAdjSP.TextMatrix(flxAdjSP.row, 12) = flxAdjPartNo.TextMatrix(flxAdjPartNo.row, 0) 'rsLoadSP!SPAREPARTID
        
        txtAdjPartNo.SelStart = Val(txtAdjPartNo.Tag)
        txtAdjPartNo.SelLength = Len(txtAdjPartNo.Text)
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->FlxPartNumber-EnterCell()")
End Sub

Private Sub flxAdjPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
       Tmp = Split(txtAdjPartNo.Tag, "|")
       For i = LBound(Tmp) To UBound(Tmp)
           flxAdjPartNo.TextMatrix(ActRow, i + 1) = Tmp(i)
       Next i
       txtAdjPartNo.Visible = False
       flxAdjPartNo.Visible = False
       flxAdjSP.SetFocus
       flxAdjSP.Col = 1
    End If
    
End Sub

Private Sub flxAdjPartNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    
    If KeyAscii = 13 Then
        Call flxAdjPartNo_Click
    End If
    
End Sub

Private Sub flxAdjPartNo_LostFocus()
Call txtAdjPartNo_LostFocus
End Sub

Private Sub flxAdjSP_Click()
On Error GoTo er

    ActRow = flxAdjSP.row
    ActCol = flxAdjSP.Col
    If formAction = "add" Or formAction = "edit" Then
        If ActCol = 0 Or ActRow = 0 Then Exit Sub
            Select Case ActCol
                Case 1
                    With cboADJTYPE
                        If flxAdjSP.TextMatrix(ActRow, 1) <> "" Then .Tag = flxAdjSP.TextMatrix(ActRow, 1)
                        .Top = Frame2.Top + flxAdjSP.CellTop + flxAdjSP.Top
                        .Left = Frame2.Left + flxAdjSP.CellLeft + flxAdjSP.Left
                        .FontSize = 7
                        '.Height = flxAdjSP.CellHeight
                        .Width = flxAdjSP.CellWidth
                        .Text = flxAdjSP.Text
                        .Visible = True
                        .SelStart = 0
                        .SelLength = Len(.Text)
                        .ZOrder
                        .SetFocus
                    End With
            End Select
            If flxAdjSP.TextMatrix(ActRow, 1) = "REPLACE" And txtPONumber.Text <> "" Then
                Select Case ActCol
                    Case 2
                        With txtOrdPartNo
                            If flxAdjSP.TextMatrix(ActRow, 2) <> "" Then .Tag = flxAdjSP.TextMatrix(ActRow, 2) & "|" & flxAdjSP.TextMatrix(ActRow, 3) & "|" & flxAdjSP.TextMatrix(ActRow, 4) & "|" & flxAdjSP.TextMatrix(ActRow, 11)
                            .Top = Frame2.Top + flxAdjSP.CellTop + flxAdjSP.Top
                            .Left = Frame2.Left + flxAdjSP.CellLeft + flxAdjSP.Left
                            '.Height = flxAdjSP.CellHeight - 50
                            .Width = flxAdjSP.CellWidth
                            .Text = flxAdjSP.Text
                            .Visible = True
                            .SelStart = 0
                            .SelLength = Len(.Text)
                            .ZOrder
                            .SetFocus
                        End With
                    Case 6
                        With txtAdjPartNo
                            If flxAdjSP.TextMatrix(ActRow, 6) <> "" Then .Tag = flxAdjSP.TextMatrix(ActRow, 6) & "|" & flxAdjSP.TextMatrix(ActRow, 7) & "|" & flxAdjSP.TextMatrix(ActRow, 8) & "|" & flxAdjSP.TextMatrix(ActRow, 9) & "|" & flxAdjSP.TextMatrix(ActRow, 12)
                            .Top = Frame2.Top + flxAdjSP.CellTop + flxAdjSP.Top
                            .Left = Frame2.Left + flxAdjSP.CellLeft + flxAdjSP.Left
                            '.Height = flxAdjSP.CellHeight - 50
                            .Width = flxAdjSP.CellWidth
                            .Text = flxAdjSP.Text
                            .Visible = True
                            .SelStart = 0
                            .SelLength = Len(.Text)
                            .ZOrder
                            .SetFocus
                        End With
                    Case 8
                        With txtAdjSO
                            If flxAdjSP.TextMatrix(ActRow, 8) <> "" Then .Tag = flxAdjSP.TextMatrix(ActRow, 8)
                            .Top = Frame2.Top + flxAdjSP.CellTop + flxAdjSP.Top
                            .Left = Frame2.Left + flxAdjSP.CellLeft + flxAdjSP.Left
                            '.Height = flxAdjSP.CellHeight - 30
                            .Width = flxAdjSP.CellWidth
                            .Text = flxAdjSP.Text
                            .Visible = True
                            .SelStart = 0
                            .SelLength = Len(.Text)
                            .ZOrder
                            .SetFocus
                        End With
                End Select
            ElseIf flxAdjSP.TextMatrix(ActRow, 1) = "NEW" Then
                Select Case ActCol
                    Case 6
                        With txtAdjPartNo
                            If flxAdjSP.TextMatrix(ActRow, 6) <> "" Then .Tag = flxAdjSP.TextMatrix(ActRow, 6) & "|" & flxAdjSP.TextMatrix(ActRow, 7) & "|" & flxAdjSP.TextMatrix(ActRow, 8) & "|" & flxAdjSP.TextMatrix(ActRow, 9) & "|" & flxAdjSP.TextMatrix(ActRow, 12)
                            .Top = Frame2.Top + flxAdjSP.CellTop + flxAdjSP.Top
                            .Left = Frame2.Left + flxAdjSP.CellLeft + flxAdjSP.Left
                            '.Height = flxAdjSP.CellHeight - 50
                            .Width = flxAdjSP.CellWidth
                            .Text = flxAdjSP.Text
                            .Visible = True
                            .SelStart = 0
                            .SelLength = Len(.Text)
                            .ZOrder
                            .SetFocus
                        End With
                    Case 8
                        With txtAdjSO
                            If flxAdjSP.TextMatrix(ActRow, 8) <> "" Then .Tag = flxAdjSP.TextMatrix(ActRow, 8)
                            .Top = Frame2.Top + flxAdjSP.CellTop + flxAdjSP.Top
                            .Left = Frame2.Left + flxAdjSP.CellLeft + flxAdjSP.Left
                            '.Height = flxAdjSP.CellHeight - 30
                            .Width = flxAdjSP.CellWidth
                            .Text = flxAdjSP.Text
                            .Visible = True
                            .SelStart = 0
                            .SelLength = Len(.Text)
                            .ZOrder
                            .SetFocus
                        End With
                End Select
            End If
    End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->flxSPARTS-click()")

End Sub

Private Sub flxAdjSP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call flxAdjSP_Click
    End If
End Sub

Private Sub flxAdjSP_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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

Private Sub flxAdjSP_Scroll()
cboADJTYPE.Visible = False
End Sub

Private Sub flxOrdPartNo_Click()
    With flxOrdPartNo
        txtOrdPartNo.Visible = True
        txtOrdPartNo.ZOrder
        StatSrch = True

        txtOrdPartNo.Text = flxOrdPartNo.TextMatrix(flxOrdPartNo.row, 4)
        flxAdjSP.TextMatrix(flxAdjSP.row, 3) = flxOrdPartNo.TextMatrix(flxOrdPartNo.row, 5) 'rsLoadPO!Description
        flxAdjSP.TextMatrix(flxAdjSP.row, 4) = flxOrdPartNo.TextMatrix(flxOrdPartNo.row, 9) 'rsLoadPO!Order
        flxAdjSP.TextMatrix(flxAdjSP.row, 11) = flxOrdPartNo.TextMatrix(flxOrdPartNo.row, 3) 'rsLoadPO!SPAREPARTID
        
        txtOrdPartNo.SelStart = Val(txtOrdPartNo.Tag)
        txtOrdPartNo.SelLength = Len(txtOrdPartNo.Text)
    End With
    
    If flxAdjSP.TextMatrix(ActRow, 3) <> "" And flxAdjSP.TextMatrix(ActRow, 6) <> "" And flxAdjSP.TextMatrix(ActRow, 8) <> "" And flxAdjSP.TextMatrix(ActRow, 1) = "REPLACE" Then
        txtOrdPartNo.Tag = "NxtRow"
        If ActRow = flxAdjSP.Rows - 1 Then
            flxAdjSP.AddItem ""
            flxAdjSP.SetFocus
            flxAdjSP.row = flxAdjSP.Rows - 1
            flxAdjSP.Col = 1
        End If
    ElseIf flxAdjSP.TextMatrix(ActRow, 3) = "" Then
        flxAdjSP.SetFocus
        flxAdjSP.Col = 2
    ElseIf flxAdjSP.TextMatrix(ActRow, 6) = "" Then
        flxAdjSP.SetFocus
        flxAdjSP.Col = 6
    ElseIf flxAdjSP.TextMatrix(ActRow, 8) = "" Then
        flxAdjSP.SetFocus
        flxAdjSP.Col = 8
    End If
        
    txtOrdPartNo.Visible = False
    flxOrdPartNo.Visible = False
End Sub

Private Sub flxOrdPartNo_EnterCell()
On Error GoTo er
    With flxOrdPartNo
        txtOrdPartNo.Visible = True
        txtOrdPartNo.ZOrder
        StatSrch = True
        
        txtOrdPartNo.Text = flxOrdPartNo.TextMatrix(flxOrdPartNo.row, 4)
        flxAdjSP.TextMatrix(flxAdjSP.row, 3) = flxOrdPartNo.TextMatrix(flxOrdPartNo.row, 5) 'rsLoadPO!Description
        flxAdjSP.TextMatrix(flxAdjSP.row, 4) = flxOrdPartNo.TextMatrix(flxOrdPartNo.row, 9) 'rsLoadPO!Order
        flxAdjSP.TextMatrix(flxAdjSP.row, 11) = flxOrdPartNo.TextMatrix(flxOrdPartNo.row, 3) 'rsLoadPO!SPAREPARTID
        
        txtOrdPartNo.SelStart = Val(txtOrdPartNo.Tag)
        txtOrdPartNo.SelLength = Len(txtOrdPartNo.Text)
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->FlxPartNumber-EnterCell()")
End Sub

Private Sub flxOrdPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
       Tmp = Split(txtOrdPartNo.Tag, "|")
       For i = LBound(Tmp) To UBound(Tmp)
           flxOrdPartNo.TextMatrix(ActRow, i + 1) = Tmp(i)
       Next i
       txtOrdPartNo.Visible = False
       flxOrdPartNo.Visible = False
       flxAdjSP.SetFocus
       flxAdjSP.Col = 1
    End If
End Sub

Private Sub flxOrdPartNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    
    If KeyAscii = 13 Then
        Call flxOrdPartNo_Click
    End If
End Sub

Private Sub flxOrdPartNo_LostFocus()
Call txtOrdPartNo_LostFocus
End Sub

Private Sub Form_Load()
    
    Call LockData(False)
    
    ActRow = 0
    ActCol = 0
    txtAdjNo.Text = "<AUTO>"
    txtADJDATE.Text = Format(ServerTime, "mm/dd/yyyy")
    cboADJTYPE.Text = ""
    cboADJTYPE.AddItem "REPLACE"
    cboADJTYPE.AddItem "NEW"
    
    '===================BUTTON POSITION===================
    If cmdTrans(0).Visible = False Then
            'Align Save Button
            cmdTrans(5).Left = cmdTrans(0).Left
            cmdTrans(5).Top = cmdTrans(0).Top
            'Align Cancel Button
            cmdTrans(6).Left = cmdTrans(1).Left
            cmdTrans(6).Top = cmdTrans(1).Top
    End If
    tmrLoadSP.Enabled = True
    Warehouse.sBar.Panels(6).Text = Me.Caption
End Sub

Private Sub Form_Resize()
Call RXResizer1_ResizeComplete
End Sub

Private Sub mnuAddRow_Click()
If formAction = "add" Or formAction = "edit" Then
    If flxAdjSP.TextMatrix(flxAdjSP.Rows - 1, 1) = "" Or flxAdjSP.TextMatrix(flxAdjSP.Rows - 1, 2) = "" Or flxAdjSP.TextMatrix(flxAdjSP.Rows - 1, 6) = "" Or flxAdjSP.TextMatrix(flxAdjSP.Rows - 1, 8) = "" Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        Exit Sub
    Else
        flxAdjSP.AddItem ""
        flxAdjSP.SetFocus
        flxAdjSP.row = flxAdjSP.Rows - 1
        flxAdjSP.Col = 1
    End If
End If
End Sub

Private Sub mnuDeleteRow_Click()
On Error GoTo er
If formAction = "add" Or formAction = "edit" Then
    ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If ans = vbYes Then
        If flxAdjSP.Rows = 3 Then
            flxAdjSP.Rows = 2
            flxAdjSP.Rows = 3
        Else
            flxAdjSP.RemoveItem ActRow
        End If
    End If
End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->mnuDelete()")
End Sub


Private Sub RXResizer1_ResizeComplete()
Call flxColChanger
End Sub

Private Sub tmrLoadPO_Timer()
On Error GoTo er

    If txtPONumber.Text <> "" Then
        DoEvents
        Qry = "call sp_mrcwarehouse.sp_POAdj_getPODetails('" & txtPONumber.Text & "')"
        If rsLoadPO.State = adStateOpen Then rsLoadPO.Close
        rsLoadPO.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    
        Set flxOrdPartNo.DataSource = rsLoadPO
        flxOrdPartNo.Refresh
        With flxOrdPartNo
            .Width = flxAdjSP.ColWidth(2)
            .ColWidth(0) = 0
            .ColWidth(1) = 0
            .ColWidth(2) = 0
            .ColWidth(3) = 0
            .ColWidth(4) = flxAdjSP.ColWidth(2)
            .ColWidth(5) = 0
            .ColWidth(6) = 0
            .ColWidth(7) = 0
            .ColWidth(8) = 0
            .ColWidth(9) = 0
        End With
    End If
    tmrLoadPO.Enabled = False
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->LoadOrderedPartNumber()")
End Sub

Private Sub tmrLoadSP_Timer()
On Error GoTo er
    DoEvents
    Qry = "Call sp_mrcwarehouse.sp_POAdj_getSpareparts()"
    If rsLoadSP.State = adStateOpen Then rsLoadSP.Close
    rsLoadSP.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic

    Set flxAdjPartNo.DataSource = rsLoadSP
    flxAdjPartNo.Refresh
    With flxAdjPartNo
        .Width = flxAdjSP.ColWidth(6)
        .ColWidth(0) = 0
        .ColWidth(1) = flxAdjSP.ColWidth(6)
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 0
    End With
    
    tmrLoadSP.Enabled = False
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->LoadPartNumber()")
End Sub

Private Sub txtAdjPartNo_Change()
On Error GoTo LocalErr
    With flxAdjPartNo
        If StatSrch Then
            StatSrch = False
            Exit Sub
        End If
        .Top = txtAdjPartNo.Height + txtAdjPartNo.Top
        .Left = txtAdjPartNo.Left
         .Width = txtAdjPartNo.Width
        .Visible = True
        .ZOrder
            
        If txtAdjPartNo.Text = "" Then .Visible = False: Exit Sub
        rsLoadSP.Find "[PARTNUMBER] LIKE '" + txtAdjPartNo.Text + "%'", , adSearchForward, 1
        If Not rsLoadSP.EOF Then
            If Not PrtBln Then
                txtAdjPartNo.Tag = Len(txtAdjPartNo.Text)
                PrtBln = True
            End If
            
            .TopRow = rsLoadSP.AbsolutePosition - 1
            .row = rsLoadSP.AbsolutePosition - 1
            .Col = 0: .ColSel = 1
            txtAdjPartNo.Text = rsLoadSP!PartNumber
            flxAdjSP.TextMatrix(flxAdjSP.row, 7) = rsLoadSP!Description
            flxAdjSP.TextMatrix(flxAdjSP.row, 9) = rsLoadSP!Cost
            flxAdjSP.TextMatrix(flxAdjSP.row, 12) = rsLoadSP!SPAREPARTID
            
            txtAdjPartNo.SelStart = Val(txtAdjPartNo.Tag)
            txtAdjPartNo.SelLength = Len(txtAdjPartNo.Text)
            PrtBln = False
        Else
            flxAdjSP.TextMatrix(ActRow, 6) = ""
            flxAdjSP.TextMatrix(ActRow, 7) = ""
            flxAdjSP.TextMatrix(ActRow, 8) = ""
            flxAdjSP.TextMatrix(ActRow, 9) = ""
            flxAdjSP.TextMatrix(ActRow, 10) = ""
            .Visible = False
        End If
   End With
   NewSP = False
   
   
   Exit Sub

LocalErr:
    If Err.Number = 30009 Then
        flxAdjPartNo.Visible = False
        txtAdjPartNo.Visible = False
        If rsLoadSP.State = adStateOpen Then rsLoadSP.Close
        tmrLoadSP.Enabled = True
        MsgBox "Program is reloading Part Number" & vbCrLf & "Please wait for a few seconds...", vbInformation + vbOKOnly, "Reloading"
    Else
        flxAdjPartNo.Visible = False
        txtAdjPartNo.Visible = False
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
    End If
End Sub

Private Sub txtAdjPartNo_GotFocus()
    If rsLoadSP.State = adStateClosed Then tmrLoadSP.Enabled = True
    With flxAdjPartNo
        .Top = txtAdjPartNo.Height + txtAdjPartNo.Top
        .Left = txtAdjPartNo.Left
        .Width = txtAdjPartNo.Width
    End With
End Sub

Private Sub txtAdjPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim Tmp() As String
    
    Select Case KeyCode
        Case vbKeyDown
            txtAdjPartNo.Tag = "skiplostfocus"
            flxAdjPartNo.Visible = True
            flxAdjPartNo.SetFocus
        Case vbKeyEscape
            txtAdjPartNo.Text = Empty
            Tmp = Split(flxAdjPartNo.Tag, "|")
            For i = LBound(Tmp) To UBound(Tmp)
                flxAdjPartNo.TextMatrix(ActRow, i + 1) = Tmp(i)
            Next i
            flxAdjPartNo.Tag = ""
            txtAdjPartNo.Tag = "skiplostfocus"
            flxAdjSP.SetFocus
            flxAdjSP.Col = 1
            txtAdjPartNo.Visible = False
            flxAdjPartNo.Visible = False
        Case vbKeyDelete
            StatSrch = True
        Case vbKeyBack
            StatSrch = True
    End Select
End Sub

Private Sub txtAdjPartNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        Call txtAdjPartNo_LostFocus
        
'        If txtAdjPartNo.Tag = "PN exist" Then Exit Sub
'        If txtAdjPartNo.Tag = "super" Then
'            txtAdjPartNo.SetFocus
'            txtAdjPartNo.Tag = ""
'            Exit Sub
'        End If
'        If txtAdjPartNo.Tag = "NxtRow" Then
'            txtAdjPartNo.Tag = ""
'            Exit Sub
'        End If
        DoEvents
        If flxAdjSP.TextMatrix(ActRow, 8) <> "" Then
            txtAdjPartNo.Tag = "NxtRow"
            If ActRow = flxAdjSP.Rows - 1 Then
                flxAdjSP.AddItem ""
                flxAdjSP.SetFocus
                flxAdjSP.Col = 1
                flxAdjSP.row = flxAdjSP.Rows - 1
            End If
        ElseIf flxAdjSP.TextMatrix(ActRow, 6) = "" Then
            flxAdjSP.SetFocus
            flxAdjSP.Col = 6
            flxAdjSP.row = ActRow
        ElseIf flxAdjSP.TextMatrix(ActRow, 8) = "" Then
            flxAdjSP.SetFocus
            flxAdjSP.Col = 8
            flxAdjSP.row = ActRow
        End If
    End If
End Sub

Private Sub txtAdjPartNo_LostFocus()
    Dim ans As Integer, i As Integer, xRow As Integer, j As Integer
    Dim Tmp() As String
    
    If txtAdjPartNo.Tag = "skiplostfocus" Then
        txtAdjPartNo.Tag = ""
        Exit Sub
    End If
    'Validate existing PartNo
    For i = 2 To flxAdjSP.Rows - 1
        If i <> ActRow Then
            If flxAdjSP.TextMatrix(i, 6) = txtAdjPartNo.Text Then
                flxAdjSP.TextMatrix(ActRow, 6) = txtAdjPartNo.Text
                txtAdjPartNo.Visible = False
                flxAdjPartNo.Visible = False
                ans = MsgBox("Part Number " & txtAdjPartNo.Text & " already exist in row " & i & vbCrLf & _
                       "Do you want to locate existing part number?", vbQuestion + vbYesNo, "Duplicate Part Number")
                If ans = vbYes Then
                    If flxAdjPartNo.Tag <> "" Then
                        Tmp = Split(flxAdjPartNo.Tag, "|")
                        For j = LBound(Tmp) To UBound(Tmp)
                            flxAdjSP.TextMatrix(ActRow, j + 1) = Tmp(j)
                        Next j
                    Else
                        For j = 1 To flxAdjSP.Cols - 1
                            flxAdjSP.TextMatrix(ActRow, j) = ""
                        Next j
                    End If
                    flxAdjSP.row = i
                    flxAdjSP.Col = flxAdjSP.Cols - 1
                    flxAdjSP.RowSel = i
                    flxAdjSP.ColSel = 0
                    flxAdjSP.TopRow = i
                    txtAdjPartNo.Tag = "PN exist"
                    Exit Sub
                End If
            End If
        End If
    Next i
    
'    'If New SP load AddNewSP form
'    If flxAdjSP.TextMatrix(ActRow, 7) = "" And txtAdjPartNo.Text <> "" Then
'        ans = MsgBox("Unrecognized Part Number!" & vbCrLf & "Do you want to add this Part No.?", vbQuestion + vbYesNo, "New Spare Part")
'        flxAdjSP.TextMatrix(ActRow, 6) = txtAdjPartNo.Text
'        txtAdjPartNo.Visible = False
'        flxAdjPartNo.Visible = False
'        If ans = vbYes Then
'
'            'Add Supplier Code
'            'If not Yamaha,Honda,Kawasaki,Suzuki
'            If txtSCODE.Text <> "BZ" And txtSCODE.Text <> "AA" And txtSCODE.Text <> "SP" And txtSCODE.Text <> "CA" Then
'                flxAdjSP.TextMatrix(ActRow, 6) = flxAdjSP.TextMatrix(ActRow, 6) + txtSCODE.Text
'            End If
'
'            SPAddnew.sRow = ActRow
'            SPAddnew.MySource = ""
'            SPAddnew.MyForm = "ADJUSTMENT"
'            SPAddnew.Show vbModal
'            NewSP = True
'        Else
'            flxAdjSP.TextMatrix(ActRow, 6) = ""
'        End If
'        Exit Sub
'    ElseIf flxAdjSP.TextMatrix(ActRow, 7) <> "" And txtAdjPartNo.Text <> "" Then 'NAAY New SP encoded sa row dapat e-remove iyang record sa FlxNewSP
'        If flxAdjSP.TextMatrix(ActRow, 15) <> "" Then
'            If Left$(flxAdjSP.TextMatrix(ActRow, 15), 4) = "NEW:" Then     'New SP already stored in FlxNewSP
'                xRow = Val(Right$(flxAdjSP.TextMatrix(ActRow, 15), Len(flxAdjSP.TextMatrix(ActRow, 15)) - 4))
'                If flxNewSP.TextMatrix(xRow, 0) <> txtAdjPartNo.Text Then
'                    For i = 0 To 10
'                        flxNewSP.TextMatrix(xRow, i) = ""
'                    Next i
'                End If
'            End If
'        End If
'        flxAdjSP.TextMatrix(ActRow, 15) = ""
'        If flxAdjSP.TextMatrix(ActRow, 1) = "NEW" And Val(flxAdjSP.TextMatrix(ActRow, 12)) > 0 Then
'            flxDel.AddItem flxAdjSP.TextMatrix(ActRow, 12)
'            flxAdjSP.TextMatrix(ActRow, 12) = ""
'        End If
'    End If
    
    If flxAdjSP.TextMatrix(ActRow, 7) = "" And txtAdjPartNo.Text <> "" Then
        ans = MsgBox("Unrecognized Part Number!", vbInformation + vbOKOnly, "New Spare Part")
        flxAdjSP.TextMatrix(ActRow, 6) = ""
        txtAdjPartNo.Visible = False
        flxAdjPartNo.Visible = False
    Else
        flxAdjSP.TextMatrix(ActRow, 6) = txtAdjPartNo.Text
        txtAdjPartNo.Visible = False
        flxAdjPartNo.Visible = False
    End If
End Sub

Private Sub txtAdjSO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtAdjSO, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then
        Call txtAdjSO_LostFocus
        DoEvents
        If flxAdjSP.TextMatrix(ActRow, 1) = "NEW" Then
            If flxAdjSP.TextMatrix(ActRow, 6) <> "" And flxAdjSP.TextMatrix(flxAdjSP.Rows - 1, 1) <> "" And flxAdjSP.TextMatrix(flxAdjSP.Rows - 1, 8) <> "" Then
                flxAdjSP.AddItem ""
                flxAdjSP.SetFocus
                flxAdjSP.Col = 1
                flxAdjSP.row = flxAdjSP.Rows - 1
            End If
        ElseIf flxAdjSP.TextMatrix(ActRow, 1) = "REPLACE" And flxAdjSP.TextMatrix(flxAdjSP.Rows - 1, 1) <> "" Then
            If flxAdjSP.TextMatrix(ActRow, 1) <> "" And flxAdjSP.TextMatrix(ActRow, 2) <> "" And flxAdjSP.TextMatrix(ActRow, 6) <> "" And flxAdjSP.TextMatrix(ActRow, 8) <> "" Then
                flxAdjSP.AddItem ""
                flxAdjSP.SetFocus
                flxAdjSP.Col = 1
                flxAdjSP.row = flxAdjSP.Rows - 1
            End If
        End If
        If flxAdjSP.TextMatrix(ActRow, 6) = "" Then
            MsgBox "No Partnumber Found!, Please Select Partnumber", vbInformation + vbOKOnly, "Validation"
            flxAdjSP.TextMatrix(ActRow, 8) = ""
            txtAdjSO.Text = ""
            flxAdjSP.SetFocus
            flxAdjSP.Col = 8
            Exit Sub
        ElseIf flxAdjSP.TextMatrix(ActRow, 8) <> "" Or txtAdjSO.Text <> "" Then
            flxAdjSP.TextMatrix(ActRow, 10) = Format(Val(flxAdjSP.TextMatrix(ActRow, 8)) * Replace(flxAdjSP.TextMatrix(ActRow, 9), ",", ""), "#,###,##0.00")
        End If
    End If
End Sub


Private Sub txtAdjSO_LostFocus()
    Dim ans As Integer, i As Integer, xRow As Integer, j As Integer
    Dim Tmp() As String
    
    If txtAdjSO.Tag = "skiplostfocus" Then
        txtAdjSO.Tag = ""
        Exit Sub
    End If
    DoEvents
    If txtAdjSO.Text <> "" Then
        flxAdjSP.TextMatrix(ActRow, 8) = txtAdjSO.Text
    Else
        flxAdjSP.SetFocus
        flxAdjSP.Col = 8
    End If
    txtAdjSO.Visible = False
End Sub

Private Sub txtAppvdBy_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub

Private Sub txtOrdPartNo_Change()
On Error GoTo LocalErr
    With flxOrdPartNo
        If StatSrch Then
            StatSrch = False
            Exit Sub
        End If
        .Top = txtOrdPartNo.Height + txtOrdPartNo.Top
        .Left = txtOrdPartNo.Left
         .Width = txtOrdPartNo.Width
        .Visible = True
        .ZOrder
            
        If txtOrdPartNo.Text = "" Then .Visible = False: Exit Sub
        rsLoadPO.Find "[PARTNUMBER] LIKE '" + txtOrdPartNo.Text + "%'", , adSearchForward, 1
        If Not rsLoadPO.EOF Then
            If Not PrtBln Then
                txtOrdPartNo.Tag = Len(txtOrdPartNo.Text)
                PrtBln = True
            End If
            
            .TopRow = rsLoadPO.AbsolutePosition - 1
            .row = rsLoadPO.AbsolutePosition - 1
            .Col = 0: .ColSel = 1
            txtOrdPartNo.Text = rsLoadPO!PartNumber
            flxAdjSP.TextMatrix(flxAdjSP.row, 3) = rsLoadPO!Description
            flxAdjSP.TextMatrix(flxAdjSP.row, 4) = rsLoadPO!Order
            flxAdjSP.TextMatrix(flxAdjSP.row, 11) = rsLoadPO!SPAREPARTID
            
            txtOrdPartNo.SelStart = Val(txtOrdPartNo.Tag)
            txtOrdPartNo.SelLength = Len(txtOrdPartNo.Text)
            PrtBln = False
        Else
            flxAdjSP.TextMatrix(ActRow, 2) = ""
            flxAdjSP.TextMatrix(ActRow, 3) = ""
            flxAdjSP.TextMatrix(ActRow, 4) = ""
            flxAdjSP.TextMatrix(ActRow, 11) = ""
            .Visible = False
        End If
   End With
   
   Exit Sub

LocalErr:
    If Err.Number = 30009 Then
        flxOrdPartNo.Visible = False
        txtOrdPartNo.Visible = False
        If rsLoadPO.State = adStateOpen Then rsLoadPO.Close
        tmrLoadPO.Enabled = True
        MsgBox "Program is reloading Part Number" & vbCrLf & "Please wait for a few seconds...", vbInformation + vbOKOnly, "Reloading"
    Else
        flxOrdPartNo.Visible = False
        txtOrdPartNo.Visible = False
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
    End If
End Sub

Private Sub txtOrdPartNo_GotFocus()
    If rsLoadPO.State = adStateClosed Then tmrLoadPO.Enabled = True
    With flxOrdPartNo
        .Top = txtOrdPartNo.Height + txtOrdPartNo.Top
        .Left = txtOrdPartNo.Left
        .Width = txtOrdPartNo.Width
    End With
End Sub

Private Sub txtOrdPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim Tmp() As String
    
    Select Case KeyCode
        Case vbKeyDown
            txtOrdPartNo.Tag = "skiplostfocus"
            flxOrdPartNo.Visible = True
            flxOrdPartNo.SetFocus
        Case vbKeyEscape
            txtOrdPartNo.Text = Empty
            Tmp = Split(flxOrdPartNo.Tag, "|")
            For i = LBound(Tmp) To UBound(Tmp)
                flxOrdPartNo.TextMatrix(ActRow, i + 1) = Tmp(i)
            Next i
            flxOrdPartNo.Tag = ""
            txtOrdPartNo.Tag = "skiplostfocus"
            flxAdjSP.SetFocus
            flxAdjSP.Col = 1
            txtOrdPartNo.Visible = False
            flxOrdPartNo.Visible = False
        Case vbKeyDelete
            StatSrch = True
        Case vbKeyBack
            StatSrch = True
    End Select
End Sub

Private Sub txtOrdPartNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        Call txtOrdPartNo_LostFocus

        flxOrdPartNo.Visible = False
        DoEvents
        If flxAdjSP.TextMatrix(ActRow, 3) <> "" And flxAdjSP.TextMatrix(ActRow, 6) <> "" And flxAdjSP.TextMatrix(ActRow, 8) <> "" And flxAdjSP.TextMatrix(ActRow, 1) = "REPLACE" Then
            txtOrdPartNo.Tag = "NxtRow"
            If ActRow = flxAdjSP.Rows - 1 Then
                flxAdjSP.AddItem ""
                flxAdjSP.SetFocus
                flxAdjSP.row = flxAdjSP.Rows - 1
                flxAdjSP.Col = 1
            End If
        ElseIf flxAdjSP.TextMatrix(ActRow, 3) = "" Then
            flxAdjSP.SetFocus
            flxAdjSP.Col = 2
            flxAdjSP.row = ActRow
        ElseIf flxAdjSP.TextMatrix(ActRow, 6) = "" Then
            flxAdjSP.SetFocus
            flxAdjSP.Col = 6
            flxAdjSP.row = ActRow
        ElseIf flxAdjSP.TextMatrix(ActRow, 8) = "" Then
            flxAdjSP.SetFocus
            flxAdjSP.Col = 8
            flxAdjSP.row = ActRow
        End If
    End If
End Sub

Private Sub txtOrdPartNo_LostFocus()
    Dim ans As Integer, i As Integer, xRow As Integer, j As Integer
    Dim Tmp() As String
    
    If txtOrdPartNo.Tag = "skiplostfocus" Then
        txtOrdPartNo.Tag = ""
        Exit Sub
    End If
    
    flxAdjSP.TextMatrix(ActRow, 2) = txtOrdPartNo.Text
    txtOrdPartNo.Visible = False
    flxOrdPartNo.Visible = False
End Sub

Private Sub txtPONumber_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxAdjSP.SetFocus
        flxAdjSP.row = 2
        flxAdjSP.Col = 1
    End If
End Sub

Private Sub txtPONumber_LostFocus()
Dim rsPO As New ADODB.Recordset
If txtPONumber.Text <> "" Then
    Qry = "call sp_mrcwarehouse.sp_POAdj_getPODetails('" & txtPONumber.Text & "')"
    
    With rsPO
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        
        If .RecordCount > 0 Then
            txtSUPPNAME.Text = IIf(IsNull(!SUPPLIERNAME), "", !SUPPLIERNAME)
            txtSCODE.Text = IIf(IsNull(!SUPPLIERCODE), "", !SUPPLIERCODE)
            tmrLoadPO.Enabled = True
        Else
            MsgBox "" & txtPONumber.Text & " does not exists!", vbInformation + vbOKOnly, "Validation"
            .Close
            Exit Sub
        End If
        .Close
    End With
End If
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub
