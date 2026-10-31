VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form BrnBackOrderPO 
   Caption         =   "Branch Back Order P.O."
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14460
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   14460
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtConfirmNo 
      Height          =   285
      Left            =   8040
      TabIndex        =   43
      Tag             =   "XYF"
      Top             =   8835
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5760
      Top             =   9240
   End
   Begin MSFlexGridLib.MSFlexGrid flxPO 
      Height          =   2895
      Left            =   15120
      TabIndex        =   37
      Top             =   1920
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5106
      _Version        =   393216
      Cols            =   8
      FormatString    =   $"BrnBackOrderPO.frx":0000
   End
   Begin VB.TextBox txtPrnCopy 
      Height          =   285
      Left            =   11520
      TabIndex        =   36
      Top             =   10200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtOrderID 
      Height          =   285
      Left            =   12360
      TabIndex        =   35
      Top             =   10200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtSuppID 
      Height          =   285
      Left            =   13200
      TabIndex        =   34
      Top             =   10200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAprv 
      Height          =   285
      Left            =   2160
      TabIndex        =   29
      Tag             =   "XYF"
      Top             =   9195
      Width           =   2775
   End
   Begin VB.TextBox txtSubmit 
      Height          =   285
      Left            =   2160
      TabIndex        =   28
      Tag             =   "XYF"
      Top             =   8835
      Width           =   2775
   End
   Begin VB.TextBox txtRemarks 
      Height          =   285
      Left            =   2160
      MaxLength       =   150
      TabIndex        =   27
      Tag             =   "XYF"
      Top             =   8400
      Width           =   7575
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   12000
      TabIndex        =   26
      Tag             =   "XYF"
      Top             =   8400
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   1770
      Left            =   120
      TabIndex        =   11
      Tag             =   "XYF"
      Top             =   720
      Width           =   14175
      Begin MSComctlLib.ProgressBar Pbar 
         Height          =   135
         Left            =   1080
         TabIndex        =   41
         Top             =   1560
         Visible         =   0   'False
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   11520
         TabIndex        =   40
         Tag             =   "XYF"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtPODate 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11520
         TabIndex        =   18
         Tag             =   "XYF"
         Top             =   538
         Width           =   2535
      End
      Begin VB.TextBox txtPONumber 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11520
         Locked          =   -1  'True
         TabIndex        =   17
         Tag             =   "XYF"
         Text            =   "<Auto>"
         Top             =   200
         Width           =   2535
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Left            =   11520
         TabIndex        =   16
         Tag             =   "XYF"
         Top             =   876
         Width           =   2535
      End
      Begin VB.TextBox txtScode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   15
         Tag             =   "XYF"
         Top             =   200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtSaddress 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   14
         Tag             =   "XYF"
         Top             =   538
         Width           =   3735
      End
      Begin VB.ComboBox cboSupplier 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2040
         TabIndex        =   13
         Tag             =   "XYF"
         Top             =   192
         Width           =   3735
      End
      Begin VB.TextBox txtTerms 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   876
         Width           =   3735
      End
      Begin lvButton.lvButtons_H cmdBBOparts 
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Tag             =   "xyf"
         ToolTipText     =   "Load BBO Parts"
         Top             =   1320
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   661
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
         ImgAlign        =   4
         Image           =   "BrnBackOrderPO.frx":0042
         cBack           =   -2147483633
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "STATUS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   10320
         TabIndex        =   39
         Tag             =   "XYF"
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "P.O. DATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   10320
         TabIndex        =   24
         Tag             =   "XYF"
         Top             =   568
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "P.O. NUMBER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   9960
         TabIndex        =   23
         Tag             =   "XYF"
         Top             =   230
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ORDER TYPE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   10320
         TabIndex        =   22
         Tag             =   "XYF"
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "SUPPLIER NAME:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   21
         Tag             =   "XYF"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "TERMS:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Tag             =   "XYF"
         Top             =   906
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "SUPPLIERADDRESS:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   568
         Width           =   1455
      End
   End
   Begin lvButton.lvButtons_H cmdNew 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&New Order"
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
      Image           =   "BrnBackOrderPO.frx":0A54
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdSummary 
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "BrnBackOrderPO.frx":0FEE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Print Order"
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
      Image           =   "BrnBackOrderPO.frx":1588
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "E&xit Order"
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
      Image           =   "BrnBackOrderPO.frx":1B22
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   375
      Left            =   8400
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "BrnBackOrderPO.frx":20BC
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   10080
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "BrnBackOrderPO.frx":2656
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdEdit 
      Height          =   375
      Left            =   1678
      TabIndex        =   6
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Edit Order"
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
      Image           =   "BrnBackOrderPO.frx":2BF0
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxParts 
      Height          =   5775
      Left            =   120
      TabIndex        =   25
      Tag             =   "XYF"
      Top             =   2520
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   10186
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      HighLight       =   2
      ScrollBars      =   2
      FormatString    =   "^    |<PART NUMBER   |<DESCRIPTION  |^QTY.   |>UNIT PRICE|>TOTAL PRICE"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "CONFIRMATION NO."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   6120
      TabIndex        =   42
      Tag             =   "XYF"
      Top             =   8880
      Width           =   1695
   End
   Begin VB.Label lbl2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "REMARKS:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1080
      TabIndex        =   33
      Tag             =   "XYF"
      Top             =   8430
      Width           =   795
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "PREPARED BY:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   32
      Tag             =   "XYF"
      Top             =   8835
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "APPROVED BY:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   31
      Tag             =   "XYF"
      Top             =   9210
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "TOTAL:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   10920
      TabIndex        =   30
      Tag             =   "XYF"
      Top             =   8415
      Width           =   975
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   2160
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14700
      m_MDI_Height    =   10380
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "MOTO>R REPUBLIC CORP."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Tag             =   "TLF"
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SPARE PARTS ORDER FORM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   2
      Left            =   11430
      TabIndex        =   7
      Tag             =   "TRF"
      Top             =   480
      Width           =   2745
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "MOTO>R REPUBLIC CORP."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   150
      TabIndex        =   9
      Tag             =   "TLF"
      Top             =   510
      Width           =   2895
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SPARE PARTS ORDER FORM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   11460
      TabIndex        =   10
      Tag             =   "TRF"
      Top             =   510
      Width           =   2745
   End
   Begin VB.Menu mnuRows 
      Caption         =   "Branch Back Order PO"
      Visible         =   0   'False
      Begin VB.Menu mnuDel 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "BrnBackOrderPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSupplier As New ADODB.Recordset

Public LockFlx As Boolean, LoadRef As Boolean
Dim ActCol As Integer, ActRow As Integer
Dim AddNew As Boolean, SuppSrch As Boolean

Private Sub BtnSaveCancel(Bln As Boolean)
    Dim i As Integer
    
    If Bln Then
        cmdNew.Visible = False
        cmdEdit.Visible = False
        cmdSummary.Visible = False
        cmdPrint.Visible = False
        cmdExit.Visible = False
        cmdSave.Visible = True
        cmdCancel.Visible = True
        cmdSave.Left = 0
        cmdCancel.Left = cmdSave.Width
        
    Else
        cmdNew.Visible = True
        cmdEdit.Visible = True
        cmdSummary.Visible = True
        cmdPrint.Visible = True
        cmdExit.Visible = True
        cmdSave.Visible = False
        cmdCancel.Visible = False
    End If
End Sub

Private Sub ClearEntries()
    cboSupplier.Text = ""
    txtScode.Text = ""
    txtSaddress.Text = ""
    txtSuppID.Text = ""
    txtPODate.Text = ""
    cboType.Text = ""
    txtRemarks.Text = ""
    txtSubmit.Text = ""
    txtAprv.Text = ""
    txtPrnCopy.Text = ""
    txtOrderID.Text = ""
    txtTerms.Text = ""
    flxParts.Rows = 1
    flxParts.Rows = 2
    flxPO.Rows = 1
End Sub

Private Sub LockData(Bln As Boolean)
    Frame1.Enabled = Not Bln
    txtRemarks.Locked = Bln
    txtSubmit.Locked = Bln
    txtAprv.Locked = Bln
    LockFlx = Bln
End Sub

Public Sub DisplaySPOrder(MyRef As String)
    Dim rsOrder As New ADODB.Recordset
    Dim rsITR As New ADODB.Recordset
    Dim rsPO As New ADODB.Recordset
    Dim Qry As String, i As Integer
    
    Qry = "SELECT SPORDER.* From SPORDER WHERE SPORDER.REFERENCE = '" & MyRef & "';"
    
    With rsOrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Call ClearEntries
            .MoveFirst
            txtPODate.Text = Format(!ORDERDATE, "mm/dd/yyyy")
            txtPONumber.Text = !REFERENCE
            cboType.Text = !ORDERTYPE
            txtConfirmNo.Text = IIf(IsNull(!CONFIRMNO), "", !CONFIRMNO)
            cboStatus.Text = IIf(IsNull(!POSTATUS), "", !POSTATUS)
            txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
            txtTerms.Text = IIf(IsNull(!Terms), "", !Terms)
            txtSubmit.Text = !SubmittedBy
            txtAprv.Text = !ApprovedBy
            txtPrnCopy.Text = IIf(IsNull(!PrintCopy), 0, !PrintCopy)
            txtSuppID.Text = IIf(IsNull(!supplierid), "", !supplierid)
            
            txtOrderID.Text = !SPORDERID
        End If
        .Close
    End With
    If txtSuppID.Text <> "" Then Call DisplaySupplier
    'mySQL
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.`ORDER`, SPITR.PRICE, SPITR.`ORDER`*SPITR.PRICE AS TOTAL, SPITR.SPAREPARTID, SPAREPARTS.APPLICMODEL " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE SPITR.REFERENCE='" & MyRef & "' and spitr.TRANSTYPE LIKE 'ORDER%' ORDER BY SPITR.SEQNO;"
    
    
    With rsITR
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxParts.Rows = 1
            .MoveFirst
            pBar.Visible = True
            pBar.Value = 0
            pBar.Max = .RecordCount
            Do While Not .EOF
                flxParts.AddItem flxParts.Rows & vbTab & !PartNumber & vbTab & !Description & vbTab & !Order & vbTab & Format$(!Price, "#,##0.00") & vbTab & Format$(!Total, "#,##0.00") & vbTab & !SPAREPARTID
                
                .MoveNext
                pBar.Value = pBar.Value + 1
                DoEvents
            Loop
        End If
        .Close
        pBar.Visible = False
    End With
    Qry = "SELECT backorder.SPAREPARTID, backorder.SUPPORDERREF, backorder.TRANSDATE, backorder.BRANCHNAME, backorder.BACKORDERQTY, backorder.POQTY, backorder.BACKORDERID, backorder.REFERENCE " & _
        "From backorder Where backorder.SUPPORDERID = " & Val(txtOrderID.Text) & " AND backorder.SUPPORDERREF = '" & MyRef & "';"
    With rsPO
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            pBar.Visible = True
            pBar.Value = 0
            pBar.Max = .RecordCount
            Do While Not .EOF
                flxPO.AddItem vbTab & !SPAREPARTID & vbTab & !REFERENCE & vbTab & Format(!TRANSDATE, "MM/DD/YYYY") & vbTab & !branchName & vbTab & !BACKORDERQTY & vbTab & !POQTY & vbTab & !BACKORDERID
                .MoveNext
                pBar.Value = pBar.Value + 1
                DoEvents
            Loop
        End If
        .Close
        pBar.Visible = False
    End With
    
    Call LockData(True)
    Call ComputeTotal
End Sub

Private Sub DisplaySupplier()
    Dim rsSupp As New ADODB.Recordset
    Dim Qry As String
    
    'mySQL
    Qry = "SELECT SUPPLIER.SUPPLIERNAME, SUPPLIER.SUPPLIERCODE, SUPPLIER.ADDRESS " & _
          "From SUPPLIER WHERE SUPPLIER.SUPPLIERID=" & txtSuppID.Text & ";"
    
    With rsSupp
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            SuppSrch = True
            cboSupplier.Text = !SUPPLIERNAME
            txtSaddress.Text = !Address
            txtScode.Text = !SUPPLIERCODE
        End If
        .Close
    End With
    
End Sub

Public Sub ComputeTotal()
    Dim Ttl As Single, i As Integer
    Dim TtlCost As Single
    
    Ttl = 0
    With flxParts
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 3)) > 0 And Val(Replace(.TextMatrix(i, 4), ",", "")) > 0 Then     'Val(.TextMatrix(i, 5)) = 0 And
                TtlCost = Val(.TextMatrix(i, 3)) * Val(Replace(.TextMatrix(i, 4), ",", ""))
                .TextMatrix(i, 5) = Format$(TtlCost, "###,##0.00")
            End If
            Ttl = Ttl + Val(Replace(.TextMatrix(i, 5), ",", ""))
        Next i
    End With
    txtTotal.Text = Format$(Ttl, "#,##0.00")
    
End Sub

Function ValidateSPOrder() As String
    Dim vMsg As String, Qry As String
    Dim i As Integer, Cntr As Integer
    
    vMsg = ""
    
    If Val(txtSuppID.Text) = 0 Then vMsg = "Invalid entry on field Supplier"
    If Not IsDate(txtPODate.Text) Then vMsg = vMsg & vbCrLf & "Invalid P.O. Date"
    'If txtPONumber.Text = "" Then
    '    vMsg = vMsg & vbCrLf & "Incomplete entry on field P.O. Number"
    'ElseIf AddNew Then
    '    Qry = "SELECT SPORDER.REFERENCE From SPORDER WHERE SPORDER.REFERENCE='" & txtPONumber.Text & "';"
    '    If GetFieldValue(Qry) <> "" Then
    '        vMsg = vMsg & vbCrLf & "P.O. Number " & txtPONumber.Text & " already exist!"
    '    End If
    'End If
    If cboType.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Order Type"
    If cboStatus.ListIndex < 0 Then vMsg = vMsg & vbCrLf & "Invalid entry on field STATUS"
    If txtSubmit.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field PREPARED BY"
    If txtAprv.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field APPROVED BY"
    If flxParts.Rows > 2 Then
        If flxParts.TextMatrix(flxParts.Rows - 1, 1) = "" And flxParts.TextMatrix(flxParts.Rows - 1, 2) = "" And flxParts.TextMatrix(flxParts.Rows - 1, 3) = "" Then flxParts.RemoveItem flxParts.Rows - 1
    Else
        If flxParts.TextMatrix(1, 1) = "" And flxParts.TextMatrix(1, 3) = "" Then vMsg = vMsg & vbCrLf & "No Spare Parts ordered!"
    End If
    ValidateSPOrder = vMsg
    
End Function


Private Sub cboStatus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdBBOparts.SetFocus
End Sub

Private Sub cboSupplier_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
    If SuppSrch Then
        SuppSrch = False
        Exit Sub
    End If
    With cboSupplier
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsSupplier.MoveFirst
            rsSupplier.Move i
            txtSuppID.Text = rsSupplier!supplierid
            txtScode.Text = IIf(IsNull(rsSupplier!SUPPLIERCODE), "", rsSupplier!SUPPLIERCODE)
            txtSaddress.Text = rsSupplier!Address
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                SuppSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtSuppID.Text = ""
            txtScode.Text = ""
            txtSaddress.Text = ""
        End If
    End With
End Sub

Private Sub cboSupplier_Click()
    Call cboSupplier_Change
End Sub

Private Sub cboSupplier_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboSupplier.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub cboSupplier_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTerms.SetFocus
End Sub



Private Sub cboType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdBBOparts.SetFocus
End Sub

Private Sub cmdBBOparts_Click()
    BrnBBOparts.Show vbModal
End Sub

Private Sub cmdCancel_Click()
    Dim ans As Integer
    
    ans = MsgBox("Do you want to Cancel Order?", vbQuestion + vbYesNo, "Cancel")
    If ans = vbNo Then Exit Sub
    Call BtnSaveCancel(False)
    
    If AddNew Then
        Call ClearEntries
    Else
        txtPODate.Locked = False
        txtPONumber.Locked = False
        txtTerms.Locked = False
        cboType.Locked = False
        cboSupplier.Locked = False
        'Call DisplaySPOrder(txtPONumber.Text)
    End If
    Call LockData(True)

End Sub

Private Sub cmdEdit_Click()

    If cboStatus.Text = "FINAL" Then
        MsgBox "Editing is not allowed because P.O. is already FINAL!", vbInformation + vbOKOnly, "Edit Denied"
        Exit Sub
    ElseIf txtPONumber.Text = "" Then
        MsgBox "No available P.O. to Edit", vbInformation + vbOKOnly, "Confirm P.O."
        Exit Sub
    End If
            
    AddNew = False
    Call LockData(False)
    Call BtnSaveCancel(True)
    txtPONumber.Locked = True
    cboType.Locked = True
    LockFlx = False
End Sub

Private Sub cmdExit_Click()
    Warehouse.sBar.Panels(6).Text = ""
    Unload Me
End Sub

Private Sub cmdNew_Click()
    Call ClearEntries
    Call LockData(False)
    Call BtnSaveCancel(True)
    txtPODate.Text = Date
    txtSubmit.Text = GetUserName
    txtAprv.Text = GetSettings("MANAGER")
    AddNew = True
    LockFlx = False
    flxParts.TextMatrix(1, 0) = 1
    cboType.ListIndex = 3
    cboType.Locked = True
    cboStatus.ListIndex = 1
    cboSupplier.SetFocus
End Sub

Function GetOrderNo() As String
    Dim rsOrder As New ADODB.Recordset
    Dim Cntr As Integer, RefExist As Boolean
    Dim Qry As String, MyOrdNo As String, MyRef As String
    
    '***Access
    'Qry = "SELECT SPORDER.SPORDERID, SPORDER.REFERENCE, SPORDER.ORDERDATE, Month([ORDERDATE]) AS MYMONTH " & _
          "From SPORDER WHERE (((Month([ORDERDATE]))=" & Month(txtPODate.Text) & ") AND ((SPORDER.SUPPLIERID)=" & txtSuppID.Text & ")) ORDER BY SPORDER.REFERENCE DESC;"
    'mySQL
    Qry = "SELECT SPORDER.SPORDERID, SPORDER.REFERENCE, SPORDER.ORDERDATE, Month(SPORDER.ORDERDATE) AS MYMONTH " & _
          "From SPORDER WHERE Month(SPORDER.ORDERDATE)='" & Month(txtPODate.Text) & "' AND SPORDER.SUPPLIERID=" & txtSuppID.Text & " ORDER BY SPORDER.REFERENCE DESC;"
    
    With rsOrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        Cntr = .RecordCount
        .Close
    End With
    If InStr(cboSupplier.Text, "HONDA") > 0 Then
        MyOrdNo = Format$(txtPODate.Text, "yyyy-mm-") & Format$(Cntr + 1, "000")
    Else
        MyOrdNo = Format$(txtPODate.Text, "yy-mm-") & Format$(Cntr + 1, "00") & txtScode.Text
    End If
    RefExist = False
    Do
        '***ACCESS
        'Qry = "SELECT SPORDER.REFERENCE From SPORDER WHERE (((SPORDER.REFERENCE)='" & MyOrdNo & "'));"
        'mySQL
        Qry = "SELECT SPORDER.REFERENCE From SPORDER WHERE SPORDER.REFERENCE='" & MyOrdNo & "';"
        If CtrlNumExist(Qry) Then
            Cntr = Cntr + 1
            If InStr(cboSupplier.Text, "HONDA") > 0 Then
                MyOrdNo = Format$(txtPODate.Text, "yyyy-mm-") & Format$(Cntr + 1, "000")
            Else
                MyOrdNo = Format$(txtPODate.Text, "yy-mm-") & Format$(Cntr + 1, "00") & txtScode.Text
            End If
            RefExist = True
        Else
            RefExist = False
        End If
    Loop While RefExist
    GetOrderNo = MyOrdNo
    
End Function

Private Sub cmdPrint_Click()
    Dim ans As Integer, i As Integer
    Dim PrnVal(0 To 21) As String, Qry As String, xBrn As String
    'Add Additional Genuine supplier
    xBrn = SysBranch
    SysBranch = "LYR Motorparts Corporation"
    SysAddress = "LYR Corporate Bldg. Gante Avenue, Tagum City"
    If InStr(cboSupplier.Text, "HONDA PHILS.") > 0 Or InStr(cboSupplier.Text, "YMPI") > 0 Or InStr(cboSupplier.Text, "KAWASAKI PHILS.") > 0 Or InStr(cboSupplier.Text, "SUZUKI PHILS.") > 0 Then
        ans = MsgBox("Do you want to Print P.O. with Amount?", vbQuestion + vbYesNo, "Print")
        If ans = vbYes Then
            Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.`ORDER`, SPITR.PRICE " & _
                "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.TRANSTYPE Like 'ORDER%' AND SPITR.REFERENCE = '" & txtPONumber.Text & "' ORDER BY SPITR.SEQNO;"
            
            PRNmc.QryPrn = Qry
            PRNmc.ParamValue1 = txtPONumber.Text
            PRNmc.ParamValue2 = cboType.Text
            PRNmc.ParamValue3 = txtPODate.Text
            PRNmc.MySource = "SPPOPRICE"
            PRNmc.Show
        Else
            Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.`ORDER` " & _
                  "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.TRANSTYPE Like 'ORDER%' AND SPITR.REFERENCE = '" & txtPONumber.Text & "' ORDER BY SPITR.SEQNO;"
            
            PRNmc.QryPrn = Qry
            PRNmc.ParamValue1 = txtPONumber.Text
            PRNmc.ParamValue2 = cboType.Text
            PRNmc.ParamValue3 = txtPODate.Text
            PRNmc.MySource = "SPPO"
            PRNmc.Show
        End If
    Else
        ans = MsgBox("Do you want to Print with Part Number?", vbQuestion + vbYesNo, "Print Part Number")
        If ans = vbYes Then
            Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.`ORDER`, SPITR.PRICE " & _
                "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.TRANSTYPE Like 'ORDER%' AND SPITR.REFERENCE = '" & txtPONumber.Text & "' ORDER BY SPITR.SEQNO;"
            
            PRNmc.QryPrn = Qry
            PRNmc.ParamValue1 = cboSupplier.Text
            PRNmc.ParamValue2 = txtPONumber.Text
            PRNmc.ParamValue3 = txtPODate.Text
            PRNmc.MySource = "SPPOreplPnum"
            PRNmc.Show
            Exit Sub
        End If
        ans = MsgBox("Do you want to Print P.O. with Amount?", vbQuestion + vbYesNo, "Print")
        If ans = vbNo Then
            Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.`ORDER` " & _
                  "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.TRANSTYPE Like 'ORDER%' AND SPITR.REFERENCE = '" & txtPONumber.Text & "' ORDER BY SPITR.SEQNO;"
            
            PRNmc.QryPrn = Qry
            PRNmc.ParamValue1 = cboSupplier.Text
            PRNmc.ParamValue2 = txtPONumber.Text
            PRNmc.ParamValue3 = txtPODate.Text
            PRNmc.MySource = "SPPOrepl"
            PRNmc.Show
        Else
            Qry = "SELECT SPAREPARTS.DESCRIPTION, SPITR.`ORDER`, SPITR.PRICE " & _
                "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.TRANSTYPE Like 'ORDER%' AND SPITR.REFERENCE = '" & txtPONumber.Text & "' ORDER BY SPITR.SEQNO;"
            
            PRNmc.QryPrn = Qry
            PRNmc.ParamValue1 = cboSupplier.Text
            PRNmc.ParamValue2 = txtPONumber.Text
            PRNmc.ParamValue3 = txtPODate.Text
            PRNmc.MySource = "SPPOreplAmt"
            PRNmc.Show
        End If
    End If
    SysBranch = xBrn
End Sub

Private Sub cmdSave_Click()
    Dim myMsg As String, myLastUpdate As String, Qry As String, myRem As String
    Dim ans As Integer
    
    myMsg = ValidateSPOrder
    If myMsg <> "" Then
        MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
        Exit Sub
    End If
    ans = MsgBox("Do you want to save Purchase Order?", vbQuestion + vbYesNo, "Save")
    If ans = vbNo Then Exit Sub
    myLastUpdate = Format(Now, "yyyy-mm-dd hh:mm:ss")
    If AddNew Then
        txtPONumber.Text = GetOrderNo
        myRem = Replace(txtRemarks.Text, "'", "`")
        LYRcon.Execute "Insert Into sporder(REFERENCE, ORDERTYPE, ORDERDATE, SUBMITTEDBY, APPROVEDBY, UPDATEDBY, LASTUPDATE, REMARKS, TERMS, SUPPLIERID, POSTATUS, CONFIRMNO)" & _
            "VALUES('" & Replace(txtPONumber.Text, "'", "`") & "', '" & cboType.Text & "', '" & Format(txtPODate.Text, "yyyy-mm-dd") & "', '" & txtSubmit.Text & "', '" & txtAprv.Text & "', " & UserID & ", '" & myLastUpdate & "', " & _
            "'" & myRem & "', '" & txtTerms.Text & "', " & Val(txtSuppID.Text) & ", '" & cboStatus.Text & "', '" & txtConfirmNo.Text & "');"
        Qry = "select sporder.SPORDERID from sporder where sporder.REFERENCE = '" & txtPONumber.Text & "' AND sporder.LASTUPDATE = '" & myLastUpdate & "';"
        txtOrderID.Text = GetFieldValue(Qry)
    Else
        LYRcon.Execute "UPDATE sporder SET ORDERDATE = '" & Format(txtPODate.Text, "yyyy-mm-dd") & "', SUBMITTEDBY = '" & txtSubmit.Text & "', APPROVEDBY = '" & txtAprv.Text & "', " & _
        "REMARKS = '" & Replace(txtRemarks.Text, "'", "`") & "', UPDATEDBY = " & UserID & ", SUPPLIERID = " & Val(txtSuppID.Text) & ", TERMS = '" & txtTerms.Text & "', LASTUPDATE = '" & myLastUpdate & "', POSTATUS = '" & cboStatus.Text & "', CONFIRMNO = '" & txtConfirmNo.Text & "' " & _
        "Where sporder.SPORDERID = " & Val(txtOrderID.Text) & ";"
        'Remove SPITR and backorder records
        LYRcon.Execute "DELETE From spitr Where spitr.REFERENCE = '" & txtPONumber.Text & "' AND spitr.TRANSTYPE = 'ORDER TEMP'"
        LYRcon.Execute "UPDATE backorder SET SUPPORDERREF = NULL, SUPPORDERID = NULL, SUPPOSTATUS = '', POQTY = 0 " & _
            "Where backorder.SUPPORDERID = " & Val(txtOrderID.Text) & ";"
        
    End If
    pBar.Visible = True
    pBar.Value = 0
    pBar.Max = flxParts.Rows + flxPO.Rows
    With flxParts
        If txtConfirmNo.Text <> "" Then
            For i = 1 To .Rows - 1
                LYRcon.Execute "Insert Into spitr( REFERENCE, SPAREPARTID, TRANSDATE, TRANSTYPE, `ORDER`, " & _
                   "PRICE, SOURCE, UPDATEDBY, LASTUPDATE, SEQNO) VALUES( '" & txtPONumber.Text & "', " & _
                   "" & flxParts.TextMatrix(i, 6) & ", '" & Format(txtPODate.Text, "yyyy-mm-dd") & "', " & _
                   "'ORDER', " & Val(flxParts.TextMatrix(i, 3)) & ", " & Val(Replace(flxParts.TextMatrix(i, 4), ",", "")) & ", " & _
                   "'" & cboSupplier.Text & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', " & i & ");"
                pBar.Value = pBar.Value + 1
                DoEvents
            Next i
        Else
           For i = 1 To .Rows - 1
                LYRcon.Execute "Insert Into spitr( REFERENCE, SPAREPARTID, TRANSDATE, TRANSTYPE, `ORDER`, " & _
                   "PRICE, SOURCE, UPDATEDBY, LASTUPDATE, SEQNO) VALUES( '" & txtPONumber.Text & "', " & _
                   "" & flxParts.TextMatrix(i, 6) & ", '" & Format(txtPODate.Text, "yyyy-mm-dd") & "', " & _
                   "'ORDER TEMP', " & Val(flxParts.TextMatrix(i, 3)) & ", " & Val(Replace(flxParts.TextMatrix(i, 4), ",", "")) & ", " & _
                   "'" & cboSupplier.Text & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', " & i & ");"
                pBar.Value = pBar.Value + 1
                DoEvents
            Next i
        End If
    End With
    With flxPO
        
        For i = 1 To .Rows - 1
            LYRcon.Execute "UPDATE backorder SET SUPPORDERREF = '" & txtPONumber.Text & "', SUPPORDERID = " & Val(txtOrderID.Text) & ", SUPPOSTATUS = '" & cboStatus.Text & "', POQTY = " & Val(.TextMatrix(i, 6)) & " " & _
                "Where backorder.BACKORDERID = " & Val(.TextMatrix(i, 7)) & " AND backorder.SPAREPARTID = " & Val(.TextMatrix(i, 1)) & ";"
            pBar.Value = pBar.Value + 1
            DoEvents
        Next i
       
    End With
        
    pBar.Visible = False
    Call BtnSaveCancel(False)
    MsgBox "Spare Parts Order successfully saved!", vbInformation + vbOKOnly, "Save"
    LockFlx = True
    
End Sub

Private Sub cmdSummary_Click()
    
    SPorderSum.Show
    Unload Me
End Sub



Private Sub flxParts_DblClick()
    Dim mySPID As Integer, i As Integer
    
    If ActCol = 3 And Val(flxParts.TextMatrix(ActRow, 3)) > 0 Then
        mySPID = Val(flxParts.TextMatrix(ActRow, 6))
        With flxPO
            BrnBBOref.flxREF.Rows = 1
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 1)) = mySPID Then
                    BrnBBOref.flxREF.AddItem .TextMatrix(i, 2) & vbTab & .TextMatrix(i, 3) & vbTab & .TextMatrix(i, 4) & vbTab & .TextMatrix(i, 5) & vbTab & .TextMatrix(i, 6)
                End If
            Next i
        End With
        LoadRef = True
        BrnBBOref.lblTotal.Caption = "TOTAL ORDER: " & flxParts.TextMatrix(ActRow, 3)
        BrnBBOref.Caption = "(" & flxParts.TextMatrix(ActRow, 1) & ") " & flxParts.TextMatrix(ActRow, 2)
        BrnBBOref.Show
    End If
End Sub

Private Sub flxParts_EnterCell()
    ActCol = flxParts.Col
    ActRow = flxParts.row
    
    If LoadRef Then
        ActCol = 3
        Call flxParts_DblClick
    End If
End Sub

Private Sub flxParts_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not LockFlx And ActRow > 0 Then
        If Button = vbRightButton Then
            Call PopupMenu(mnuRows)
        End If
    End If
End Sub



Private Sub Form_Load()
   
    cboType.AddItem "(01) - URGENT"
    cboType.AddItem "(04) - STOCK"
    cboType.AddItem "(05) - WARRANTY"
    cboType.AddItem "(06) - BRANCH B.O."
    
    cboStatus.AddItem "FINAL"
    cboStatus.AddItem "TEMPORARY"
    LockFlx = True
    Call LockData(True)
    Warehouse.sBar.Panels(6).Text = "Spare Parts Order Form"
    Timer1.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsSupplier.State = adStateOpen Then rsSupplier.Close
    Warehouse.sBar.Panels(6).Text = ""
End Sub

Private Sub mnuDel_Click()
    Dim ans As Integer, i As Integer, SPid As Integer
    
    ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If ans = vbYes Then
        If ActRow = 1 And flxParts.Rows = 2 Then
            flxParts.Rows = 1
            flxParts.Rows = 2
            flxParts.TextMatrix(1, 0) = 1
        Else
            'Get ID and remove if flxPO
            SPid = Val(flxParts.TextMatrix(ActRow, 6))
            For i = flxPO.Rows - 1 To 1 Step -1 'Delete SpareParts in FlxPO before deleting the Row
                If Val(flxPO.TextMatrix(i, 1)) = SPid Then  'If SPID is found
                    If i = 1 And flxPO.Rows = 2 Then    'If there is only 1 record in Flxpo
                        flxPO.Rows = 1
                    Else
                        flxPO.RemoveItem i
                    End If
                End If
            Next i
            flxParts.RemoveItem ActRow
            For i = 1 To flxParts.Rows - 1      'Reset Row number
                flxParts.TextMatrix(i, 0) = i
            Next i
        End If
        Call ComputeTotal
    End If
End Sub

Private Sub RXResizer1_ResizeComplete()
    With flxParts
        
        .Cols = 14
        .ColWidth(0) = 800
        .ColWidth(1) = (.Width - 800) * 0.23
        .ColWidth(2) = (.Width - 800) * 0.3
        .ColWidth(3) = (.Width - 800) * 0.1
        .ColWidth(4) = (.Width - 800) * 0.17
        .ColWidth(5) = (.Width - 800) * 0.17
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        If .Cols > 10 Then .ColWidth(10) = 0
        If .Cols > 11 Then .ColWidth(11) = 0
        If .Cols > 12 Then .ColWidth(12) = 0
        .TextMatrix(0, 1) = "PART NUMBER"
        .TextMatrix(0, 2) = "DESCRIPTION"
        .TextMatrix(0, 3) = "QTY."
        .TextMatrix(0, 4) = "UNIT PRICE"
        .TextMatrix(0, 5) = "TOTAL PRICE"
        .TextMatrix(0, 6) = ""
        .TextMatrix(0, 7) = ""
        .TextMatrix(0, 8) = ""
        '.ColAlignment(4) = 7
        '.ColAlignment(5) = 7
    End With
End Sub

Private Sub Timer1_Timer()
    Dim Qry As String
    
    Qry = "SELECT SUPPLIER.SUPPLIERID, SUPPLIER.SUPPLIERNAME, SUPPLIER.SUPPLIERCODE, SUPPLIER.ADDRESS " & _
          "From SUPPLIER WHERE supplier.SUPPLIERNAME NOT LIKE '%(INACTIVE)%' ORDER BY SUPPLIER.SUPPLIERNAME;"
    With rsSupplier
        
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            'cboSupplier.Clear
            Do While Not .EOF
                cboSupplier.AddItem !SUPPLIERNAME
                .MoveNext
            Loop
        End If
    End With
    Timer1.Enabled = False
End Sub


Private Sub txtPODate_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789/-"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        cboType.SetFocus
    End If

End Sub

Private Sub txtPODate_LostFocus()
    If txtPODate.Text <> "" Then
        If Not IsDate(txtPODate.Text) Then
            MsgBox "Invalid P.O. Date! Please enter a valid date.", vbExclamation + "Validation Error"
        End If
    End If
End Sub

Private Sub txtPONumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPODate.SetFocus
End Sub

Private Sub txtTerms_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPODate.SetFocus
End Sub
