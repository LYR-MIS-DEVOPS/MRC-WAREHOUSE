VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRFP 
   Caption         =   "REQUEST FOR PAYMENTS"
   ClientHeight    =   9810
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   14880
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtFlxValueRFP 
      Height          =   285
      Left            =   360
      TabIndex        =   52
      Tag             =   "XYZ"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   37
      Tag             =   "XYZ"
      Top             =   8520
      Width           =   13215
      Begin VB.TextBox txtCheckNo 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   42
         Tag             =   "XYZ"
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox cboChckedBy 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   41
         Tag             =   "XYZ"
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtApprvdBy 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   11760
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Tag             =   "XYZ"
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtNotedBy 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   11760
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   39
         Tag             =   "XYZ"
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtPrepBy 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Tag             =   "XYZ"
         Top             =   240
         Width           =   3015
      End
      Begin MSMask.MaskEdBox txtChckNoDate 
         Height          =   315
         Left            =   4320
         TabIndex        =   43
         Tag             =   "XYZ"
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12632256
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "CHECK NO.:"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Tag             =   "XYZ"
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "MM/DD/YYYY"
         Height          =   255
         Left            =   6120
         TabIndex        =   49
         Tag             =   "XYZ"
         Top             =   915
         Width           =   1095
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "CHECK  DATE:"
         Height          =   255
         Left            =   3240
         TabIndex        =   48
         Tag             =   "XYZ"
         Top             =   855
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "APPROVED BY:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10440
         TabIndex        =   47
         Tag             =   "XYZ"
         Top             =   735
         Width           =   1455
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "NOTED BY:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10800
         TabIndex        =   46
         Tag             =   "XYZ"
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "CHECKED BY:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   45
         Tag             =   "XYZ"
         Top             =   285
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "PREPARED BY:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Tag             =   "XYZ"
         Top             =   285
         Width           =   1455
      End
   End
   Begin VB.Timer tmrGetRecord 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9120
      Top             =   0
   End
   Begin VB.ComboBox cboFlxParticular 
      Height          =   315
      ItemData        =   "frmRFP.frx":0000
      Left            =   6960
      List            =   "frmRFP.frx":0002
      TabIndex        =   36
      Top             =   6840
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.TextBox txtFlxDeductReb 
      Height          =   285
      Left            =   12480
      TabIndex        =   35
      Top             =   6840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox cboFlxType 
      Height          =   315
      ItemData        =   "frmRFP.frx":0004
      Left            =   5040
      List            =   "frmRFP.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   6840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtFlxValueDeduct 
      Height          =   285
      Left            =   720
      TabIndex        =   33
      Top             =   6840
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Frame framHoldButtton 
      Caption         =   "DEDUCTION / ADD - BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   22
      Tag             =   "xyz"
      Top             =   6240
      Width           =   14655
      Begin VB.TextBox txtNetPay 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   12600
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Tag             =   "XYZ"
         Text            =   "0.00"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtTotalDed 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   12600
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Tag             =   "XYZ"
         Text            =   "0.00"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtTotAdd 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   4440
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Tag             =   "XYZ"
         Text            =   "0.00"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox txtTotLess 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Tag             =   "XYZ"
         Text            =   "0.00"
         Top             =   1560
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid flxDeduct 
         Height          =   1215
         Left            =   120
         TabIndex        =   6
         Tag             =   "XYZ"
         Top             =   240
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   2143
         _Version        =   393216
         Cols            =   5
         RowHeightMin    =   315
         BackColor       =   16777215
         BackColorBkg    =   12632256
         HighLight       =   2
         GridLines       =   2
         AllowUserResizing=   3
         FormatString    =   $"frmRFP.frx":001D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL NET PAY:"
         Height          =   255
         Left            =   11280
         TabIndex        =   31
         Tag             =   "XYZ"
         Top             =   1860
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL DEDUCTION:"
         Height          =   255
         Left            =   11040
         TabIndex        =   30
         Tag             =   "XYZ"
         Top             =   1485
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL ADD:"
         Height          =   255
         Left            =   3480
         TabIndex        =   27
         Tag             =   "XYZ"
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL LESS:"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Tag             =   "XYZ"
         Top             =   1560
         Width           =   1095
      End
   End
   Begin VB.Frame framHoldMid 
      Height          =   4455
      Left            =   120
      TabIndex        =   16
      Tag             =   "xyz"
      Top             =   1680
      Width           =   14655
      Begin VB.TextBox txtTotInvAmnt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   12600
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Tag             =   "XYZ"
         Text            =   "0.00"
         Top             =   3960
         Width           =   1935
      End
      Begin MSFlexGridLib.MSFlexGrid flxRFP 
         Height          =   3735
         Left            =   120
         TabIndex        =   5
         Tag             =   "XYZ"
         Top             =   240
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   6588
         _Version        =   393216
         Cols            =   10
         RowHeightMin    =   315
         BackColor       =   16777215
         BackColorBkg    =   12632256
         WordWrap        =   -1  'True
         HighLight       =   2
         GridLines       =   2
         AllowUserResizing=   3
         FormatString    =   $"frmRFP.frx":0127
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL:"
         Height          =   255
         Left            =   12000
         TabIndex        =   23
         Tag             =   "XYZ"
         Top             =   4040
         Width           =   615
      End
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   15120
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "txtHoldGrayBackColor"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   15120
      TabIndex        =   13
      Text            =   "HoldBackColor"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   661
      Caption         =   "&NEW"
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
      Image           =   "frmRFP.frx":01FC
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   7
      Top             =   120
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   661
      Caption         =   "&EDIT"
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
      Image           =   "frmRFP.frx":0796
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   8
      Top             =   120
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   661
      Caption         =   "&PRINT"
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
      Image           =   "frmRFP.frx":0D30
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   3720
      TabIndex        =   9
      Top             =   120
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   661
      Caption         =   "SU&MMARY"
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
      Image           =   "frmRFP.frx":6522
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   5160
      TabIndex        =   10
      Top             =   120
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "&EXIT FORM"
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
      Image           =   "frmRFP.frx":6ABC
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   15120
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      Caption         =   "&CANCEL"
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
      Image           =   "frmRFP.frx":C2AE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   15120
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      Caption         =   "&SAVE"
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
      Image           =   "frmRFP.frx":C848
      cBack           =   -2147483633
   End
   Begin VB.Frame framHoldTop 
      Height          =   1095
      Left            =   120
      TabIndex        =   15
      Tag             =   "xyz"
      Top             =   600
      Width           =   14655
      Begin VB.ComboBox cboSoldTo 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   2
         Tag             =   "XYZ"
         Top             =   600
         Width           =   4575
      End
      Begin VB.ComboBox cboSupName 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   1
         Tag             =   "XYZ"
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox txtRFPNo 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   11280
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Tag             =   "XYZ"
         Top             =   240
         Width           =   1935
      End
      Begin MSMask.MaskEdBox txtRFPDate 
         Height          =   315
         Left            =   11280
         TabIndex        =   4
         Tag             =   "XYZ"
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12632256
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Caption         =   "MM/DD/YYYY"
         Height          =   255
         Left            =   13320
         TabIndex        =   21
         Tag             =   "XYZ"
         Top             =   675
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "RFP DATE:"
         Height          =   255
         Left            =   10080
         TabIndex        =   20
         Tag             =   "XYZ"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "RFP NUMBER:"
         Height          =   255
         Left            =   10080
         TabIndex        =   19
         Tag             =   "XYZ"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "SOLD TO:"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Tag             =   "XYZ"
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER NAME:"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Tag             =   "XYZ"
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSMask.MaskEdBox txtFlxDate 
      Height          =   315
      Left            =   4680
      TabIndex        =   51
      Tag             =   "XYZ"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   7320
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   15120
      m_MDI_Height    =   10380
   End
   Begin VB.Menu mnuPopUpFormat 
      Caption         =   "PopUpMenu"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuADD 
         Caption         =   "&ADD ROW"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "&DELETE ROW"
      End
   End
End
Attribute VB_Name = "frmRFP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmpRs As New ADODB.Recordset
Dim formAction As String
Dim didRecordEntry As Boolean

Dim AllArr3sBranch() As String
Dim arr3sBranch() As String

Dim AllArrSupplier() As String
Dim arrSupplier() As String

Dim AllArrCheckedBy() As String
Dim arrCheckedBy() As String

Dim SUPPLIERID As Integer

Dim SearchcboSup As Boolean
Dim SearchcboShipTo As Boolean
Dim SearchcboCheckedBy As Boolean
Dim SearchcboParticular As Boolean

Dim DtlsCol As Integer
Dim DtlsRow As Integer
Dim flexGridObject As String
Dim ans As String
Dim checkDate As String
Dim RFPID As Integer

Public Sub ShowRecord(ByVal pRFPNO As String)
On Error GoTo er
    Dim i As Integer
    
    Call modAPI.sp_load_rfp_info(pRFPNO, tmpRs)
    If tmpRs.State = 0 Then Exit Sub
    
    If tmpRs.RecordCount <= 0 Then
        tmpRs.Close
        Exit Sub
    End If
    
    tmpRs.MoveFirst
    
            cboSupName.Text = tmpRs!SUPPLIERNAME
            cboSoldTo.Text = tmpRs!soldto
            RFPID = tmpRs!RFPID
            txtRFPNO.Text = tmpRs!RFPNO
            txtRFPDate.Text = Format(tmpRs!RFPDATE, "mm/dd/yyyy")
            
            txtTotInvAmnt.Text = Format(tmpRs!TotalAmount, "#,##0.00")
            txtTotalDed.Text = Format(tmpRs!TotalDeduction, "#,##0.00")
            txtNetPay.Text = Format(tmpRs!TotalNetPay, "#,##0.00")

            txtPrepBy.Text = tmpRs!PREPBY
            cboChckedBy.Text = tmpRs!chekedby
            txtNotedBy.Text = tmpRs!NOTEDBY
            txtApprvdBy.Text = tmpRs!apprvdby
            
            txtCheckNo.Text = tmpRs!CHECKNO
            If IsNull(tmpRs!checkDate) = False Then txtChckNoDate.Text = Format(tmpRs!checkDate, "mm/dd/yyyy")
            
            
        While Not tmpRs.EOF
        

          If flxRFP.TextMatrix(1, 1) = "" Then
                flxRFP.TextMatrix(1, 1) = tmpRs!shippedto
                flxRFP.TextMatrix(1, 2) = tmpRs!SPCRRNO
                flxRFP.TextMatrix(1, 3) = tmpRs!DATERECEIVED
                flxRFP.TextMatrix(1, 4) = tmpRs!PACKINGNO
                flxRFP.TextMatrix(1, 5) = Format(tmpRs!INVOICEDATE, "mm/dd/yyyy")
                flxRFP.TextMatrix(1, 6) = tmpRs!InvoiceNO
                flxRFP.TextMatrix(1, 7) = tmpRs!PONO
                flxRFP.TextMatrix(1, 8) = tmpRs!WayBillno
                flxRFP.TextMatrix(1, 9) = Format(tmpRs!INVOICEAMNT, "#,##0.00")
  
            Else
                flxRFP.AddItem _
                vbTab & tmpRs!shippedto & _
                vbTab & tmpRs!SPCRRNO & _
                vbTab & tmpRs!DATERECEIVED & _
                vbTab & tmpRs!PACKINGNO & _
                vbTab & Format(tmpRs!INVOICEDATE, "mm/dd/yyyy") & _
                vbTab & tmpRs!InvoiceNO & _
                vbTab & tmpRs!PONO & _
                vbTab & tmpRs!WayBillno & _
                vbTab & Format(tmpRs!INVOICEAMNT, "#,##0.00")
 
            End If
            
        tmpRs.MoveNext
    Wend
    
    
    If tmpRs.State <> 0 Then tmpRs.Close


    DoEvents
    Call ShowDeduction(RFPID)
    Call EnabledSelectedButton
    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->ShowRecord()")
End Sub

Public Sub ShowDeduction(ByVal pRFPID As Integer)
On Error GoTo er
    Dim i As Integer
    
    Call modAPI.sp_load_rfp_deduction(pRFPID, tmpRs)
    If tmpRs.State = 0 Then Exit Sub
    
    If tmpRs.RecordCount <= 0 Then
        tmpRs.Close
        Exit Sub
    End If
    
    tmpRs.MoveFirst
    
            txtTotAdd.Text = Format(tmpRs!TotalAdd, "#,##0.00")
            txtTotLess.Text = Format(tmpRs!TotalLess, "#,##0.00")
            
        While Not tmpRs.EOF

            If flxDeduct.TextMatrix(1, 1) = "" Then
                flxDeduct.TextMatrix(1, 1) = tmpRs!Description
                flxDeduct.TextMatrix(1, 2) = tmpRs!Type
                flxDeduct.TextMatrix(1, 3) = tmpRs!Particulars
                flxDeduct.TextMatrix(1, 4) = tmpRs!deductionreb

  
            Else
                flxDeduct.AddItem _
                vbTab & tmpRs!Description & _
                vbTab & tmpRs!Type & _
                vbTab & tmpRs!Particulars & _
                vbTab & tmpRs!deductionreb

            End If
            
            
        tmpRs.MoveNext
    Wend
    
    
    If tmpRs.State <> 0 Then tmpRs.Close
    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->ShowDeduction()")
End Sub
Private Sub ComputeTotAddLess()
On Error GoTo er
    Dim i As Integer
    Dim TotalAdd As Double
    Dim TotalLess As Double
    
    For i = 1 To flxDeduct.Rows - 1
        If flxDeduct.TextMatrix(i, 2) = "ADD" Then
            TotalAdd = TotalAdd + Val(Replace(flxDeduct.TextMatrix(i, 4), ",", ""))
        ElseIf flxDeduct.TextMatrix(i, 2) = "LESS" Then
            TotalLess = TotalLess + Val(Replace(flxDeduct.TextMatrix(i, 4), ",", ""))
        End If
    Next
    
    txtTotAdd.Text = Format(TotalAdd, "#,##0.00")
    txtTotLess.Text = Format(TotalLess, "#,##0.00")
    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Erro->ComputeTotAddLess()")
End Sub

Private Sub ComputeTotalNetPay()
On Error GoTo er
    Dim i As Integer
    Dim j As Integer
    Dim TotalInvoice As Double
    Dim TotalDeduct As Double
    

    For i = 1 To flxRFP.Rows - 1
        TotalInvoice = TotalInvoice + Val(Replace(flxRFP.TextMatrix(i, 9), ",", ""))
    Next
    
    txtTotInvAmnt.Text = Format(TotalInvoice, "#,##0.00")
    
    For j = 1 To flxDeduct.Rows - 1
        'If flxDeduct.TextMatrix(j, 2) = "LESS" Then
            TotalDeduct = TotalDeduct + Val(Replace(flxDeduct.TextMatrix(j, 4), ",", ""))
        'End If
    Next
    
    txtTotalDed.Text = Format(TotalDeduct, "#,##0.00")
    
    
    txtNetPay.Text = Format(TotalInvoice - TotalDeduct, "#,##0.00")
    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->ComputeTotalNetPay()")
End Sub

Private Sub SetButtonVisibility(DataEntry As Boolean)
'cmdTrans Legend
'0 ->NEW
'1 ->EDIT
'2 ->PRINT
'3 ->SUMMRY
'4 ->EXITs
'5 ->SAVE
'6 ->CANCEL
    
    Dim i As Integer
    
    'hide all button
    For i = 0 To 6
        cmdTrans(i).Visible = False
    Next i
    
    If DataEntry Then
        cmdTrans(5).Visible = True
        cmdTrans(6).Visible = True
    Else
        For i = 0 To 4
            cmdTrans(i).Visible = True
        Next i
        
    End If
    
End Sub

Sub ClearEntryForm(Focus1stObject As Boolean)
    
    cboSupName.Text = ""
    cboSoldTo.Text = ""
    txtRFPNO.Text = ""
    
    txtRFPDate.Mask = ""
    txtRFPDate.Text = ""
    txtRFPDate.Mask = "##/##/####"
    
    txtTotInvAmnt.Text = ""
    txtTotLess.Text = ""
    txtTotAdd.Text = ""
    txtTotalDed.Text = ""
    txtNetPay.Text = ""
    cboChckedBy.Text = ""
    txtPrepBy.Text = EmployeesName
    txtNotedBy.Text = GetSettings("RFP Noted By")
    txtApprvdBy.Text = GetSettings("RFP Approved By")
    
    txtCheckNo.Text = ""
    
    txtChckNoDate.Mask = ""
    txtChckNoDate.Text = ""
    txtChckNoDate.Mask = "##/##/####"
    
    flxRFP.Rows = 1
    flxRFP.Rows = 2
    
    flxDeduct.Rows = 1
    flxDeduct.Rows = 2

    If Focus1stObject = True Then cmdTrans(0).SetFocus
                              
End Sub
Sub LockedEntryObject(DataEntry As Boolean)

    DataEntry = Not DataEntry
    
        cboSupName.Locked = DataEntry
        cboSoldTo.Locked = DataEntry
        txtRFPNO.Locked = DataEntry
        txtCheckNo.Locked = DataEntry
        cboChckedBy.Locked = DataEntry
                                                                                                                                                                                                                           
    If cboSupName.Locked Then
    
        cboSupName.BackColor = txtHoldGrayBackColor.BackColor
        cboSoldTo.BackColor = txtHoldGrayBackColor.BackColor
        txtRFPNO.BackColor = txtHoldGrayBackColor.BackColor
        cboChckedBy.BackColor = txtHoldGrayBackColor.BackColor
        txtRFPDate.BackColor = txtHoldGrayBackColor.BackColor
        txtCheckNo.BackColor = txtHoldGrayBackColor.BackColor
        txtChckNoDate.BackColor = txtHoldGrayBackColor.BackColor
        
        txtRFPDate.Enabled = False
        txtChckNoDate.Enabled = False
        
    Else
        
        cboSupName.BackColor = txtHoldWhiteBackColor.BackColor
        cboSoldTo.BackColor = txtHoldWhiteBackColor.BackColor
        txtRFPNO.BackColor = txtHoldWhiteBackColor.BackColor
        cboChckedBy.BackColor = txtHoldWhiteBackColor.BackColor
        txtRFPDate.BackColor = txtHoldWhiteBackColor.BackColor
        txtCheckNo.BackColor = txtHoldWhiteBackColor.BackColor
        txtChckNoDate.BackColor = txtHoldWhiteBackColor.BackColor
        
        txtRFPDate.Enabled = True
        txtChckNoDate.Enabled = True
        
    End If

End Sub

Private Sub cboChckedBy_Change()
On Error GoTo er
    Dim i As Integer, j As Integer
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
    If didRecordEntry = False Then Exit Sub
    
    If SearchcboCheckedBy Then
        SearchcboCheckedBy = False
        Exit Sub
    End If
    
    With cboChckedBy
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters

        If i <> -1 And i <= .ListCount Then
            strTotal = .List(i)
    
            DoEvents
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
               SearchcboCheckedBy = True
               .SelText = Right$(strTotal, j)
               .SelStart = Len(strPartial)
               .SelLength = j
            End If

        End If

    End With
        
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error ->cboChckedBy_Change()")
End Sub

Private Sub cboChckedBy_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Or KeyCode = 46 Then
        SearchcboCheckedBy = True
    End If
End Sub


Private Sub cboChckedBy_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboChckedBy, KeyAscii, ValidChars)
    
    If KeyAscii = 13 Then txtCheckNo.SetFocus
End Sub

Private Sub cboFlxParticular_Change()
On Error GoTo er
    Dim i As Integer, j As Integer
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
    If didRecordEntry = False Then Exit Sub
    
    If SearchcboParticular Then
        SearchcboParticular = False
        Exit Sub
    End If

    With cboFlxParticular
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters

        If i <> -1 And i <= .ListCount Then
            strTotal = .List(i)
    
            DoEvents
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
               SearchcboSup = True
               .SelText = Right$(strTotal, j)
               .SelStart = Len(strPartial)
               .SelLength = j
            End If

        End If

    End With
        
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error ->cboFlxParticular_Change()")
End Sub

Private Sub cboFlxParticular_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Or KeyCode = 46 Then
        SearchcboParticular = True
    End If
End Sub


Private Sub cboFlxType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxDeduct.TextMatrix(flxDeduct.row, 2) = cboFlxType.Text
        cboFlxType.Visible = False
        flxDeduct.SetFocus
    End If
End Sub


Private Sub cboFlxType_LostFocus()
    Call cboFlxType_KeyPress(13)
End Sub


Private Sub cboFlxParticular_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        Dim LessValue As Double
        
        If flxDeduct.TextMatrix(flxDeduct.row, 2) = "LESS" Then
        
            LessValue = (Val(Replace(txtTotInvAmnt.Text, ",", "")) / 1.12) * 0.01
            flxDeduct.TextMatrix(flxDeduct.row, 4) = Format(LessValue, "#,##0.00")
    
        End If
        
        flxDeduct.TextMatrix(flxDeduct.row, 3) = cboFlxParticular.Text
        cboFlxParticular.Visible = False
        flxDeduct.SetFocus
        Call ComputeTotalNetPay
        Call ComputeTotAddLess
        
    End If
End Sub


Private Sub cboFlxParticular_LostFocus()
    Call cboFlxParticular_KeyPress(13)
End Sub


Private Sub cboParticular_Change()

End Sub

Private Sub cboSoldTo_Change()
On Error GoTo er
    Dim i As Integer, j As Integer
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
    'If didRecordEntry = False Then Exit Sub
    
    If SearchcboShipTo Then
        SearchcboShipTo = False
        Exit Sub
    End If

    
    'branchID = 0
    With cboSoldTo
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters

        If i <> -1 And i <= .ListCount Then
            strTotal = .List(i)
            
            'branchID = arr3sBranch(i, 2)
    
            DoEvents
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
               SearchcboShipTo = True
               .SelText = Right$(strTotal, j)
               .SelStart = Len(strPartial)
               .SelLength = j
            End If
        Else
            'branchID = 0
        End If

    End With
        
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error ->cboSoldTo_Change()")
End Sub

Private Sub cboSoldTo_Click()
'If didRecordEntry = False Then Exit Sub
'
'    If cboSoldTo.ListIndex <> -1 Then
'        branchID = arrSupplier(cboSoldTo.ListIndex, 0)
'    End If
End Sub


Private Sub cboSoldTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Or KeyCode = 46 Then
        SearchcboShipTo = True
    End If
End Sub

Private Sub cboSoldTo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboSoldTo, KeyAscii, ValidChars)
    
    If KeyAscii = 13 Then
'        If cboSoldTo.ListIndex <> -1 Then
'            branchID = arrSupplier(cboSoldTo.ListIndex, 0)
'        End If
        txtRFPNO.SetFocus
    End If
End Sub

Private Sub cboSupName_Change()
On Error GoTo er
    Dim i As Integer, j As Integer
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
    'If didRecordEntry = False Then Exit Sub
    
    If SearchcboSup Then
        SearchcboSup = False
        Exit Sub
    End If

    
    SUPPLIERID = 0
    With cboSupName
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters

        If i <> -1 And i <= .ListCount Then
            strTotal = .List(i)
            
            SUPPLIERID = arrSupplier(i, 0)
    
            DoEvents
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
               SearchcboSup = True
               .SelText = Right$(strTotal, j)
               .SelStart = Len(strPartial)
               .SelLength = j
            End If
            
        Else
            SUPPLIERID = 0
        End If

    End With
        
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error ->cboSupName_Change()")
End Sub

Private Sub cboSupName_Click()
If didRecordEntry = False Then Exit Sub

    If cboSupName.ListIndex <> -1 Then
        SUPPLIERID = arrSupplier(cboSupName.ListIndex, 0)
    End If
End Sub


Private Sub cboSupName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Or KeyCode = 46 Then
        SearchcboSup = True
    End If
End Sub

Private Sub cboSupName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboSupName, KeyAscii, ValidChars)
    
    If KeyAscii = 13 Then
        If cboSupName.ListIndex <> -1 Then
            SUPPLIERID = arrSupplier(cboSupName.ListIndex, 0)
        End If
        cboSoldTo.SetFocus
    End If
End Sub


Private Sub CreatTempTable()
On Error GoTo er
Dim Qry As String

        Qry = "Drop Temporary Table IF  Exists sp_acctpayinvcng.tmp_rfpdtls;"
        Call LYRcon.Execute(Qry)
        
        DoEvents
        Qry = "Create Temporary Table IF Not Exists sp_acctpayinvcng.tmp_rfpdtls( " & _
        "   shippedto       varchar(60), " & _
        "   spcrrno         int(10),   " & _
        "   datereceived    date,  " & _
        "   packingno       varchar(60),   " & _
        "   invoicedate     date,  " & _
        "   invoiceno       varchar(45),  " & _
        "   pono            varchar(45),  " & _
        "   waybillno       varchar(45),  " & _
        "   invoiceamnt     Decimal(16, 2));"
        
        'Text1.Text = Qry
        Call LYRcon.Execute(Qry)
        

        Qry = "Drop Temporary Table IF  Exists sp_acctpayinvcng.tmp_rfpdeduct;"
        Call LYRcon.Execute(Qry)
        
        DoEvents
        Qry = "Create Temporary Table IF Not Exists sp_acctpayinvcng.tmp_rfpdeduct( " & _
        "   description     varchar(100), " & _
        "   type            varchar(10),   " & _
        "   datereceived    date,  " & _
        "   particulars     varchar(100),   " & _
        "   deductionreb    decimal(16, 2),  " & _
        "   totalless       decimal(16, 2),  " & _
        "   totaladd        decimal(16, 2));"
        
        'Text1.Text = Qry
        Call LYRcon.Execute(Qry)
'
    
    Call InsertTempTable
                
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CreatTempTable()")
End Sub
Private Sub InsertTempTable()
On Error GoTo er
Dim Qry As String
    
        Qry = "Insert Into sp_acctpayinvcng.tmp_rfpdtls(shippedto, spcrrno, datereceived, packingno, invoicedate, invoiceno, pono, waybillno, invoiceamnt) Values "

        DtlsRow = 0
        For DtlsRow = 1 To flxRFP.Rows - 1
            
                If DtlsRow > 1 Then Qry = Qry & ","
    
                Qry = Qry & "('" & flxRFP.TextMatrix(DtlsRow, 1) & "','" & flxRFP.TextMatrix(DtlsRow, 2) & "','" & Format(flxRFP.TextMatrix(DtlsRow, 3), "YYYY-mm-dd") & _
                            "','" & Trim(flxRFP.TextMatrix(DtlsRow, 4)) & "','" & Format(flxRFP.TextMatrix(DtlsRow, 5), "YYYY-mm-dd") & "','" & Trim(flxRFP.TextMatrix(DtlsRow, 6)) & _
                            "','" & Trim(flxRFP.TextMatrix(DtlsRow, 7)) & "','" & Trim(flxRFP.TextMatrix(DtlsRow, 8)) & "','" & Val(Replace(flxRFP.TextMatrix(DtlsRow, 9), ",", "")) & "')"
         Next
    
        'InputBox Qry, Qry, Qry
        Call LYRcon.Execute(Qry)
        
        Qry = "Insert Into sp_acctpayinvcng.tmp_rfpdeduct(description, `type`, particulars, deductionreb, totalless, totaladd) Values "

        DtlsRow = 0
        For DtlsRow = 1 To flxDeduct.Rows - 1
            
                If DtlsRow > 1 Then Qry = Qry & ","
    
                Qry = Qry & "('" & flxDeduct.TextMatrix(DtlsRow, 1) & "','" & flxDeduct.TextMatrix(DtlsRow, 2) & _
                            "','" & flxDeduct.TextMatrix(DtlsRow, 3) & "','" & Val(Replace(Trim(flxDeduct.TextMatrix(DtlsRow, 4)), ",", "")) & _
                            "', '" & Val(Replace(txtTotLess.Text, ",", "")) & "', '" & Val(Replace(txtTotAdd.Text, ",", "")) & "')"
         Next
    
        'InputBox Qry, Qry, Qry
        Call LYRcon.Execute(Qry)
        
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->InsertTempTable()")
End Sub
Private Sub cmdTrans_Click(Index As Integer)
On Error GoTo er
'cmdTrans Legend
'0 ->NEW
'1 ->EDIT
'2 ->PRINT
'3 ->SUMMRY
'4 ->EXIT
'5 ->SAVE
'6 ->CANCEL

    Dim ans As VbMsgBoxResult
    Dim level As String
    Dim message As String
    

    If Index = 0 Then 'NEW
        formAction = "NEW"
        didRecordEntry = True
        Call SetButtonVisibility(True)
        Call LockedEntryObject(True)
        Call ClearEntryForm(False)
        cboSupName.SetFocus
        
    ElseIf Index = 1 Then 'EDIT
        formAction = "EDIT"
        didRecordEntry = True
        Call SetButtonVisibility(True)
        Call LockedEntryObject(True)
        txtPrepBy.Text = EmployeesName
        
        cboSupName.SetFocus
        
    ElseIf Index = 2 Then 'PRINT
       Dim Qry1 As String
       Dim Qry2 As String
       Dim tmpRs1 As New ADODB.Recordset
       Dim tmpRs2 As New ADODB.Recordset
       Screen.MousePointer = vbHourglass
       Dim rpt As New rptRFP

        Qry1 = "Call sp_acctpayinvcng.sp_load_rfp_info('" & txtRFPNO.Text & "');"
        DoEvents
        'InputBox Qry, Qry, Qry
        If tmpRs1.State <> 0 Then tmpRs1.Close
        Set tmpRs1 = LYRcon.Execute(Qry1)

        Qry2 = "SELECT * FROM mktg3s.rfpdeduct INNER JOIN mktg3s.rfpmain ON rfpdeduct.rfpid = rfpmain.rfpid where rfpdeduct.rfpid = (" & RFPID & ");"
        DoEvents
        'InputBox Qry, Qry, Qry
        If tmpRs2.State <> 0 Then tmpRs2.Close
        Set tmpRs2 = LYRcon.Execute(Qry2)

        DoEvents
        'Main report
        With rpt
            .DiscardSavedData
             DoEvents
            .Database.SetDataSource tmpRs1
            .ReadRecords
        End With
        
        'Subreport Deductions
        With rpt.Subreport1.OpenSubreport
            .DiscardSavedData
             DoEvents
            .Database.SetDataSource tmpRs2
            .ReadRecords
        End With

       frmPrint.CRViewer1.ReportSource = rpt
       frmPrint.CRViewer1.ViewReport
       frmPrint.CRViewer1.Zoom 150
       frmPrint.CRViewer1.EnableExportButton = False
       Screen.MousePointer = vbDefault

       DoEvents
       frmPrint.Caption = "PRINT" & " " & "-" & " " & "[" & "REQUEST FOR PAYMENTS" & "]"
       frmPrint.WindowState = vbMaximized
       frmPrint.MySource = "RFP"
       frmPrint.Show vbModal
        
    ElseIf Index = 3 Then 'SUMMARY
        frm_OnProcess2.OwnerForm = "RFP"
        FrmSummary.mySrc = "RFP"
        FrmSummary.Show
        formAction = "EXIT"
        Unload Me

    ElseIf Index = 4 Then 'EXIT
        formAction = "EXIT"
        
        ans = MsgBox("ARE YOU SURE YOU WANT TO EXIT?", vbQuestion + vbYesNo, Me.Caption)
        If ans = vbNo Then
            formAction = ""
            Exit Sub
        End If
        
        Unload Me
        
    ElseIf Index = 5 Then  'SAVE
        If ValidateEntry() <> "" Then
            MsgBox ValidateEntry, vbExclamation, Me.Caption
            Exit Sub
        End If


        Screen.MousePointer = vbHourglass

        Call CreatTempTable

        DoEvents
        If formAction = "EDIT" Then

            ans = MsgBox("DO YOU WANT TO SAVE YOUR CHANGES?", vbQuestion + vbYesNo, Me.Caption)
            If ans = vbNo Then Exit Sub

            Call modAPI.sp_rfp_update(RFPID, Trim(txtRFPNO.Text), txtRFPDate.Text, SUPPLIERID, cboSoldTo.Text, cboChckedBy.Text _
                                     , txtNotedBy.Text, txtApprvdBy.Text, txtCheckNo.Text, txtChckNoDate.Text, txtTotInvAmnt.Text _
                                     , txtTotalDed.Text, txtNetPay.Text, level, message)

        ElseIf formAction = "NEW" Then

            ans = MsgBox("PLEASE MAKE SURE THAT THE DATA IS CORRECT!" & _
                vbCrLf & _
                vbCrLf & "DO YOU WISH TO CONTINUE?", vbYesNo + vbQuestion, Me.Caption)
            If ans = vbNo Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            
            Call modAPI.sp_rfp_create(Trim(txtRFPNO.Text), txtRFPDate.Text, SUPPLIERID, cboSoldTo.Text, cboChckedBy.Text _
                                     , txtNotedBy.Text, txtApprvdBy.Text, txtCheckNo.Text, txtChckNoDate.Text, txtTotInvAmnt.Text _
                                     , txtTotalDed.Text, txtNetPay.Text, level, message)

       End If


        Screen.MousePointer = vbDefault

        If Trim(level) = "" Then
            Exit Sub
        ElseIf level = "Granted" Or level = "Warning" Then
            If level = "Granted" Then
                MsgBox message, vbInformation, "SAVE."
            Else
                MsgBox message, vbExclamation, "SAVE!"
                Exit Sub
            End If
        Else
            MsgBox message, vbCritical, "Error!"
            Exit Sub
        End If


        Call LockedEntryObject(False)
        Call SetButtonVisibility(False)
        Call EnabledSelectedButton

        didRecordEntry = False

        formAction = ""
        
    ElseIf Index = 6 Then 'CANCEL Button
        ans = MsgBox("Cancel the currrent operation?", vbYesNo + vbQuestion, Me.Caption)
        If ans = vbNo Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        Call LockedEntryObject(False)
        Call SetButtonVisibility(False)
        Call ClearEntryForm(True)
        Call EnabledSelectedButton
        
        didRecordEntry = False
        formAction = ""

    End If
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdTrans_Click()")
End Sub


Private Sub flxDeduct_Click()
    flexGridObject = "flxDeduct"
End Sub

Private Sub flxDeduct_DblClick()
On Error GoTo er

    If didRecordEntry = False Then Exit Sub
   
    DtlsCol = flxDeduct.Col
    If DtlsCol = 1 Then
       With txtFlxValueDeduct
        .Top = framHoldButtton.Top + flxDeduct.CellTop + flxDeduct.Top
        .Left = framHoldButtton.Left + flxDeduct.CellLeft + flxDeduct.Left
        .Width = flxDeduct.CellWidth
        .Height = flxDeduct.CellHeight
        .Text = flxDeduct.TextMatrix(flxDeduct.row, DtlsCol)
        .Visible = True
        .BorderStyle = 1
        .SelStart = 0
        .SelLength = Len(.Text)
        .ZOrder
        .SetFocus
      End With
    ElseIf DtlsCol = 2 Then
       With cboFlxType
        .Top = framHoldButtton.Top + flxDeduct.CellTop + flxDeduct.Top
        .Left = framHoldButtton.Left + flxDeduct.CellLeft + flxDeduct.Left
        .Width = flxDeduct.CellWidth
        .Visible = True
        .ZOrder
        .SetFocus
      End With
    ElseIf DtlsCol = 3 Then
       With cboFlxParticular
        .Top = framHoldButtton.Top + flxDeduct.CellTop + flxDeduct.Top
        .Left = framHoldButtton.Left + flxDeduct.CellLeft + flxDeduct.Left
        .Width = flxDeduct.CellWidth
        .Text = flxDeduct.TextMatrix(flxDeduct.row, DtlsCol)
        .Visible = True
        .ZOrder
        .SetFocus
      End With
    ElseIf DtlsCol = 4 Then
       With txtFlxDeductReb
        .Top = framHoldButtton.Top + flxDeduct.CellTop + flxDeduct.Top
        .Left = framHoldButtton.Left + flxDeduct.CellLeft + flxDeduct.Left
        .Width = flxDeduct.CellWidth
        .Height = flxDeduct.CellHeight
        .Text = Val(Replace(Trim(flxDeduct.TextMatrix(flxDeduct.row, DtlsCol)), ",", ""))
        .Alignment = vbRightJustify
        .Visible = True
        .BorderStyle = 1
        .SelStart = 0
        .SelLength = Len(.Text)
        .ZOrder
        .SetFocus
      End With
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error ->flxDeduct_DblClick()")
End Sub

Private Sub flxDeduct_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxDeduct_DblClick
End Sub

Private Sub flxDeduct_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If didRecordEntry = False Then Exit Sub
    If Button = vbRightButton Then
        PopupMenu mnuPopUpFormat, vbPopupMenuRightButton
        flexGridObject = "flxDeduct"
    End If
End Sub


Private Sub flxRFP_Click()
    flexGridObject = "flxRFP"
End Sub

Private Sub flxRFP_DblClick()
On Error GoTo er

    If didRecordEntry = False Then Exit Sub
   
    DtlsCol = flxRFP.Col
    DtlsRow = flxRFP.row
    If DtlsCol = 2 Or DtlsCol = 6 Or DtlsCol = 7 Or DtlsCol = 8 Then
       With txtFlxValueRFP
        .Top = framHoldMid.Top + flxRFP.CellTop + flxRFP.Top
        .Left = framHoldMid.Left + flxRFP.CellLeft + flxRFP.Left
        .Width = flxRFP.CellWidth
        .Height = flxRFP.CellHeight
        .Text = flxRFP.TextMatrix(flxRFP.row, DtlsCol)
        .Visible = True
        .BorderStyle = 1
        .SelStart = 0
        .SelLength = Len(.Text)
        .ZOrder
        .SetFocus
      End With
    ElseIf DtlsCol = 5 Then
       With txtFlxDate
        .Top = framHoldMid.Top + flxRFP.CellTop + flxRFP.Top
        .Left = framHoldMid.Left + flxRFP.CellLeft + flxRFP.Left
        .Width = flxRFP.CellWidth
        .Height = flxRFP.CellHeight
        If flxRFP.TextMatrix(flxRFP.row, 5) <> "" Then .Text = flxRFP.TextMatrix(flxRFP.row, DtlsCol)
        .Visible = True
        .BorderStyle = 1
        .SelStart = 0
        .SelLength = Len(.Text)
        .ZOrder
        .SetFocus
      End With
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error ->flxRFP_DblClick()")
End Sub


Private Sub flxRFP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxRFP_DblClick
End Sub


Private Sub flxRFP_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If didRecordEntry = False Then Exit Sub
    If Button = vbRightButton Then
        PopupMenu mnuPopUpFormat, vbPopupMenuRightButton
        flexGridObject = "flxRFP"
    End If
End Sub

Private Sub Form_Load()
On Error GoTo er

    DoEvents
    Call modAPI.sp_checkedBy_list(tmpRs)
    If tmpRs.State = 1 Then
        If tmpRs.RecordCount > 0 Then
            Call modObjTools.fillArray(AllArrCheckedBy, tmpRs)
        End If
    End If
    If tmpRs.State <> 0 Then tmpRs.Close

    DoEvents
    Call modObjTools.addToListObjectUsingArrayWithCols(cboChckedBy, arrCheckedBy, AllArrCheckedBy, 0, 0)
    
    DoEvents
    Call modAPI.sp_supplier_list(tmpRs)
    If tmpRs.State = 1 Then
        If tmpRs.RecordCount > 0 Then
            Call modObjTools.fillArray(AllArrSupplier, tmpRs)
        End If
    End If
    If tmpRs.State <> 0 Then tmpRs.Close
    
    DoEvents
    Call modObjTools.addToListObjectUsingArrayWithCols(cboSupName, arrSupplier, AllArrSupplier, 0, 1)
    
'    DoEvents
'    Call modAPI.sp_3s_branch_list(tmpRs)
'    If tmpRs.State = 1 Then
'        If tmpRs.RecordCount > 0 Then
'            Call modObjTools.fillArray(AllArr3sBranch, tmpRs)
'        End If
'    End If
'    If tmpRs.State <> 0 Then tmpRs.Close
'
'    DoEvents
'    Call modObjTools.addToListObjectUsingArrayWithCols(cboSoldTo, arr3sBranch, AllArr3sBranch, 2, 1)
    
    cboSoldTo.AddItem "MOTORHUB CORPORATION"
    cboSoldTo.AddItem "BINOS MARKETING CORPORATION"
    cboSoldTo.AddItem "BINOS MARKETING"
    cboSoldTo.AddItem "LYR MARKETING CORPORATION"
    
    cboFlxParticular.AddItem "DEBIT MEMO"
    cboFlxParticular.AddItem "CREDIT MEMO"
    cboFlxParticular.AddItem "EXPANDED WITH HOLDING TAX 1%"
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->Form_Load()")
End Sub

Private Sub Form_Resize()
    Call RXResizer1_ResizeComplete
End Sub



Public Sub GetSPCRRRecord(ByVal pSPCRRNO As Integer)
On Error GoTo er
    
    Call modAPI.sp_get_spcrr_info(pSPCRRNO, tmpRs)
    If tmpRs.State = 0 Then Exit Sub
    
    If tmpRs.RecordCount <= 0 Then
        MsgBox "NO RECORD FOUND!", vbInformation, Me.Caption
        Screen.MousePointer = vbDefault
        tmpRs.Close
        Exit Sub
    ElseIf tmpRs!INVSTATUS = "PENDING" Then
        MsgBox "SPCRR number status was still pending!", vbExclamation, Me.Caption
        Screen.MousePointer = vbDefault
        tmpRs.Close
        Exit Sub
    End If

    tmpRs.MoveFirst
    
        While Not tmpRs.EOF
            'BRANCHNAME, SUPINVID, DATERCVD, PACKINGNO, PONO, WAYBILLNO, INVOICETTLPAY
            flxRFP.TextMatrix(flxRFP.row, 1) = tmpRs!branchName
            flxRFP.TextMatrix(flxRFP.row, 2) = tmpRs!SUPINVID
            flxRFP.TextMatrix(flxRFP.row, 3) = tmpRs!DATERCVD
            flxRFP.TextMatrix(flxRFP.row, 4) = tmpRs!PACKINGNO
            flxRFP.TextMatrix(flxRFP.row, 7) = tmpRs!PONO
            flxRFP.TextMatrix(flxRFP.row, 8) = tmpRs!WayBillno
            flxRFP.TextMatrix(flxRFP.row, 9) = Format(tmpRs!INVOICETTLPAY, "#,##0.00")
    tmpRs.MoveNext
    Wend
    

    If tmpRs.State <> 0 Then tmpRs.Close
    
    txtFlxValueRFP.Text = ""
    txtFlxValueRFP.Visible = False
    Screen.MousePointer = vbDefault
    
Exit Sub
er:
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->GetSPCRRRecord()")
End Sub

Private Sub mnuAdd_Click()
On Error GoTo er
    
    If didRecordEntry = False Then Exit Sub
    
    If flexGridObject = "flxRFP" Then
        flxRFP.AddItem "" & _
        vbTab & "" & _
        vbTab & "" & _
        vbTab & "" & _
        vbTab & "" & _
        vbTab & "" & _
        vbTab & "" & _
        vbTab & "" & _
        vbTab & "" & _
        vbTab & ""
        
        flxRFP.Col = 2
        flxRFP.SetFocus
    ElseIf flexGridObject = "flxDeduct" Then
        flxDeduct.AddItem "" & _
        vbTab & "" & _
        vbTab & "" & _
        vbTab & "" & _
        vbTab & ""
        
        flxDeduct.Col = 1
        flxDeduct.SetFocus
    End If

Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->mnuADD_Click()")
End Sub

Private Sub mnuDel_Click()
On Error GoTo er

    If didRecordEntry = False Then Exit Sub
    
    ans = MsgBox("Delete row?", vbYesNo + vbQuestion, Me.Caption)
    If ans = vbNo Then Exit Sub
    
    
    If flexGridObject = "flxRFP" Then
        If flxRFP.Rows = 2 Then
            
            flxRFP.Rows = 1
            flxRFP.Rows = 2
           
        Else
            
            flxRFP.RemoveItem (flxRFP.RowSel)
            
        End If
        
    ElseIf flexGridObject = "flxDeduct" Then
        If flxDeduct.Rows = 2 Then
            
            flxDeduct.Rows = 1
            flxDeduct.Rows = 2
           
        Else
            
            flxDeduct.RemoveItem (flxDeduct.RowSel)
            
        End If
    End If
    
    Call ComputeTotalNetPay
    Call ComputeTotAddLess
    
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error ->mnuDel_Click()"
End Sub


Private Sub RXResizer1_ResizeComplete()
On Error Resume Next
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->PRINT
'3 ->SUMMRY
'4 ->EXIT
'5 ->SAVE
'6 ->CANCEL

    'Align Save Button
    cmdTrans(5).Left = cmdTrans(0).Left
    cmdTrans(5).Top = cmdTrans(0).Top

    'Align Cancel Button
    cmdTrans(6).Left = Me.Width - 1600
    cmdTrans(6).Top = cmdTrans(0).Top
    
    
    With flxRFP
        .Clear
        .Rows = 2
        .Cols = 10
        .WordWrap = True
        .AllowUserResizing = flexResizeBoth
        
        '               0       1        2           3              4            5             6            7           8                9
        .FormatString = "  |<SHIPPED TO |^SPCRR NO. |^DATE RECEIVED |^PACKING NO. |^INVOICE DATE |^INVOICE NO. |^P.O. NO. |^WAYBILL NO. |>INVOICE AMOUNT "
        
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.21
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.1
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.08
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.1
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.08
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.1
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.1
        .ColWidth(8) = (.Width - .ColWidth(0)) * 0.1
        .ColWidth(9) = (.Width - .ColWidth(0)) * 0.11
    End With
    
    With flxDeduct
        .Clear
        .Rows = 2
        .Cols = 5
        '               0       1           2           3              4
        .FormatString = "  |<DESCRIPTION |^TYPE |<PARTICULARS |>DEDUCTION/REBATE "
        
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.26
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.1
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.47
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.15
    End With
End Sub


Private Sub tmrGetRecord_Timer()
On Error GoTo er
    If tmrGetRecord.Enabled = False Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    tmrGetRecord.Enabled = False
    
    DoEvents
    Call ShowRecord(txtRFPNO.Text)
    
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    
    Screen.MousePointer = Default
    
Exit Sub
er:
    tmrGetRecord.Enabled = False
    Screen.MousePointer = Default
End Sub

Private Sub txtChckNoDate_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtChckNoDate, KeyAscii, NumbersOnly)
    
    If KeyAscii = 13 Then
        If didRecordEntry Then cmdTrans(5).SetFocus
    End If
End Sub


Private Sub txtCheckNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtChckNoDate.SetFocus
End Sub


Private Sub txtFlxDate_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFlxDate, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then
        If IsDate(txtFlxDate.Text) Then
            flxRFP.TextMatrix(DtlsRow, 5) = txtFlxDate.Text
        Else
            flxRFP.TextMatrix(DtlsRow, 5) = ""
        End If
        
        If flxRFP.row <> flxRFP.Rows - 1 Then
            flxRFP.row = DtlsRow + 1
        End If
        flxRFP.SetFocus
        txtFlxDate.Visible = False

    End If
End Sub


Private Sub txtFlxDate_LostFocus()
    Call txtFlxDate_KeyPress(13)
End Sub


Private Sub txtFlxDeductReb_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFlxDeductReb, KeyAscii, NumbersWithDecimalPoint)
    If KeyAscii = 13 Then

        flxDeduct.TextMatrix(flxDeduct.row, 4) = Format(txtFlxDeductReb.Text, "#,##0.00")
        flxDeduct.SetFocus
        txtFlxDeductReb.Visible = False
        Call ComputeTotalNetPay
        Call ComputeTotAddLess
    End If
End Sub


Private Sub txtFlxDeductReb_LostFocus()
    Call txtFlxDeductReb_KeyPress(13)
End Sub


Private Sub txtFlxValueDeduct_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxDeduct.TextMatrix(flxDeduct.row, 1) = txtFlxValueDeduct.Text
        txtFlxValueDeduct.Visible = False
        flxDeduct.SetFocus
    End If
End Sub


Private Sub txtFlxValueDeduct_LostFocus()
    Call txtFlxValueDeduct_KeyPress(13)
End Sub


Private Sub txtFlxValueRFP_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(txtFlxValueRFP, KeyAscii, ValidCodeChars)
    If KeyAscii = 13 Then
        
        If DtlsCol = 2 And txtFlxValueRFP.Text <> "" Then
            Call GetSPCRRRecord(Trim(txtFlxValueRFP.Text))
            Call ComputeTotalNetPay
        ElseIf DtlsCol = 6 Then
            flxRFP.TextMatrix(DtlsRow, 6) = txtFlxValueRFP.Text
        ElseIf DtlsCol = 7 Then
            flxRFP.TextMatrix(DtlsRow, 7) = txtFlxValueRFP.Text
        ElseIf DtlsCol = 8 Then
            flxRFP.TextMatrix(DtlsRow, 8) = txtFlxValueRFP.Text
        End If
        
        If flxRFP.row <> flxRFP.Rows - 1 Then
            flxRFP.row = DtlsRow + 1
        End If
        flxRFP.SetFocus
        txtFlxValueRFP.Visible = False
        
    End If
End Sub




Private Sub txtFlxValueRFP_LostFocus()
    Call txtFlxValueRFP_KeyPress(13)
End Sub




Private Sub txtRFPDate_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtRFPDate, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then
        flxRFP.Col = 2
        flxRFP.row = 1
        flxRFP.SetFocus
    End If
End Sub


Private Sub txtRFPNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtRFPNO, KeyAscii, ValidCodeChars)
    If KeyAscii = 13 Then txtRFPDate.SetFocus
End Sub

Public Sub EnabledSelectedButton()
'cmdTrans Legend
'0 ->NEW
'1 ->EDIT
'2 ->PRINT
'3 ->SUMMRY
'4 ->EXIT
'5 ->SAVE
'6 ->CANCEL
    
    If cboSupName.Text <> "" Then
        If txtCheckNo.Text = "" Then
            cmdTrans(1).Enabled = True 'Edit Button
        Else
            cmdTrans(1).Enabled = False 'Edit Button
        End If
        cmdTrans(2).Enabled = True 'PRINT Button
    Else
        cmdTrans(1).Enabled = False 'Edit Button
        cmdTrans(2).Enabled = False 'PRINT Button
    End If
        
 
End Sub

Function ValidateEntry() As String
    Dim vMsg As String, Qry As String
    Dim i As Integer, Cntr As Integer
    
    vMsg = ""
    
    If cboSupName.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field supplier name."
    If SUPPLIERID = 0 Then vMsg = vMsg & vbCrLf & "Supplier name doesn't exist!"
    
    If cboSoldTo.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field sold to."

    If Not IsDate(txtRFPDate.Text) Then
        vMsg = vMsg & vbCrLf & "Invalid RFP date."
    End If
    

    If Not IsDate(txtChckNoDate.Text) Then
        'vMsg = vMsg & vbCrLf & "Invalid due date."
        txtChckNoDate.Text = "00/00/0000"
    End If
    

    If flxRFP.TextMatrix(1, 1) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field shipped to."
    If flxRFP.TextMatrix(1, 2) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field spcrr no."
    If flxRFP.TextMatrix(1, 3) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field date received."
    If flxRFP.TextMatrix(1, 4) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field packning no."
    If flxRFP.TextMatrix(1, 5) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field invoice date."
    If flxRFP.TextMatrix(1, 6) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field invoice no."
    If flxRFP.TextMatrix(1, 7) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field P.O. no."
    If flxRFP.TextMatrix(1, 8) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field waybill no."
    If flxRFP.TextMatrix(1, 9) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field invoice amount."

    
    If txtPrepBy.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field prepared by name."
    If cboChckedBy.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field checked by name."
    If txtNotedBy.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field noted by name."
    If txtApprvdBy.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field approved by name."
    
    ValidateEntry = vMsg
    
End Function

