VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmNewPricing 
   Caption         =   "PRICING DISCOUNT"
   ClientHeight    =   9465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9465
   ScaleWidth      =   16305
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8160
      Top             =   0
   End
   Begin VB.Timer tmAutoLoad 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8760
      Top             =   0
   End
   Begin VB.Frame SPpricing 
      Height          =   8775
      Left            =   120
      TabIndex        =   0
      Tag             =   "XYF"
      Top             =   600
      Width           =   16095
      Begin VB.ComboBox cboStat 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         ItemData        =   "frmNewPricing.frx":0000
         Left            =   12360
         List            =   "frmNewPricing.frx":000A
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Tag             =   "XYZ"
         Top             =   7320
         Width           =   3255
      End
      Begin VB.TextBox txtAppBy 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   12360
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Tag             =   "XYZ"
         Text            =   "LORENZO C. RUBINOS JR."
         Top             =   8040
         Width           =   3255
      End
      Begin VB.TextBox txtRemarks 
         BackColor       =   &H00C0C0C0&
         Height          =   615
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Tag             =   "XYZ"
         Top             =   7200
         Width           =   6015
      End
      Begin VB.TextBox txtEncodedBy 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Tag             =   "XYZ"
         Top             =   7920
         Width           =   3015
      End
      Begin VB.TextBox txtLastUpdate 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Tag             =   "XYZ"
         Top             =   8295
         Width           =   3015
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Left            =   2210
         TabIndex        =   6
         Top             =   270
         Width           =   255
      End
      Begin VB.ComboBox cboSupplier 
         Height          =   315
         Left            =   12480
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   240
         Width           =   3495
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "All Suppliers"
         Height          =   315
         Left            =   14760
         TabIndex        =   3
         Tag             =   "XYF"
         Top             =   600
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Left            =   12480
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtFindPart 
         Height          =   375
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   6240
         TabIndex        =   4
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "&Find"
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
         Image           =   "frmNewPricing.frx":001E
         cBack           =   -2147483633
      End
      Begin MSFlexGridLib.MSFlexGrid flxSRP 
         Height          =   6015
         Left            =   120
         TabIndex        =   10
         Tag             =   "XYZ"
         Top             =   960
         Width           =   15855
         _ExtentX        =   27966
         _ExtentY        =   10610
         _Version        =   393216
         Cols            =   15
         RowHeightMin    =   315
         BackColor       =   16777215
         BackColorBkg    =   12632256
         HighLight       =   2
         GridLines       =   2
         AllowUserResizing=   1
         FormatString    =   $"frmNewPricing.frx":5810
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
      Begin VB.Label lblAppBy 
         BackStyle       =   0  'Transparent
         Caption         =   "APPROVED BY:"
         Height          =   255
         Left            =   10920
         TabIndex        =   20
         Tag             =   "XYZ"
         Top             =   8040
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "ENCODED BY:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Tag             =   "XYZ"
         Top             =   7920
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "REMARKS:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Tag             =   "XYZ"
         Top             =   7200
         Width           =   1095
      End
      Begin VB.Label lblStat 
         BackStyle       =   0  'Transparent
         Caption         =   "STATUS:"
         Height          =   255
         Left            =   10920
         TabIndex        =   17
         Tag             =   "XYZ"
         Top             =   7320
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "LAST UPDATE:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Tag             =   "XYZ"
         Top             =   8280
         Width           =   1215
      End
      Begin VB.Label lblFindParts 
         Caption         =   "FIND PART DESCRIPTION:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   285
         Width           =   2055
      End
      Begin VB.Label Label12 
         Caption         =   "SUPPLIER:"
         Height          =   255
         Left            =   11520
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "TYPE:"
         Height          =   255
         Left            =   11520
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   600
         Width           =   495
      End
   End
   Begin lvButton.lvButtons_H cmdSaveExcel 
      Height          =   375
      Left            =   3720
      TabIndex        =   21
      Top             =   120
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "&SAVE EXCEL"
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
      Image           =   "frmNewPricing.frx":58AF
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   661
      Caption         =   "&ADD"
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
      Image           =   "frmNewPricing.frx":5E49
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   23
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
      Image           =   "frmNewPricing.frx":63E3
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   24
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
      Image           =   "frmNewPricing.frx":697D
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   5280
      TabIndex        =   25
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
      Image           =   "frmNewPricing.frx":C16F
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   9480
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   16545
      m_MDI_Height    =   10035
   End
End
Attribute VB_Name = "frmNewPricing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
