VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmAcctPayInvcg 
   Caption         =   "SUPPLIER PRODUCT COST RECEIVING RECEIPT "
   ClientHeight    =   10950
   ClientLeft      =   165
   ClientTop       =   -1545
   ClientWidth     =   15270
   ControlBox      =   0   'False
   Icon            =   "frmAcctPayInvcg.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9240
      Top             =   0
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   5160
      TabIndex        =   78
      Top             =   120
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   661
      Caption         =   "&REFRESH"
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
      Image           =   "frmAcctPayInvcg.frx":57E2
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtTag 
      Height          =   375
      Left            =   15960
      TabIndex        =   77
      Text            =   "txtTag"
      Top             =   5040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer tmrGetRecord 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10560
      Top             =   0
   End
   Begin VB.TextBox txtFlxVat 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   15960
      MaxLength       =   16
      TabIndex        =   65
      Tag             =   "XYZ"
      Text            =   "txtFlxVat"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1095
      Left            =   5880
      TabIndex        =   59
      Top             =   1440
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1931
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "FIND PART NUMBER"
      TabPicture(0)   =   "frmAcctPayInvcg.frx":5D7C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblClose"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdFind"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtFind"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.TextBox txtFind 
         Height          =   405
         Left            =   120
         TabIndex        =   60
         Top             =   525
         Width           =   4095
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   4320
         TabIndex        =   61
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "Find"
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
         Image           =   "frmAcctPayInvcg.frx":5D98
         cBack           =   -2147483633
      End
      Begin VB.Label lblClose 
         Alignment       =   2  'Center
         Caption         =   " ( X )"
         Height          =   255
         Left            =   4920
         MousePointer    =   99  'Custom
         TabIndex        =   62
         Top             =   120
         Width           =   375
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   7
      Left            =   15960
      TabIndex        =   43
      Top             =   480
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
      Image           =   "frmAcctPayInvcg.frx":B58A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   15960
      TabIndex        =   44
      Top             =   120
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
      Image           =   "frmAcctPayInvcg.frx":BB24
      cBack           =   -2147483633
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
      Image           =   "frmAcctPayInvcg.frx":C0BE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   9
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
      Image           =   "frmAcctPayInvcg.frx":C658
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   10
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
      Image           =   "frmAcctPayInvcg.frx":CBF2
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   3720
      TabIndex        =   11
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
      Image           =   "frmAcctPayInvcg.frx":123E4
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   6600
      TabIndex        =   12
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
      Image           =   "frmAcctPayInvcg.frx":1297E
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtFlxDisType 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   15960
      MaxLength       =   50
      TabIndex        =   50
      Tag             =   "XYZ"
      Text            =   "txtFlxDisType"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFlxValue 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   15960
      MaxLength       =   16
      TabIndex        =   49
      Tag             =   "XYZ"
      Text            =   "txtFlxValue"
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFlxInvPrice 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   15960
      MaxLength       =   16
      TabIndex        =   48
      Tag             =   "XYZ"
      Text            =   "txtFlxInvPrice"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   15960
      TabIndex        =   46
      Text            =   "HoldBackColor"
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   15960
      Locked          =   -1  'True
      TabIndex        =   45
      Text            =   "txtHoldGrayBackColor"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraHoldBottom 
      Height          =   8175
      Left            =   120
      TabIndex        =   27
      Tag             =   "XYF"
      Top             =   2640
      Width           =   15015
      Begin VB.TextBox txtDisNote 
         BackColor       =   &H00C0C0C0&
         Height          =   2175
         Left            =   5160
         Locked          =   -1  'True
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   75
         Tag             =   "XYZ"
         Top             =   3720
         Width           =   2895
      End
      Begin VB.TextBox txtTotPayINV 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   13455
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Tag             =   "XYZ"
         Text            =   "0.00"
         Top             =   5880
         Width           =   1455
      End
      Begin VB.TextBox txtTotVat 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   13455
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   58
         Tag             =   "XYZ"
         Text            =   "0.00"
         Top             =   5520
         Width           =   1455
      End
      Begin VB.TextBox txtOtherAmnt 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   13455
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   66
         Tag             =   "XYZ"
         Text            =   "0.00"
         Top             =   5160
         Width           =   1455
      End
      Begin VB.TextBox txtLessINV 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   13455
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Tag             =   "XYZ"
         Text            =   "0.00"
         Top             =   4800
         Width           =   1455
      End
      Begin VB.TextBox txtLessINVPer3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   13455
         Locked          =   -1  'True
         MaxLength       =   7
         ScrollBars      =   2  'Vertical
         TabIndex        =   56
         Tag             =   "XYZ"
         Text            =   "0%"
         Top             =   4440
         Width           =   1455
      End
      Begin VB.TextBox txtLessINVPer2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   13455
         Locked          =   -1  'True
         MaxLength       =   7
         ScrollBars      =   2  'Vertical
         TabIndex        =   53
         Tag             =   "XYZ"
         Text            =   "0%"
         Top             =   4080
         Width           =   1455
      End
      Begin VB.TextBox txtLessINVPer1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   13455
         Locked          =   -1  'True
         MaxLength       =   7
         ScrollBars      =   2  'Vertical
         TabIndex        =   51
         Tag             =   "XYZ"
         Text            =   "0%"
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox txtGrossTotINV 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   13455
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Tag             =   "XYZ"
         Text            =   "0.00"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.ComboBox cboStat 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         ItemData        =   "frmAcctPayInvcg.frx":18170
         Left            =   8160
         List            =   "frmAcctPayInvcg.frx":1817A
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Tag             =   "XYZ"
         Top             =   6360
         Width           =   975
      End
      Begin VB.TextBox txtAppBy 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   3720
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Tag             =   "XYZ"
         Text            =   "LORENZO C. RUBINOS JR."
         Top             =   7440
         Width           =   2535
      End
      Begin VB.TextBox txtDateRecv 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   12840
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Tag             =   "XYZ"
         Top             =   7680
         Width           =   2055
      End
      Begin VB.TextBox txtPrepBy 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Tag             =   "XYZ"
         Top             =   7440
         Width           =   2535
      End
      Begin VB.TextBox txtRecBy 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   8160
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Tag             =   "XYZ"
         Top             =   7680
         Width           =   2535
      End
      Begin VB.TextBox txtNotes 
         BackColor       =   &H00C0C0C0&
         Height          =   855
         Left            =   8160
         Locked          =   -1  'True
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Tag             =   "XYZ"
         Top             =   6720
         Width           =   6735
      End
      Begin MSFlexGridLib.MSFlexGrid flxAPI 
         Height          =   3135
         Left            =   120
         TabIndex        =   4
         Tag             =   "XYZ"
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   14
         RowHeightMin    =   315
         BackColor       =   16777215
         BackColorBkg    =   12632256
         HighLight       =   2
         GridLines       =   2
         AllowUserResizing=   3
         FormatString    =   $"frmAcctPayInvcg.frx":1818E
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
      Begin MSFlexGridLib.MSFlexGrid flxDiscount 
         Height          =   2655
         Left            =   120
         TabIndex        =   5
         Tag             =   "XYZ"
         Top             =   3480
         Width           =   4950
         _ExtentX        =   8731
         _ExtentY        =   4683
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         FixedCols       =   0
         RowHeightMin    =   315
         BackColorBkg    =   14737632
         HighLight       =   2
         GridLines       =   2
         AllowUserResizing=   3
         FormatString    =   "     |^INVOICE DISCOUNT |^VALUE  |^DISCOUNT TYPE               "
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
      Begin VB.Frame Frame1 
         Height          =   1695
         Left            =   8160
         TabIndex        =   68
         Tag             =   "XYZ"
         Top             =   3600
         Width           =   3975
         Begin VB.TextBox txtChckNo 
            BackColor       =   &H00C0C0C0&
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            MaxLength       =   20
            ScrollBars      =   2  'Vertical
            TabIndex        =   71
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtVcherNo 
            BackColor       =   &H00C0C0C0&
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            MaxLength       =   20
            ScrollBars      =   2  'Vertical
            TabIndex        =   70
            Top             =   600
            Width           =   1815
         End
         Begin MSMask.MaskEdBox txtChckDate 
            Height          =   375
            Left            =   2040
            TabIndex        =   69
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BackColor       =   12632256
            Enabled         =   0   'False
            MaxLength       =   14
            Mask            =   "## / ## / ####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "CHECK VOUCHER NO. :"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "CHECK NO. :"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "CHECK DATE :"
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   960
            Width           =   1455
         End
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "DISCOUNT NOTE:"
         Height          =   255
         Left            =   5160
         TabIndex        =   76
         Tag             =   "XYZ"
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "OTHER AMOUNT:"
         Height          =   255
         Left            =   11880
         TabIndex        =   67
         Tag             =   "XYZ"
         Top             =   5280
         Width           =   1455
      End
      Begin VB.Label lblVat 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "VAT:"
         Height          =   255
         Left            =   11880
         TabIndex        =   57
         Tag             =   "XYZ"
         Top             =   5640
         Width           =   1455
      End
      Begin VB.Label lblPer3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "% 3:"
         Height          =   255
         Left            =   11880
         TabIndex        =   55
         Tag             =   "XYZ"
         Top             =   4500
         Width           =   1455
      End
      Begin VB.Label lblPer2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "% 2:"
         Height          =   255
         Left            =   11880
         TabIndex        =   54
         Tag             =   "XYZ"
         Top             =   4140
         Width           =   1455
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "AMOUNT :"
         Height          =   255
         Left            =   11880
         TabIndex        =   52
         Tag             =   "XYZ"
         Top             =   4860
         Width           =   1455
      End
      Begin VB.Label lblAppBy 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "APPROVED BY:"
         Height          =   255
         Left            =   3720
         TabIndex        =   42
         Tag             =   "XYZ"
         Top             =   7800
         Width           =   2535
      End
      Begin VB.Label lblDateRecv 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DATE RECEIVED:"
         Height          =   255
         Left            =   11280
         TabIndex        =   41
         Tag             =   "XYZ"
         Top             =   7680
         Width           =   1455
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PREPARED BY:"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Tag             =   "XYZ"
         Top             =   7800
         Width           =   2535
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "RECEIVED BY:"
         Height          =   255
         Left            =   6600
         TabIndex        =   39
         Tag             =   "XYZ"
         Top             =   7680
         Width           =   1455
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NOTES:"
         Height          =   615
         Left            =   7320
         TabIndex        =   38
         Tag             =   "XYZ"
         Top             =   6720
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "STATUS:"
         Height          =   255
         Left            =   7200
         TabIndex        =   37
         Tag             =   "XYZ"
         Top             =   6360
         Width           =   855
      End
      Begin VB.Label llblTotPay 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL PAYABLE:"
         Height          =   255
         Left            =   11880
         TabIndex        =   36
         Tag             =   "XYZ"
         Top             =   5940
         Width           =   1455
      End
      Begin VB.Label lblPer1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "% 1:"
         Height          =   255
         Left            =   11880
         TabIndex        =   35
         Tag             =   "XYZ"
         Top             =   3780
         Width           =   1455
      End
      Begin VB.Label lblGrossTot 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "GROSS TOTAL:"
         Height          =   255
         Left            =   11880
         TabIndex        =   34
         Tag             =   "XYZ"
         Top             =   3420
         Width           =   1455
      End
   End
   Begin VB.Frame fraHoldTop 
      Height          =   2175
      Left            =   120
      TabIndex        =   13
      Tag             =   "XYF"
      Top             =   480
      Width           =   15015
      Begin VB.ComboBox cboInvType 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         ItemData        =   "frmAcctPayInvcg.frx":18241
         Left            =   12720
         List            =   "frmAcctPayInvcg.frx":1824E
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Tag             =   "XYZ"
         Top             =   1750
         Width           =   1935
      End
      Begin VB.TextBox txtPONo 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Tag             =   "XYZ"
         Top             =   1365
         Width           =   3855
      End
      Begin VB.TextBox txtWayBillNo 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Tag             =   "XYZ"
         Top             =   990
         Width           =   3855
      End
      Begin VB.TextBox txtDlvrdBy 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Tag             =   "XYZ"
         Top             =   615
         Width           =   3855
      End
      Begin VB.TextBox txtSupName 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Tag             =   "XYZ"
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox txtDueDate 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   12720
         Locked          =   -1  'True
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Tag             =   "XYZ"
         Top             =   1365
         Width           =   1935
      End
      Begin VB.TextBox txtTerms 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   12720
         Locked          =   -1  'True
         MaxLength       =   4
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Tag             =   "XYZ"
         Top             =   990
         Width           =   1335
      End
      Begin VB.TextBox txtPckngDate 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   12720
         Locked          =   -1  'True
         MaxLength       =   11
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Tag             =   "XYZ"
         Top             =   615
         Width           =   1935
      End
      Begin VB.TextBox txtPckngNo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   12720
         Locked          =   -1  'True
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Tag             =   "XYZ"
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "INVOICE TYPE:"
         Height          =   255
         Left            =   11280
         TabIndex        =   64
         Tag             =   "XYZ"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblDays 
         BackStyle       =   0  'Transparent
         Caption         =   "DAYS"
         Height          =   255
         Left            =   14160
         TabIndex        =   47
         Tag             =   "XYZ"
         Top             =   1050
         Width           =   495
      End
      Begin VB.Label lblDueDate 
         BackStyle       =   0  'Transparent
         Caption         =   "DUE DATE:"
         Height          =   255
         Left            =   11280
         TabIndex        =   26
         Tag             =   "XYZ"
         Top             =   1425
         Width           =   1215
      End
      Begin VB.Label lblTerms 
         BackStyle       =   0  'Transparent
         Caption         =   "TERMS:"
         Height          =   255
         Left            =   11280
         TabIndex        =   25
         Tag             =   "XYZ"
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label lblPckngDate 
         BackStyle       =   0  'Transparent
         Caption         =   "PACKING DATE:"
         Height          =   255
         Left            =   11280
         TabIndex        =   24
         Tag             =   "XYZ"
         Top             =   675
         Width           =   1335
      End
      Begin VB.Label lblPckngNo 
         BackStyle       =   0  'Transparent
         Caption         =   "PACKING NO.:"
         Height          =   255
         Left            =   11280
         TabIndex        =   23
         Tag             =   "XYZ"
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "P.O. NUMBER:"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Tag             =   "XYZ"
         Top             =   1425
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "WAY BILL NO.:"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Tag             =   "XYZ"
         Top             =   1050
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DELIVERED BY:"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Tag             =   "XYZ"
         Top             =   675
         Width           =   1455
      End
      Begin VB.Label lblStartBookNo 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER NAME:"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Tag             =   "XYZ"
         Top             =   300
         Width           =   1455
      End
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   9840
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   15390
      m_MDI_Height    =   11430
      OriginalBorder  =   2
      OriginalCaption =   25
   End
   Begin VB.Menu mnuInvFrm 
      Caption         =   "Invoice"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuPopUpFormat 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuSTFindRow 
         Caption         =   "FIND PART NUMBER"
         Shortcut        =   ^F
      End
   End
End
Attribute VB_Name = "frmAcctPayInvcg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmpRs As New ADODB.Recordset
Dim PrtSrch As Boolean

Dim didRecordEntry As Boolean
Dim formAction As String
Dim DtlsCol As Integer
Dim DtlsRow As Integer

Dim flxRowNo As Integer
Dim flxColNo As Integer
Dim flxTTLCol As Integer

Dim sqlPckno As String
Dim sqlPONo As String

Sub DisNoteValue()
On Error GoTo er
    Dim DisNote1 As String
    Dim DisNote2 As String
    Dim DisNote3 As String
    Dim DisNote4 As String
    Dim DisNote5 As String
    Dim Note As String
    
    If flxDiscount.TextMatrix(1, 3) <> "" And flxDiscount.TextMatrix(1, 0) = "[X]" Then
        DisNote1 = "LESS" & " " & flxDiscount.TextMatrix(1, 2) & " - " & flxDiscount.TextMatrix(1, 3)
        Note = DisNote1
    End If
    If flxDiscount.TextMatrix(2, 3) <> "" And flxDiscount.TextMatrix(2, 0) = "[X]" Then
        DisNote2 = "LESS" & " " & flxDiscount.TextMatrix(2, 2) & " - " & flxDiscount.TextMatrix(2, 3)
        
        If DisNote1 <> "" Then
            Note = DisNote1 & vbNewLine & DisNote2
        Else
            Note = DisNote2
        End If
        
    End If
    If flxDiscount.TextMatrix(3, 3) <> "" And flxDiscount.TextMatrix(3, 0) = "[X]" Then
        DisNote3 = "LESS" & " " & flxDiscount.TextMatrix(3, 2) & " - " & flxDiscount.TextMatrix(3, 3)
        
        If DisNote1 <> "" And DisNote2 <> "" Then
            Note = DisNote1 & vbNewLine & DisNote2 & vbNewLine & DisNote3
        Else
            Note = DisNote3
        End If
        
    End If
    If flxDiscount.TextMatrix(4, 3) <> "" And flxDiscount.TextMatrix(4, 0) = "[X]" Then
        DisNote4 = "LESS" & " " & flxDiscount.TextMatrix(4, 2) & " - " & flxDiscount.TextMatrix(4, 3)
                
        If DisNote1 <> "" And DisNote2 <> "" And DisNote3 <> "" Then
            Note = DisNote1 & vbNewLine & DisNote2 & vbNewLine & DisNote3 & vbNewLine & DisNote4
        Else
            Note = DisNote4
        End If
        
    End If
    If flxDiscount.TextMatrix(5, 3) <> "" And flxDiscount.TextMatrix(5, 0) = "[X]" Then
        DisNote5 = "LESS" & " " & flxDiscount.TextMatrix(5, 2) & " - " & flxDiscount.TextMatrix(5, 3)
        Note = DisNote1 & vbNewLine & DisNote2 & vbNewLine & DisNote3 & vbNewLine & DisNote4 & vbNewLine & DisNote5
    End If
    
    txtDisNote.Text = Note


Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error -> DisNoteValue()")
End Sub

Public Sub ShowRecord(ByVal pPACKINGNO As String)
On Error GoTo er
    Dim i As Integer
    
    Call modAPI.sp_load_packing_info(pPACKINGNO, tmpRs)
    If tmpRs.State = 0 Then Exit Sub
    
    If tmpRs.RecordCount <= 0 Then
        tmpRs.Close
        Exit Sub
    End If
    
    
    tmpRs.MoveFirst
    
            txtSupName.Text = tmpRs!Supplier
            txtDlvrdBy.Text = tmpRs!Shipper
            txtWayBillNo.Text = tmpRs!WayBillno
            txtPONo.Text = tmpRs!PACKINGLIST
            txtPckngNo.Text = tmpRs!PACKINGNO
            txtPckngDate.Text = tmpRs!PACKDATE
            txtRecBy.Text = tmpRs!receivedby
            txtDateRecv.Text = tmpRs!DATEREC
           
            txtTerms.Text = tmpRs!Terms
            txtDueDate.Text = tmpRs!DueDate
            txtNotes.Text = tmpRs!INVNOTES
            txtLessINVPer1.Text = tmpRs!INVLESPER
            txtLessINVPer2.Text = tmpRs!INVLESPER2
            txtLessINVPer3.Text = tmpRs!INVLESPER3
            txtLessINV.Text = Format(tmpRs!INVOICELESS, "#,##0.00")
            txtAppBy.Text = tmpRs!ApprovedBy
            txtChckNo.Text = tmpRs!CHCKNO
            txtVcherNo.Text = tmpRs!VOCHERNO
            txtTotVat.Text = Format(tmpRs!VatValue, "#,##0.00")
            txtTotPayINV.Text = Format(tmpRs!INVOICETTLPAY, "#,##0.00")
            txtOtherAmnt.Text = tmpRs!OTHERAMNT
            txtDisNote.Text = tmpRs!disnote
            
            If IsNull(tmpRs!ChckDate) = True Or tmpRs!ChckDate = "" Then
                txtChckDate.Mask = ""
                txtChckDate.Text = ""
                txtChckDate.Mask = "## / ## / ####"
            Else
                txtChckDate.Mask = Format(tmpRs!ChckDate, "mm / dd / yyyy")
            End If
            
            Call modObjTools.changeListIndex(cboStat, tmpRs!INVSTATUS)
            
            Call modObjTools.changeListIndex(cboInvType, tmpRs!InvType)
            
            flxDiscount.TextMatrix(1, 2) = tmpRs!DISC1VAL
            flxDiscount.TextMatrix(2, 2) = tmpRs!DISC2VAL
            flxDiscount.TextMatrix(3, 2) = tmpRs!DISC3VAL
            flxDiscount.TextMatrix(4, 2) = tmpRs!DISC4VAL
            flxDiscount.TextMatrix(5, 2) = tmpRs!DISC5VAL
            flxDiscount.TextMatrix(6, 2) = tmpRs!vat

            flxDiscount.TextMatrix(1, 3) = tmpRs!DISC1TYPE
            flxDiscount.TextMatrix(2, 3) = tmpRs!DISC2TYPE
            flxDiscount.TextMatrix(3, 3) = tmpRs!DISC3TYPE
            flxDiscount.TextMatrix(4, 3) = tmpRs!DISC4TYPE
            flxDiscount.TextMatrix(5, 3) = tmpRs!DISC5TYPE
       
       

        While Not tmpRs.EOF
        
'            flxAPI.TextMatrix(DtlsRow, 6) = Format(tmpRs!INVOICEPRICE, "#,##0.00")
'            flxAPI.TextMatrix(DtlsRow, 7) = Format(tmpRs!TTLDISCOUNT, "#,##0.00")
'            flxAPI.TextMatrix(DtlsRow, 8) = Format(tmpRs!VATAMOUNT, "#,##0.00")
'            flxAPI.TextMatrix(DtlsRow, 9) = Format(tmpRs!NewCost, "#,##0.00")
'              If tmpRs!MrcCost <> "" Then
'                    flxAPI.TextMatrix(DtlsRow, 10) = Format(tmpRs!MrcCost, "#,##0.00")
'              End If
'            flxAPI.TextMatrix(DtlsRow, 11) = Format(tmpRs!PriceDiff, "#,##0.00")
'            flxAPI.TextMatrix(DtlsRow, 12) = Format(tmpRs!PriceDiffPer, "#,##0.00")
'            flxAPI.TextMatrix(DtlsRow, 13) = Format(tmpRs!INVOICETOTAL, "#,##0.00")


'          1           2            3          4            5           6            7       8        9          10          11          12             13           14            15
'" |<PART NUMBER |<DESCRIPTION |^WHS QTY |^BRANCH QTY |^ORDER QTY. |^INV. PRICE |^DISCOUNT |^VAT |^NEW COST |^OLD COST |^PRICE DIFF.|^PRICE DIFF.% |^INV TOTAL " 'SPAREPARTS ID  SPRR NO
            
            If flxAPI.TextMatrix(1, 1) = "" Then
                flxAPI.TextMatrix(1, 1) = tmpRs!PARTNUMBER
                flxAPI.TextMatrix(1, 2) = tmpRs!Description
                flxAPI.TextMatrix(1, 3) = tmpRs!whsqty
                flxAPI.TextMatrix(1, 4) = tmpRs!bqty
                flxAPI.TextMatrix(1, 5) = tmpRs!qty
                flxAPI.TextMatrix(1, 6) = Format(tmpRs!INVOICEPRICE, "#,##0.00")
                flxAPI.TextMatrix(1, 7) = Format(tmpRs!TTLDISCOUNT, "#,##0.00")
                flxAPI.TextMatrix(1, 8) = Format(tmpRs!VATAMOUNT, "#,##0.00")
                flxAPI.TextMatrix(1, 9) = Format(tmpRs!NewCost, "#,##0.00")
                flxAPI.TextMatrix(1, 10) = Format(tmpRs!MrcCost, "#,##0.00")
                flxAPI.TextMatrix(1, 11) = Format(tmpRs!PRICEDIFF, "#,##0.00")
                flxAPI.TextMatrix(1, 12) = tmpRs!PRICEDIFFPER
                flxAPI.TextMatrix(1, 13) = Format(tmpRs!INVOICETOTAL, "#,##0.00")
                flxAPI.TextMatrix(1, 14) = tmpRs!SPAREPARTSID
                flxAPI.TextMatrix(1, 15) = tmpRs!SPRRNO
            Else
                flxAPI.AddItem _
                vbTab & tmpRs!PARTNUMBER & _
                vbTab & tmpRs!Description & _
                vbTab & tmpRs!whsqty & _
                vbTab & tmpRs!bqty & _
                vbTab & tmpRs!qty & _
                vbTab & Format(tmpRs!INVOICEPRICE, "#,##0.00") & _
                vbTab & Format(tmpRs!TTLDISCOUNT, "#,##0.00") & _
                vbTab & Format(tmpRs!VATAMOUNT, "#,##0.00") & _
                vbTab & Format(tmpRs!NewCost, "#,##0.00") & _
                vbTab & Format(tmpRs!MrcCost, "#,##0.00") & _
                vbTab & Format(tmpRs!PRICEDIFF, "#,##0.00") & _
                vbTab & tmpRs!PRICEDIFFPER & _
                vbTab & Format(tmpRs!INVOICETOTAL, "#,##0.00") & _
                vbTab & tmpRs!SPAREPARTSID & _
                vbTab & tmpRs!SPRRNO
            End If

    
            
            For i = 1 To 5
                If flxDiscount.TextMatrix(i, 2) = "" Then
                    flxDiscount.TextMatrix(i, 0) = "[  ]"
                Else
                    flxDiscount.TextMatrix(i, 0) = "[X]"
                End If
            Next
            
        tmpRs.MoveNext
    Wend
    
    DoEvents
    If flxAPI.Rows = 1 Then
        flxAPI.Rows = 2
    End If
    
    If tmpRs.State <> 0 Then tmpRs.Close
    
    
    DoEvents
    Call Total
    'Call TotalGross

    DoEvents
    Call EnabledSelectedButton
    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->ShowRecord()")
End Sub

Public Sub RefreshRecord(ByVal pPACKINGNO As String)
On Error GoTo er
    Dim row As Integer
    
    Call modAPI.sp_refresh_packing_info(pPACKINGNO, tmpRs)
    If tmpRs.State = 0 Then Exit Sub
    
    If tmpRs.RecordCount <= 0 Then
        tmpRs.Close
        Exit Sub
    End If

    row = 1
    tmpRs.MoveFirst
        
        While Not tmpRs.EOF
        
'          1           2            3          4            5           6            7       8        9          10          11          12             13           14            15
'" |<PART NUMBER |<DESCRIPTION |^WHS QTY |^BRANCH QTY |^ORDER QTY. |^INV. PRICE |^DISCOUNT |^VAT |^NEW COST |^OLD COST |^PRICE DIFF.|^PRICE DIFF.% |^INV TOTAL " 'SPAREPARTS ID  SPRR NO
                
            
                flxAPI.TextMatrix(row, 3) = tmpRs!whsqty
                flxAPI.TextMatrix(row, 4) = tmpRs!bqty
                flxAPI.TextMatrix(row, 5) = tmpRs!OrderQty
                flxAPI.TextMatrix(row, 10) = Format(tmpRs!Cost, "#,##0.00")
                
                row = row + 1
    tmpRs.MoveNext
        Wend
        
    
    
    
    
    If tmpRs.State <> 0 Then tmpRs.Close
    
    DoEvents
    Call Total
    'Call TotalGross

    DoEvents
    Call EnabledSelectedButton
    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->ShowRecord()")
End Sub


Public Sub GetRecord(ByVal pPACKINGNO As String)
On Error GoTo er
    Dim Level As String
    Dim message As String

    Call modAPI.sp_get_packing_info(pPACKINGNO, tmpRs, Level, message)
    If tmpRs.State = 0 Then Exit Sub
    
    If Level = "Warning" Then
        MsgBox message, vbInformation, "GET PACKING INFO"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If tmpRs.RecordCount <= 0 Then
        MsgBox "NO RECORD FOUND!", vbInformation, "PACKING NO."
        
        Screen.MousePointer = vbDefault
        'Highlight the txtRefNo
        txtPckngNo.SelStart = 0
        txtPckngNo.SelLength = Len(txtPckngNo.Text)
        txtPckngNo.SetFocus
    
        tmpRs.Close
        Exit Sub
    End If
    
    flxAPI.Rows = 1
    flxAPI.Rows = 2
    
    Call flxAPIColsChanger
    
'          1           2            3          4            5           6            7       8        9          10          11          12             13           14            15
'" |<PART NUMBER |<DESCRIPTION |^WHS QTY |^BRANCH QTY |^ORDER QTY. |^INV. PRICE |^DISCOUNT |^VAT |^NEW COST |^OLD COST |^PRICE DIFF.|^PRICE DIFF.% |^INV TOTAL " 'SPAREPARTS ID  SPRR NO
   
    tmpRs.MoveFirst
    
            txtSupName.Text = tmpRs!Supplier
            txtDlvrdBy.Text = tmpRs!Shipper
            txtWayBillNo.Text = tmpRs!WayBillno
            txtPONo.Text = tmpRs!PACKINGLIST
            txtPckngNo.Text = tmpRs!PACKINGNO
            txtPckngDate.Text = tmpRs!PACKDATE
            txtRecBy.Text = tmpRs!receivedby
            txtDateRecv.Text = tmpRs!DATEREC
    
        While Not tmpRs.EOF
            If flxAPI.TextMatrix(1, 1) = "" Then
                flxAPI.TextMatrix(1, 1) = tmpRs!PARTNUMBER
                flxAPI.TextMatrix(1, 2) = tmpRs!Description
                flxAPI.TextMatrix(1, 3) = tmpRs!whsqty
                flxAPI.TextMatrix(1, 4) = tmpRs!bqty
                flxAPI.TextMatrix(1, 5) = tmpRs!OrderQty
                flxAPI.TextMatrix(1, 6) = ""
                flxAPI.TextMatrix(1, 7) = ""
                flxAPI.TextMatrix(1, 8) = ""
                flxAPI.TextMatrix(1, 9) = ""
                flxAPI.TextMatrix(1, 10) = Format(tmpRs!Cost, "#,##0.00")
                flxAPI.TextMatrix(1, 11) = ""
                flxAPI.TextMatrix(1, 12) = ""
                flxAPI.TextMatrix(1, 13) = ""
                flxAPI.TextMatrix(1, 14) = tmpRs!SPAREPARTID
                flxAPI.TextMatrix(1, 15) = tmpRs!SPRRNO
            Else
                flxAPI.AddItem _
                vbTab & tmpRs!PARTNUMBER & _
                vbTab & tmpRs!Description & _
                vbTab & tmpRs!whsqty & _
                vbTab & tmpRs!bqty & _
                vbTab & tmpRs!OrderQty & _
                vbTab & "" & _
                vbTab & "" & _
                vbTab & "" & _
                vbTab & "" & _
                vbTab & Format(tmpRs!Cost, "#,##0.00") & _
                vbTab & "" & _
                vbTab & "" & _
                vbTab & "" & _
                vbTab & tmpRs!SPAREPARTID & _
                vbTab & tmpRs!SPRRNO
            End If
        
    tmpRs.MoveNext
    Wend
    
    DoEvents
    
    If flxAPI.Rows = 1 Then
        flxAPI.Rows = 2
    End If
    
    If tmpRs.State <> 0 Then tmpRs.Close
    
    If txtTerms.Locked = False Then
        txtTerms.SetFocus
    End If
    'Call ShowRecord(pPACKINGNO)
    
    txtPrepBy.Text = EmployeesName
    
    Screen.MousePointer = vbDefault
    
Exit Sub
er:
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->GetRecord()")
End Sub

Private Sub SetButtonVisibility(DataEntry As Boolean)
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->PRINT
'3 ->SUMMRY
'4 ->REFRESH
'5 ->EXIT
'6 ->SAVE
'7 ->CANCEL
    
    Dim i As Integer
    
    'hide all button
    For i = 0 To 7
        cmdTrans(i).Visible = False
    Next i
    
    If DataEntry Then
        cmdTrans(6).Visible = True
        cmdTrans(7).Visible = True
    Else
        For i = 0 To 5
            cmdTrans(i).Visible = True
        Next i
        
    End If
    
End Sub

Sub LockedEntryObject(DataEntry As Boolean)

    DataEntry = Not DataEntry
        
        txtPckngNo.Locked = DataEntry
        txtTerms.Locked = DataEntry
        txtDueDate.Locked = DataEntry
        txtLessINVPer1.Locked = DataEntry
        txtLessINVPer2.Locked = DataEntry
        txtLessINVPer3.Locked = DataEntry
        txtNotes.Locked = DataEntry
        txtAppBy.Locked = DataEntry
        txtChckNo.Locked = DataEntry
        txtVcherNo.Locked = DataEntry
        txtOtherAmnt.Locked = DataEntry
        txtDisNote.Locked = DataEntry
                                                                                                                                                                                                                                 
    If txtPckngNo.Locked Then
   
        txtPckngNo.BackColor = txtHoldGrayBackColor.BackColor
        txtTerms.BackColor = txtHoldGrayBackColor.BackColor
        txtDueDate.BackColor = txtHoldGrayBackColor.BackColor
        txtLessINVPer1.BackColor = txtHoldGrayBackColor.BackColor
        txtLessINVPer2.BackColor = txtHoldGrayBackColor.BackColor
        txtLessINVPer3.BackColor = txtHoldGrayBackColor.BackColor
        txtNotes.BackColor = txtHoldGrayBackColor.BackColor
        txtAppBy.BackColor = txtHoldGrayBackColor.BackColor
        txtChckNo.BackColor = txtHoldGrayBackColor.BackColor
        txtVcherNo.BackColor = txtHoldGrayBackColor.BackColor
        txtOtherAmnt.BackColor = txtHoldGrayBackColor.BackColor
        txtDisNote.BackColor = txtHoldGrayBackColor.BackColor
    Else
        
        txtPckngNo.BackColor = txtHoldWhiteBackColor.BackColor
        txtTerms.BackColor = txtHoldWhiteBackColor.BackColor
        txtDueDate.BackColor = txtHoldWhiteBackColor.BackColor
        txtLessINVPer1.BackColor = txtHoldWhiteBackColor.BackColor
        txtLessINVPer2.BackColor = txtHoldWhiteBackColor.BackColor
        txtLessINVPer3.BackColor = txtHoldWhiteBackColor.BackColor
        txtNotes.BackColor = txtHoldWhiteBackColor.BackColor
        txtAppBy.BackColor = txtHoldWhiteBackColor.BackColor
        txtChckNo.BackColor = txtHoldWhiteBackColor.BackColor
        txtVcherNo.BackColor = txtHoldWhiteBackColor.BackColor
        txtOtherAmnt.BackColor = txtHoldWhiteBackColor.BackColor
        txtDisNote.BackColor = txtHoldWhiteBackColor.BackColor
    End If

End Sub

Private Sub CreatTempTable()
On Error GoTo er
        Dim Qry As String

        Qry = "Drop Temporary Table IF Exists sp_acctpayinvcng.tmp_suppinvdtls;"
        Call LYRcon.Execute(Qry)
        DoEvents
        Qry = "Create Temporary Table IF Not Exists sp_acctpayinvcng.tmp_suppinvdtls( " & _
        "   PACKINGNO           varchar(20),        " & _
        "   SPAREPARTSID        INT(10),            " & _
        "   QTY                 INT(10),            " & _
        "   MRCCOST             DECIMAL(16,4),      " & _
        "   PRICEDIFF           DECIMAL(16,4),      " & _
        "   PRICEDIFFPER        VARCHAR(10),        " & _
        "   INVOICEPRICE        DECIMAL(16,4),      " & _
        "   TTLDISCOUNT         DECIMAL(16,4),      " & _
        "   VATAMOUNT           DECIMAL(16,4),      " & _
        "   NEWCOST             DECIMAL(16,4),      " & _
        "   INVOICETOTAL        DECIMAL(16,4),      " & _
        "   WHSQTY              INT(10),            " & _
        "   BQTY                INT(10)             " & _
        ");"

        'Text1.Text = Qry
        Call LYRcon.Execute(Qry)
        
        
        DoEvents
        Call InsertTempTable
        
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error ->CreatTempTable()")
End Sub

Private Sub InsertTempTable()
On Error GoTo er
        Dim Qry As String
        
        Dim SQNO As Integer
        Dim i As Integer
        
'          1           2            3          4            5           6            7       8        9          10          11          12             13           14            15
'" |<PART NUMBER |<DESCRIPTION |^WHS QTY |^BRANCH QTY |^ORDER QTY. |^INV. PRICE |^DISCOUNT |^VAT |^NEW COST |^OLD COST |^PRICE DIFF.|^PRICE DIFF.% |^INV TOTAL " 'SPAREPARTS ID  SPRR NO

        
        Qry = "Insert Into sp_acctpayinvcng.tmp_suppinvdtls(PACKINGNO, SPAREPARTSID,WHSQTY, BQTY, QTY, INVOICEPRICE, TTLDISCOUNT, VATAMOUNT, NEWCOST, MRCCOST, PRICEDIFF,PRICEDIFFPER, INVOICETOTAL) Values"

        For i = 1 To flxAPI.Rows - 1
            DoEvents
            If i > 1 Then Qry = Qry & ","
        
        Qry = Qry & "('" & txtPckngNo.Text & "','" & Trim(flxAPI.TextMatrix(i, 14)) & _
                        "','" & Trim(flxAPI.TextMatrix(i, 3)) & "','" & Trim(flxAPI.TextMatrix(i, 4)) & "','" & Trim(flxAPI.TextMatrix(i, 5)) & _
                        "','" & Val(Replace(Trim(flxAPI.TextMatrix(i, 6)), ",", "")) & "','" & Val(Replace(Trim(flxAPI.TextMatrix(i, 7)), ",", "")) & "','" & Val(Replace(Trim(flxAPI.TextMatrix(i, 8)), ",", "")) & _
                        "','" & Val(Replace(Trim(flxAPI.TextMatrix(i, 9)), ",", "")) & "','" & Val(Replace(Trim(flxAPI.TextMatrix(i, 10)), ",", "")) & "','" & Val(Replace(Trim(flxAPI.TextMatrix(i, 11)), ",", "")) & _
                        "','" & flxAPI.TextMatrix(i, 12) & "','" & Val(Replace(Trim(flxAPI.TextMatrix(i, 13)), ",", "")) & "')"

        Next
        'Text2.Text = Qry
        Call LYRcon.Execute(Qry)
        'InputBox Qry, Qry, Qry
        
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error ->InsertTempTable()")
End Sub

Function ValidateEntry() As String
    Dim vMsg As String, Qry As String
    Dim i As Integer, Cntr As Integer
    
    vMsg = ""
    
    If txtPckngNo.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field packing no."
    If txtTerms.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field terms."
    If txtOtherAmnt.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on other amount."
    
    If txtDueDate.Text = "" Then
        vMsg = vMsg & vbCrLf & "Incomplete entry on Due date."
    ElseIf Not IsDate(txtDueDate.Text) Then
        vMsg = vMsg & vbCrLf & "Invalid due date."
    End If
    
    If Not IsDate(txtDueDate.Text) Then
        vMsg = vMsg & vbCrLf & "Invalid due date."
    End If
    

'    If Not IsDate(txtChckDate.Text) Or txtChckDate.Mask <> "## / ## / ####" Then
'        vMsg = vMsg & vbCrLf & "Invalid check date."
'    End If
            
    '        1            2            3          4            5            6             7                8          9        10       11          12
    '|<PART NUMBER |<DESCRIPTION |^WHS QTY |^BRANCH QTY |^ORDER QTY. |^MRC COST |^PRICE DIFFERENCE |^INV. PRICE |^DISCOUNT |^VAT  |^NEW COST  |^INV TOTAL
    
'    If flxAPI.TextMatrix(flxAPI.Rows - 1, 1) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field PART NUMBER."
'    If flxAPI.TextMatrix(flxAPI.Rows - 1, 2) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field DESCRIPTION"
'    If flxAPI.TextMatrix(flxAPI.Rows - 1, 3) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field WHS QTY."
'    If flxAPI.TextMatrix(flxAPI.Rows - 1, 4) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field BRANCH QTY."
'    If flxAPI.TextMatrix(flxAPI.Rows - 1, 5) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field ORDER QTY."
'    If flxAPI.TextMatrix(flxAPI.Rows - 1, 6) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field MRC COST"
        
    ValidateEntry = vMsg
    
End Function

Public Sub EnabledSelectedButton()
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->PRINT
'3 ->SUMMRY
'4 ->REFRESH
'5 ->EXIT
'6 ->SAVE
'7 ->CANCEL
    
    cmdTrans(1).Enabled = False 'Edit Button
    
    If cboStat.Text = "FINAL" Then
        cmdTrans(1).Enabled = False 'EDIT Button
        cmdTrans(2).Enabled = True  'PRINT Button
        cmdTrans(4).Enabled = False 'REFRESH Button
    ElseIf cboStat.Text = "PENDING" Then
        cmdTrans(1).Enabled = True 'Edit Button
        cmdTrans(2).Enabled = True 'PRINT Button
        cmdTrans(4).Enabled = True  'REFRESH Button
    Else
        cmdTrans(1).Enabled = False
        cmdTrans(2).Enabled = False
    End If
 
End Sub

Private Sub cboInvType_Click()

    If cboInvType.Text = "VAT EXCLUSIVE" Then
        lblVat.Caption = "ADD VAT"
        lblPer1.Caption = "ADD %1"
        lblPer2.Caption = "ADD %2"
        lblPer3.Caption = "ADD %3"
        lblAmount.Caption = "ADD AMOUNT"
    Else
        lblVat.Caption = "LESS VAT"
        lblPer1.Caption = "LESS %1"
        lblPer2.Caption = "LESS %2"
        lblPer3.Caption = "LESS %3"
        lblAmount.Caption = "LESS AMOUNT"
    End If
    
'        flxAPI.SetFocus
'        flxAPI.Row = 1
'        flxAPI.Col = 6

End Sub

Private Sub cboInvType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxAPI.SetFocus
        flxAPI.Col = 8
    End If
End Sub

Private Sub cmdFind_Click()

    Dim i As Integer

    If Trim(txtFind.Text) = "" Then Exit Sub
    With flxAPI
            For i = 1 To flxAPI.Rows - 1
                If Trim(flxAPI.TextMatrix(i, 1)) = Trim(txtFind.Text) Then
                    .Col = 0: .row = i
                    .ColSel = 15: .RowSel = i
                    .TopRow = i
                    .HighLight = flexHighlightAlways
                    Exit Sub
                End If
            Next i
    End With
    MsgBox "No match found!", vbInformation + vbOKOnly, "Find"
    
End Sub

Private Sub cmdTrans_Click(Index As Integer)
On Error GoTo er
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->PRINT
'3 ->SUMMRY
'4 ->REFRESH
'5 ->EXIT
'6 ->SAVE
'7 ->CANCEL

    Dim Ans As VbMsgBoxResult
    Dim Level As String
    Dim message As String
    
    Dim disvalue1 As String
    Dim disvalue2 As String
    Dim disvalue3 As String
    Dim disvalue4 As String
    Dim disvalue5 As String
    Dim DisType1 As String
    Dim DisType2 As String
    Dim DisType3 As String
    Dim DisType4 As String
    Dim DisType5 As String

    If Index = 0 Then 'ADD
        formAction = "NEW"
        Call SetButtonVisibility(True)
        Call LockedEntryObject(True)
        
        didRecordEntry = True
        
        Call ClearEntryForm(True)
        Call flxDscntRowValue
        
        txtChckDate.Enabled = True
        txtChckDate.BackColor = txtHoldWhiteBackColor.BackColor
        txtChckDate.Format = "mm / dd / yyyy"
        
        cboInvType.Locked = False
        cboInvType.BackColor = txtHoldWhiteBackColor.BackColor
    
    ElseIf Index = 1 Then 'EDIT
        formAction = "EDIT"
        Call LockedEntryObject(True)
        Call SetButtonVisibility(True)
        
        cboStat.Locked = False
        cboStat.BackColor = txtHoldWhiteBackColor.BackColor
        
        cboInvType.Locked = False
        cboInvType.BackColor = txtHoldWhiteBackColor.BackColor
        
        txtChckDate.Enabled = True
        txtChckDate.BackColor = txtHoldWhiteBackColor.BackColor
        txtChckDate.Format = "mm / dd / yyyy"
        
        didRecordEntry = True
    
    ElseIf Index = 2 Then 'PRINT
       Dim tmpRs1 As New ADODB.Recordset
       Dim tmpRs2 As New ADODB.Recordset
       Screen.MousePointer = vbHourglass
       Dim rpt As New rptAcctPayInvcng
    
        'PROCEDURE sp_acctpayinvcng.sp_print_packing_info(
        '  pPACKINGNO VarChar(20)
        '  )
        
        sqlPckno = "Call sp_acctpayinvcng.sp_print_packing_info('" & txtPckngNo.Text & "');"
        DoEvents
        If tmpRs1.State <> 0 Then tmpRs1.Close
        Set tmpRs1 = LYRcon.Execute(sqlPckno)
        
        
        sqlPONo = "SELECT DISTINCT sptr.packinglist FROM mrcwarehouse.supplierinvoice as supinvc INNER JOIN mrcwarehouse.spitr as sptr on supinvc.packingno = sptr.reference WHERE supinvc.packingno= '" & txtPckngNo.Text & "'; "
        DoEvents
        If tmpRs2.State <> 0 Then tmpRs2.Close
        Set tmpRs2 = LYRcon.Execute(sqlPONo)
        
        
        DoEvents
        'Main report
        With rpt
            .DiscardSavedData
             DoEvents
            .Database.SetDataSource tmpRs1
            .ReadRecords
        End With
       
        'Subreport Cash Attachment
        With rpt.Subreport1.OpenSubreport
            .DiscardSavedData
             DoEvents
            .Database.SetDataSource tmpRs2
            .ReadRecords
        End With
    
       frmPrint.CRViewer1.ReportSource = rpt
       frmPrint.CRViewer1.ViewReport
       frmPrint.CRViewer1.Zoom 150
       frmPrint.CRViewer1.EnableExportButton = True
       Screen.MousePointer = vbDefault
       
       DoEvents
       frmPrint.Caption = "PRINT" & " " & "-" & " " & "[" & "ACCOUNTS PAYABLE INVOICING" & "]"
       frmPrint.WindowState = vbMaximized
       frmPrint.MySource = "Supplier Invoice"
       frmPrint.Show vbModal
        
    ElseIf Index = 3 Then 'SUMMARY
        frm_OnProcess2.OwnerForm = "Costing"
        FrmSummary.mySrc = "Costing"
        FrmSummary.Show
        formAction = "EXIT"
        Unload Me
        
    ElseIf Index = 4 Then 'REFRESH
        formAction = "REFRESH"
        DoEvents
        Screen.MousePointer = vbHourglass
        tmrRefresh.Enabled = True
        DoEvents
        Call Load(frm_OnProcess2)
        frm_OnProcess2.Show vbModal

    ElseIf Index = 5 Then 'EXIT
        formAction = "EXIT"
        
        Ans = MsgBox("ARE YOU SURE YOU WANT TO EXIT?", vbQuestion + vbYesNo, Me.Caption)
        If Ans = vbNo Then
            formAction = ""
            Exit Sub
        End If
        
        Unload Me
        
    ElseIf Index = 6 Then  'SAVE
        Dim ChckDate As String
        
        If ValidateEntry() <> "" Then
            MsgBox ValidateEntry, vbExclamation, Me.Caption
            Exit Sub
        End If
        
'        DoEvents
'        If DisabledTestConnection = False Then
'            If modConnection.TestConnection() = False Then Exit Sub
'        End If
        
            Screen.MousePointer = vbHourglass
            
            'filter the discount value
            If flxDiscount.TextMatrix(1, 2) <> "" And flxDiscount.TextMatrix(1, 0) = "[X]" Then
                disvalue1 = Trim(flxDiscount.TextMatrix(1, 2))
            End If
            If flxDiscount.TextMatrix(2, 2) <> "" And flxDiscount.TextMatrix(2, 0) = "[X]" Then
                disvalue2 = Trim(flxDiscount.TextMatrix(2, 2))
            End If
            If flxDiscount.TextMatrix(3, 2) <> "" And flxDiscount.TextMatrix(3, 0) = "[X]" Then
                disvalue3 = Trim(flxDiscount.TextMatrix(3, 2))
            End If
            If flxDiscount.TextMatrix(4, 2) <> "" And flxDiscount.TextMatrix(4, 0) = "[X]" Then
                disvalue4 = Trim(flxDiscount.TextMatrix(4, 2))
            End If
            If flxDiscount.TextMatrix(5, 2) <> "" And flxDiscount.TextMatrix(5, 0) = "[X]" Then
                disvalue5 = Trim(flxDiscount.TextMatrix(5, 2))
            End If
            
                        'filter the discount type
            If flxDiscount.TextMatrix(1, 3) <> "" And flxDiscount.TextMatrix(1, 0) = "[X]" Then
                DisType1 = Trim(flxDiscount.TextMatrix(1, 3))
            End If
            If flxDiscount.TextMatrix(2, 3) <> "" And flxDiscount.TextMatrix(2, 0) = "[X]" Then
                DisType2 = Trim(flxDiscount.TextMatrix(2, 3))
            End If
            If flxDiscount.TextMatrix(3, 3) <> "" And flxDiscount.TextMatrix(3, 0) = "[X]" Then
                DisType3 = Trim(flxDiscount.TextMatrix(3, 3))
            End If
            If flxDiscount.TextMatrix(4, 3) <> "" And flxDiscount.TextMatrix(4, 0) = "[X]" Then
                DisType4 = Trim(flxDiscount.TextMatrix(4, 3))
            End If
            If flxDiscount.TextMatrix(5, 3) <> "" And flxDiscount.TextMatrix(5, 0) = "[X]" Then
                DisType5 = Trim(flxDiscount.TextMatrix(5, 3))
            End If
 
        Call CreatTempTable
        
        DoEvents
        If formAction = "EDIT" Then
        
            Ans = MsgBox("DO YOU WANT TO SAVE YOUR CHANGES?", vbQuestion + vbYesNo, Me.Caption)
            If Ans = vbNo Then Exit Sub

            
            Call modAPI.sp_acctpayinvcng_update(Trim(flxAPI.TextMatrix(1, 15)), txtSupName.Text, txtPckngNo.Text, txtTerms.Text, Format(txtDueDate.Text, "yyyy-mm-dd") _
                        , Trim(flxDiscount.TextMatrix(1, 1)), disvalue1, DisType1, Trim(flxDiscount.TextMatrix(2, 1)), disvalue2, DisType2 _
                        , Trim(flxDiscount.TextMatrix(3, 1)), disvalue3, DisType3, Trim(flxDiscount.TextMatrix(4, 1)), disvalue4, DisType4 _
                        , Trim(flxDiscount.TextMatrix(5, 1)), disvalue5, DisType5, Trim(flxDiscount.TextMatrix(6, 2)), txtTotVat.Text _
                        , txtPrepBy.Text, txtAppBy.Text, cboInvType.Text, cboStat.Text, txtNotes.Text, txtLessINVPer1.Text, txtLessINVPer2.Text _
                        , txtLessINVPer3.Text, Val(Replace(txtLessINV.Text, ",", "")), Val(Replace(txtOtherAmnt.Text, ",", "")), Val(Replace(txtTotPayINV.Text, ",", "")) _
                        , txtRecBy.Text, Format(txtDateRecv.Text, "yyyy-mm-dd"), txtChckNo.Text, txtVcherNo.Text, txtChckDate.Text, txtDisNote.Text, Level, message)
            
        ElseIf formAction = "NEW" Then
        
            Ans = MsgBox("PLEASE MAKE SURE THAT THE DATA IS CORRECT!" & _
                vbCrLf & _
                vbCrLf & "DO YOU WISH TO CONTINUE?", vbYesNo + vbQuestion, Me.Caption)
            If Ans = vbNo Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
    
            
            If txtChckDate.Mask = "## / ## / ####" And txtChckDate.Text = "__ / __ / ____" Then
                ChckDate = ""
            Else
                ChckDate = txtChckDate.Text
            End If

            Call modAPI.sp_acctpayinvcng_create(Trim(flxAPI.TextMatrix(1, 15)), txtPckngNo.Text, txtTerms.Text, Format(txtDueDate.Text, "yyyy-mm-dd") _
                        , Trim(flxDiscount.TextMatrix(1, 1)), disvalue1, DisType1, Trim(flxDiscount.TextMatrix(2, 1)), disvalue2, DisType2 _
                        , Trim(flxDiscount.TextMatrix(3, 1)), disvalue3, DisType3, Trim(flxDiscount.TextMatrix(4, 1)), disvalue4, DisType4 _
                        , Trim(flxDiscount.TextMatrix(5, 1)), disvalue5, DisType5, Trim(flxDiscount.TextMatrix(6, 2)), txtTotVat.Text, txtPrepBy.Text _
                        , txtAppBy.Text, cboInvType.Text, cboStat.Text, txtNotes.Text, txtLessINVPer1.Text, txtLessINVPer2.Text, txtLessINVPer3.Text _
                        , Val(Replace(txtLessINV.Text, ",", "")), Val(Replace(txtOtherAmnt.Text, ",", "")), Val(Replace(txtTotPayINV.Text, ",", "")) _
                        , txtRecBy.Text, Format(txtDateRecv.Text, "yyyy-mm-dd"), txtChckNo.Text, txtVcherNo.Text, ChckDate, txtDisNote.Text, Level, message)
            
       
       End If
       
        
        Screen.MousePointer = vbDefault
        
        If Trim(Level) = "" Then
            Exit Sub
        ElseIf Level = "Granted" Or Level = "Warning" Then
            If Level = "Granted" Then
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
        
        cboStat.Locked = True
        cboStat.BackColor = txtHoldGrayBackColor.BackColor
        
        cboInvType.Locked = True
        cboInvType.BackColor = txtHoldGrayBackColor.BackColor
        
        txtChckDate.Enabled = False
        txtChckDate.BackColor = txtHoldGrayBackColor.BackColor
        txtChckDate.Format = "mm / dd / yyyy"
        
        Call SetButtonVisibility(False)
        
        Call EnabledSelectedButton
        
        didRecordEntry = False

        formAction = ""
        
    ElseIf Index = 7 Then 'CANCEL Button
        Ans = MsgBox("Cancel the currrent operation?", vbYesNo + vbQuestion, Me.Caption)
        If Ans = vbNo Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        Call LockedEntryObject(False)
        Call SetButtonVisibility(False)
        Call ClearEntryForm(True)
        Call flxDscntRowValue
        
        cmdTrans(1).Enabled = False
        cmdTrans(2).Enabled = False
        
        cboStat.Locked = True
        cboStat.BackColor = txtHoldGrayBackColor.BackColor
        
        cboInvType.Locked = True
        cboInvType.BackColor = txtHoldGrayBackColor.BackColor
        
        txtChckDate.Enabled = False
        txtChckDate.BackColor = txtHoldGrayBackColor.BackColor
        txtChckDate.Format = "mm / dd / yyyy"
        
        didRecordEntry = False

        formAction = ""

    End If

Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error ->cmdTrans_Click()")
End Sub

Sub flxAPIColsChanger()
On Error Resume Next
    With flxAPI
        flxTTLCol = 16
        .Clear
        .Rows = 2
        .Cols = flxTTLCol
        
        '               0       1             2            3          4            5           6              7      8       9         10            11          12             13          14            15
        .FormatString = " |<PARTNUMBER |<DESCRIPTION |^WHS QTY |^BRANCH QTY |^ORDER QTY. |^INV. PRICE |^DISCOUNT |^VAT |^NEW COST |^OLD COST |^PRICE DIFF.|^% |^INV TOTAL " 'SPAREPARTS ID  SPRR NO
        
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.1   'PART NUMBER
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.16   'DESCRIPTION
        
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.06  'WHS QTY
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.06  'BRANCH QTY
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.06  'ORDER QTY
        
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.07  'INV. PRICE
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.06  'DISCOUNT
        .ColWidth(8) = (.Width - .ColWidth(0)) * 0.06  'VAT
        .ColWidth(9) = (.Width - .ColWidth(0)) * 0.07  'NEW COST
        .ColWidth(10) = (.Width - .ColWidth(0)) * 0.07 'OLD COST
        .ColWidth(11) = (.Width - .ColWidth(0)) * 0.06 'PRICE DIFF.
        .ColWidth(12) = (.Width - .ColWidth(0)) * 0.06 'PRICE DIFF.%
        .ColWidth(13) = (.Width - .ColWidth(0)) * 0.08 'INV TOTAL
        .ColWidth(14) = 0 'SPAREPARTS ID
        .ColWidth(15) = 0 'SPRR NO
        
    End With
    
End Sub

Sub flxDscntColsChanger()
On Error Resume Next
    With flxDiscount
        .Clear
        .Rows = 7
        .Cols = 4

        .FormatString = "^INVOICE DISCOUNT |<INVOICE DISCOUNT |^VALUE  |^DISCOUNT TYPE"

        .ColWidth(0) = 400
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.35    'INVOICE DISCOUN
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.15    'VALUE
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.47      'DISCOUNT TYPE

        'Merge Rows
        .TextMatrix(0, 0) = "INVOICE DISCOUNT"
        .TextMatrix(0, 1) = "INVOICE DISCOUNT"

        .MergeCells = 1
        .MergeRow(0) = True
        .Height = .RowHeightMin * 8
    End With
    
    Call flxDscntRowValue
    
End Sub

Sub flxDscntRowValue()

    flxDiscount.TextMatrix(1, 0) = "[  ]"
    flxDiscount.TextMatrix(2, 0) = "[  ]"
    flxDiscount.TextMatrix(3, 0) = "[  ]"
    flxDiscount.TextMatrix(4, 0) = "[  ]"
    flxDiscount.TextMatrix(5, 0) = "[  ]"
    
    flxDiscount.TextMatrix(1, 1) = "DISCOUNT 1"
    flxDiscount.TextMatrix(2, 1) = "DISCOUNT 2"
    flxDiscount.TextMatrix(3, 1) = "DISCOUNT 3"
    flxDiscount.TextMatrix(4, 1) = "DISCOUNT 4"
    flxDiscount.TextMatrix(5, 1) = "DISCOUNT 5"
    flxDiscount.TextMatrix(6, 1) = "VAT"
    flxDiscount.TextMatrix(6, 2) = "0.00%"
    
End Sub

Sub ClearEntryForm(Focus1stObject As Boolean)
    
    txtSupName.Text = ""
    txtDlvrdBy.Text = ""
    txtWayBillNo.Text = ""
    txtPONo.Text = ""
    
    txtTerms.Text = ""
    txtDueDate.Text = ""
    txtPckngDate.Text = ""
    txtGrossTotINV.Text = ""
    txtLessINV.Text = "0.00"
    txtLessINVPer1.Text = "0%"
    txtLessINVPer2.Text = "0%"
    txtLessINVPer3.Text = "0%"
    txtTotPayINV.Text = "0.00"
    If cboStat.ListIndex >= 0 Then cboStat.ListIndex = 0
    If cboInvType.ListIndex >= 0 Then cboInvType.ListIndex = 0
    txtNotes.Text = ""
    txtDisNote.Text = ""
    txtTotVat.Text = "0.00"
    txtRecBy.Text = ""
    txtDateRecv.Text = ""
    txtPrepBy.Text = ""
    txtAppBy.Text = "LORENZO C. RUBINOS JR."
    txtOtherAmnt.Text = ""
    
    txtChckDate.Mask = ""
    txtChckDate.Text = ""
    txtChckDate.Mask = "## / ## / ####"
    
    txtChckNo.Text = ""
    txtVcherNo.Text = ""
    
    flxAPI.Rows = 1
    flxAPI.Rows = 2
    
    DtlsRow = 0
    For DtlsRow = 1 To flxDiscount.Rows - 2
        flxDiscount.TextMatrix(DtlsRow, 2) = ""
        flxDiscount.TextMatrix(DtlsRow, 3) = ""
    Next
    
    If formAction <> "REFRESH" Then
        txtPckngNo.Text = ""
        
        If didRecordEntry = False Then
            If Focus1stObject = True Then cmdTrans(0).SetFocus
        Else
            If Focus1stObject = True Then txtPckngNo.SetFocus
        End If
    End If
                              
End Sub

Private Sub flxAPI_Click()
    If Trim(flxAPI.TextMatrix(1, 1)) <> "" Then
        flxRowNo = flxAPI.row
    Else
        flxRowNo = 0
    End If
End Sub

Private Sub flxAPI_DblClick()
On Error GoTo er

    If didRecordEntry = False Then Exit Sub
   
    DtlsCol = flxAPI.Col
    If DtlsCol = 6 Then ' INV. PRICE
       With txtFlxInvPrice
        .Top = fraHoldBottom.Top + flxAPI.CellTop + flxAPI.Top
        .Left = fraHoldBottom.Left + flxAPI.CellLeft + flxAPI.Left
        .Width = flxAPI.CellWidth
        .Height = flxAPI.CellHeight
        .Text = Val(Replace(Trim(flxAPI.TextMatrix(flxAPI.row, 6)), ",", ""))  'INV. PRICE
        .Visible = True
        .BorderStyle = 1
        .SelStart = 0
        .SelLength = Len(.Text)
        .ZOrder
        .SetFocus
      End With

        'txtFlxInvPrice.Visible = False

    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error ->flxAPI_DblClick()")
End Sub

Private Sub flxAPI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxAPI_DblClick
End Sub

Private Sub Computations()
On Error GoTo er

    Dim Value1 As Double
    Dim Value2 As Double
    Dim Value3 As Double
    Dim Value4 As Double
    Dim Value5 As Double
    Dim DisVat As Double
    
    Dim invprice As Double
    Dim TotalDis As Double
    Dim NewCost As Double
    
    Dim disvalue1 As Double
    Dim disvalue2 As Double
    Dim disvalue3 As Double
    Dim disvalue4 As Double
    Dim disvalue5 As Double
    
    Dim Discount1 As Double
    Dim Discount2 As Double
    Dim Discount3 As Double
    Dim Discount4 As Double
    Dim Discount5 As Double
    Dim VatValue As Double
        
      'Convert percent to decimal in discount value
        If flxDiscount.TextMatrix(1, 2) <> "" And flxDiscount.TextMatrix(1, 0) = "[X]" Then
            Value1 = Left(Trim(flxDiscount.TextMatrix(1, 2)), Len(Trim(flxDiscount.TextMatrix(1, 2))) - 1)
            Value1 = Value1 / 100
        End If
        If flxDiscount.TextMatrix(2, 2) <> "" And flxDiscount.TextMatrix(2, 0) = "[X]" Then
            Value2 = Left(Trim(flxDiscount.TextMatrix(2, 2)), Len(Trim(flxDiscount.TextMatrix(2, 2))) - 1)
            Value2 = Value2 / 100
        End If
        If flxDiscount.TextMatrix(3, 2) <> "" And flxDiscount.TextMatrix(3, 0) = "[X]" Then
            Value3 = Left(Trim(flxDiscount.TextMatrix(3, 2)), Len(Trim(flxDiscount.TextMatrix(3, 2))) - 1)
            Value3 = Value3 / 100
        End If
        If flxDiscount.TextMatrix(4, 2) <> "" And flxDiscount.TextMatrix(4, 0) = "[X]" Then
            Value4 = Left(Trim(flxDiscount.TextMatrix(4, 2)), Len(Trim(flxDiscount.TextMatrix(4, 2))) - 1)
            Value4 = Value4 / 100
        End If
        If flxDiscount.TextMatrix(5, 2) <> "" And flxDiscount.TextMatrix(5, 0) = "[X]" Then
            Value5 = Left(Trim(flxDiscount.TextMatrix(5, 2)), Len(Trim(flxDiscount.TextMatrix(5, 2))) - 1)
            Value5 = Value5 / 100
        End If
        
    If cboInvType.Text = "" Then
        MsgBox "Please select invoice type.", vbInformation, Me.Caption
        cboInvType.SetFocus
        Exit Sub
    End If

    'Loop for Computation
    DtlsRow = 0
    For DtlsRow = 1 To flxAPI.Rows - 1
        
        invprice = 0
        Discount1 = 0
        Discount2 = 0
        Discount3 = 0
        Discount4 = 0
        Discount5 = 0

        invprice = Val(Replace(Trim(flxAPI.TextMatrix(DtlsRow, 6)), ",", ""))
        
        
        
        If cboInvType.Text = "VAT INCLUSIVE" Then
        
            DisVat = 1.12
            NewCost = invprice / DisVat
            VatValue = invprice - NewCost
            
        ElseIf cboInvType.Text = "VAT EXCLUSIVE" Then
        
            VatValue = 0
            NewCost = invprice

        End If

        If Value1 <> 0 Then 'Computation for Invoice Discount 1
            Discount1 = NewCost * Value1
            disvalue1 = NewCost - Discount1
            'VatValue = VatValue - Discount1

        End If

        If Value2 <> 0 Then 'Computation for Invoice Discount 2
            If Discount1 = 0 Then
                Discount2 = NewCost * Value2
                disvalue2 = NewCost - Discount2
                'VatValue = DisVat - Discount2
            Else
                Discount2 = disvalue1 * Value2
                disvalue2 = disvalue1 - Discount2
                'VatValue = DisVat - Discount2
            End If

        End If

        If Value3 <> 0 Then 'Computation for Invoice Discount 3
            If Discount2 = 0 Then
                Discount3 = NewCost * Value3
                disvalue3 = NewCost - Discount3
                'VatValue = DisVat - Discount3
            Else
                Discount3 = Discount2 * Value3
                disvalue3 = disvalue2 - Discount3
                'VatValue = DisVat - Discount3
            End If

        End If

        If Value4 <> 0 Then 'Computation for Invoice Discount 4
            If Discount3 = 0 Then
                Discount4 = NewCost * Value4
                disvalue4 = NewCost - Discount4
                'VatValue = DisVat - Discount4
            Else
                Discount4 = Discount3 * Value4
                disvalue4 = disvalue3 - Discount4
                'VatValue = DisVat - Discount4
            End If

        End If

        If Value5 <> 0 Then 'Computation for Invoice Discount 5
            If Discount4 = 0 Then
                Discount5 = NewCost * Value5
                disvalue5 = NewCost - Discount5
                'VatValue = DisVat - Discount5
            Else
                Discount5 = Discount4 * Value5
                disvalue5 = disvalue4 - Discount5
                'VatValue = DisVat - Discount5
            End If

        End If

        'Sum of all discount
        TotalDis = Discount1 + Discount2 + Discount3 + Discount4 + Discount5

        'Computation of new cost

        NewCost = NewCost - TotalDis

        flxAPI.TextMatrix(DtlsRow, 7) = Format(TotalDis, "#,##0.00")

        If VatValue < 0 Then
            VatValue = 0
        End If
        flxAPI.TextMatrix(DtlsRow, 8) = Format(VatValue, "#,##0.00")

        flxAPI.TextMatrix(DtlsRow, 9) = Format(NewCost, "#,##0.00")

    Next

    DoEvents
    Call Total

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error ->Computations()")
End Sub

Sub Total()
On Error GoTo er
    Dim qty As Integer
    Dim invtotal As Double
    Dim invprice As Double
    Dim NewCost As Double
    Dim CostDiff As Double
    Dim CostDiffPer As Double
    Dim MrcCost As Double

'    Dim decvalue1 As Double
'    Dim decvalue2 As Double
'    Dim decvalue3 As Double
'    Dim decvalue4 As Double
'    Dim decvalue5 As Double
'
'    Dim disvalue1 As Double
'    Dim disvalue2 As Double
'    Dim disvalue3 As Double
'    Dim disvalue4 As Double
'    Dim disvalue5 As Double
'
'    Dim disinvprice1 As Double
'    Dim disinvprice2 As Double
'    Dim disinvprice3 As Double
'    Dim disinvprice4 As Double
'    Dim disinvprice5 As Double

'       1             2            3          4            5           6             7      8       9         10            11      12      13          14            15
'|<PART NUMBER |<DESCRIPTION |^WHS QTY |^BRANCH QTY |^ORDER QTY. |^INV. PRICE |^DISCOUNT |^VAT |^NEW COST |^OLD COST |^PRICE DIFF.|^% |^INV TOTAL " 'SPAREPARTS ID  SPRR NO

    DtlsRow = 0
    For DtlsRow = 1 To flxAPI.Rows - 1

        qty = Trim(flxAPI.TextMatrix(DtlsRow, 5))

        invprice = Val(Replace(Trim(flxAPI.TextMatrix(DtlsRow, 6)), ",", ""))

        'vat = Val(Replace(Trim(flxAPI.TextMatrix(DtlsRow, 9)), ",", ""))
        'invprice = txtFlxInvPrice.Text

        'Computation for INV TOTAL
'        If flxDiscount.TextMatrix(1, 2) <> "" And flxDiscount.TextMatrix(1, 0) = "[X]" Or flxDiscount.TextMatrix(1, 2) = "0.00%" Then   'Get all discount value with discount
'            decvalue1 = Left(Trim(flxDiscount.TextMatrix(1, 2)), Len((flxDiscount.TextMatrix(1, 2))) - 1)
'            decvalue1 = decvalue1 / 100 'Convert to decimal
'
'            disvalue1 = invprice * decvalue1
'            disinvprice1 = invprice - disvalue1
'            invtotal = qty * disinvprice1
'
'        End If
'
'        If flxDiscount.TextMatrix(2, 2) <> "" And flxDiscount.TextMatrix(2, 0) = "[X]" Or flxDiscount.TextMatrix(2, 2) = "0.00%" Then 'Get all discount value with discount
'            decvalue2 = Left(Trim(flxDiscount.TextMatrix(2, 2)), Len((flxDiscount.TextMatrix(2, 2))) - 1)
'            decvalue2 = decvalue2 / 100 'Convert to decimal
'
'            disvalue2 = disinvprice1 * decvalue2
'            disinvprice2 = disinvprice1 - disvalue2
'            invtotal = qty * disinvprice2
'
'        End If
'
'        If flxDiscount.TextMatrix(3, 2) <> "" And flxDiscount.TextMatrix(3, 0) = "[X]" Or flxDiscount.TextMatrix(3, 2) = "0.00%" Then 'Get all discount value with discount
'            decvalue3 = Left(Trim(flxDiscount.TextMatrix(3, 2)), Len((flxDiscount.TextMatrix(3, 2))) - 1)
'            decvalue3 = decvalue3 / 100 'Convert to decimal
'
'            disvalue3 = disinvprice2 * decvalue3
'            disinvprice3 = disinvprice2 - disvalue3
'            invtotal = qty * disinvprice3
'
'        End If
'
'        If flxDiscount.TextMatrix(4, 2) <> "" And flxDiscount.TextMatrix(4, 0) = "[X]" Or flxDiscount.TextMatrix(4, 2) = "0.00%" Then  'Get all discount value with discount
'            decvalue4 = Left(Trim(flxDiscount.TextMatrix(4, 2)), Len((flxDiscount.TextMatrix(4, 2))) - 1)
'            decvalue4 = decvalue4 / 100 'Convert to decimal
'
'            disvalue4 = disinvprice2 * decvalue4
'            disinvprice4 = disvalue3 - disvalue4
'            invtotal = qty * disinvprice4
'
'        End If
'
'        If flxDiscount.TextMatrix(5, 2) <> "" And flxDiscount.TextMatrix(5, 0) = "[X]" Or flxDiscount.TextMatrix(5, 2) = "0.00%" Then 'Get all discount value with discount
'            decvalue5 = Left(Trim(flxDiscount.TextMatrix(5, 2)), Len((flxDiscount.TextMatrix(5, 2))) - 1)
'            decvalue5 = decvalue5 / 100 'Convert to decimal
'
'            disvalue5 = disinvprice4 * decvalue5
'            disinvprice5 = disvalue4 - disvalue5
'            invtotal = qty * disinvprice5
'
'        End If

        
        invtotal = invprice * qty
        
        
        flxAPI.TextMatrix(DtlsRow, 13) = Format(invtotal, "#,##0.00")
        
        NewCost = Val(Replace(Trim(flxAPI.TextMatrix(DtlsRow, 9)), ",", ""))
        
        MrcCost = Val(Replace(Trim(flxAPI.TextMatrix(DtlsRow, 10)), ",", ""))
        
        If MrcCost = 0 Then
            MrcCost = 1
        End If
        
        CostDiff = NewCost - MrcCost

        CostDiffPer = (CostDiff / MrcCost) * 100
        
        flxAPI.TextMatrix(DtlsRow, 11) = Format(CostDiff, "#,##0.00")
        flxAPI.TextMatrix(DtlsRow, 12) = Format(CostDiffPer, "##0.00") & "%"
        
    Next
        
        DoEvents
        Call TotalGross
        DoEvents
        Call TotalPay
        
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error ->Total()")
End Sub


Sub TotalGross()
    Dim grosstotinv As Double
    Dim invtotal As Double
    Dim i As Integer
    
    For i = 1 To flxAPI.Rows - 1
    
        invtotal = Val(Replace(Trim(flxAPI.TextMatrix(i, 13)), ",", ""))
        
        grosstotinv = grosstotinv + invtotal
    
    Next
    
        txtGrossTotINV.Text = Format(grosstotinv, "#,##0.00")
        
End Sub

Private Sub flxAPI_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    'If didRecordEntry = False Then Exit Sub
    If Button = vbRightButton Then
        PopupMenu mnuPopUpFormat, vbPopupMenuRightButton
    End If
End Sub


Private Sub flxDiscount_DblClick()
On Error GoTo er
    
    If didRecordEntry = False Then Exit Sub
    
    '                          0                 1             2           3
    '.FormatString = "<INVOICE DISCOUNT |<INVOICE DISCOUNT |^VALUE  |^DISCOUNT TYPE"
    DtlsCol = flxDiscount.Col
    
    If DtlsCol = 0 Then 'INVOICE DISCOUNT
    
        
        If flxDiscount.TextMatrix(flxDiscount.row, 2) = "" Or flxDiscount.TextMatrix(flxDiscount.row, 0) = "" Then Exit Sub
    
        If flxDiscount.TextMatrix(flxDiscount.row, 0) = "[  ]" Then
            flxDiscount.TextMatrix(flxDiscount.row, 0) = "[X]"
            
            DoEvents
            Call Computations
            
            DoEvents
            Call DisNoteValue
            
        Else
            flxDiscount.TextMatrix(flxDiscount.row, 0) = "[  ]"
            
            DoEvents
            Call Computations
            
            txtDisNote.Text = ""
            DoEvents
            Call DisNoteValue
            
        End If
    
    ElseIf DtlsCol = 2 Then 'VALUE
    
       If flxDiscount.row = 6 Then
            With txtFlxVat
                .Top = fraHoldBottom.Top + flxDiscount.CellTop + flxDiscount.Top
                .Left = fraHoldBottom.Left + flxDiscount.CellLeft + flxDiscount.Left
                .Width = flxDiscount.CellWidth
                .Height = flxDiscount.CellHeight
                .Visible = True
                
                If flxDiscount.TextMatrix(flxDiscount.row, 2) <> "" Then
                    .Text = Left(Trim(flxDiscount.TextMatrix(flxDiscount.row, 2)), Len((flxDiscount.TextMatrix(flxDiscount.row, 2))) - 1)
                Else
                    .Text = ""
                End If
                
                .BorderStyle = 1
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
              End With
            Exit Sub
       Else
       
             With txtFlxValue
              .Top = fraHoldBottom.Top + flxDiscount.CellTop + flxDiscount.Top
              .Left = fraHoldBottom.Left + flxDiscount.CellLeft + flxDiscount.Left
              .Width = flxDiscount.CellWidth
              .Height = flxDiscount.CellHeight
              .Visible = True
              
              If flxDiscount.TextMatrix(flxDiscount.row, 2) <> "" Then
                  .Text = Left(Trim(flxDiscount.TextMatrix(flxDiscount.row, 2)), Len((flxDiscount.TextMatrix(flxDiscount.row, 2))) - 1)
              Else
                  .Text = ""
              End If
              
              .BorderStyle = 1
              .SelStart = 0
              .SelLength = Len(.Text)
              .ZOrder
              .SetFocus
            End With
      
       End If
       

        
    ElseIf DtlsCol = 3 Then 'DISCOUNT TYPE
    
       If flxDiscount.row = 6 Then Exit Sub
       
       With txtFlxDisType
        .Top = fraHoldBottom.Top + flxDiscount.CellTop + flxDiscount.Top
        .Left = fraHoldBottom.Left + flxDiscount.CellLeft + flxDiscount.Left
        .Width = flxDiscount.CellWidth
        .Height = flxDiscount.CellHeight
        .Text = flxDiscount.TextMatrix(flxDiscount.row, 3)  'DISCOUNT TYPE
        .Visible = True
        .BorderStyle = 1
        .SelLength = Len(.Text)
        .SelStart = 0
        .SelLength = Len(txtFlxDisType.Text)
        .ZOrder
        .SetFocus
      End With
      
    Else
    
    txtFlxInvPrice.Visible = False

    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error ->flxDiscount_DblClick()")
End Sub



Private Sub flxDiscount_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxDiscount_DblClick
End Sub


Private Sub Form_Load()
On Error GoTo er

     DisabledTestConnection = True
     
     cboStat.ListIndex = 0
        
     txtPrepBy.Text = EmployeesName
     
     flxColNo = 1
     
     Me.Caption = "SUPPLIER PRODUCT COST RECEIVING RECEIPT"
     Warehouse.sBar.Panels(6).Text = Me.Caption
     
Exit Sub
er:
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->Form_Load()")
End Sub

Private Sub Form_Resize()
    Call RXResizer1_ResizeComplete
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo er
    Dim Ans As VbMsgBoxResult
    
    If formAction = "" Then
    
        Ans = MsgBox("EXIT APPLICATION?", vbQuestion + vbYesNo, Me.Caption)
        If Ans = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    
    End If
    
    formAction = ""
    
Exit Sub
er:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub lblClose_Click()
    SSTab1.Visible = False
    txtFind.Text = ""
End Sub

Private Sub mnuSTFindRow_Click()
    SSTab1.Visible = True
    txtFind.SetFocus
End Sub

Private Sub RXResizer1_ResizeComplete()
'On Error Resume Next
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->PRINT
'3 ->SUMMRY
'4 ->REFRESH
'5 ->EXIT
'6 ->SAVE
'7 ->CANCEL

    'Align Save Button
    cmdTrans(6).Left = cmdTrans(0).Left
    cmdTrans(6).Top = cmdTrans(0).Top

    'Align Cancel Button
    cmdTrans(7).Left = Me.Width - 1600
    cmdTrans(7).Top = cmdTrans(0).Top
    
    fraHoldTop.Width = Me.Width - 510
    fraHoldBottom.Width = Me.Width - 510
    
    If txtSupName.Text = "" Then
        DoEvents
        Call flxAPIColsChanger
        DoEvents
        Call flxDscntColsChanger
    End If
    
End Sub




Private Sub tmrGetRecord_Timer()
On Error GoTo er
    If tmrGetRecord.Enabled = False Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    tmrGetRecord.Enabled = False
    
    DoEvents
    Call ShowRecord(txtPckngNo.Text)
    
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

Private Sub tmrRefresh_Timer()
On Error GoTo er
    If tmrRefresh.Enabled = False Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    tmrRefresh.Enabled = False
    
    DoEvents
    Call RefreshRecord(txtPckngNo.Text)
    
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    
    Screen.MousePointer = Default
    
Exit Sub
er:
    tmrRefresh.Enabled = False
    Screen.MousePointer = Default
End Sub


Private Sub txtChckDate_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtVcherNo, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then txtNotes.SetFocus
End Sub

Private Sub txtChckNo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtChckNo, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtVcherNo.SetFocus
End Sub

Private Sub txtDisNote_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub

Private Sub txtDueDate_KeyPress(KeyAscii As Integer)
    If didRecordEntry = False Then
        KeyAscii = 0
        Beep!
        Exit Sub
    Else
        KeyAscii = Asc(UCase$(Chr(KeyAscii)))
        Call ObjKeyTrapping.Validation(txtDueDate, KeyAscii, NumbersOnly)
    End If
    
    If KeyAscii = 13 Then
    
        cboInvType.SetFocus
    
    End If
End Sub


Private Sub cboStat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNotes.SetFocus
End Sub


Private Sub txtFind_Change()
On Error GoTo er
    Dim Qry As String
    
    If PrtSrch Then
        PrtSrch = False
        Exit Sub
    End If
    
'    If tmpRs.State = adStateOpen Then tmpRs.Close
'    tmpRs.Open
'
'    If txtFind.Text = "" Then Exit Sub
'    tmpRs.Find "[PARTNUMBER] LIKE '" + txtFind.Text + "%'", , adSearchForward, 1
'    If Not tmpRs.EOF Then
'        If Not PrtSrch Then
'            txtTag.Text = Len(txtFind.Text)
'            PrtSrch = True
'        End If
'        txtFind.Text = tmpRs!PARTNUMBER
'        txtFind.SelStart = Val(txtTag.Text)
'        txtFind.SelLength = Len(txtFind.Text)
'     End If

    If tmpRs.State = adStateOpen Then tmpRs.Close
    
    Qry = "SELECT DISTINCT(PARTNUMBER) FROM mrcwarehouse.spareparts WHERE PARTNUMBER LIKE '" + txtFind.Text + "%';"
    'InputBox Qry, Qry, Qry
    tmpRs.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic

    If tmpRs.RecordCount > 0 Then
        If Not PrtSrch Then
            txtTag.Text = Len(txtFind.Text)
            PrtSrch = True
        End If
        
        txtFind.Text = tmpRs!PARTNUMBER
        txtFind.SelStart = Val(txtTag.Text)
        txtFind.SelLength = Len(txtFind.Text)
    End If
    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtFind_Change()")
End Sub

Private Sub txtFind_DblClick()
    'Highlight the txtLessINVPer1
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
    txtFind.SetFocus
End Sub


Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Or KeyCode = 8 Then
        PrtSrch = True
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub txtFlxDisType_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))

    If KeyAscii = 13 Then
    
        Call DisNoteValue
        
        flxDiscount.TextMatrix(flxDiscount.row, 3) = txtFlxDisType.Text
        txtFlxDisType.Visible = False
        
        flxDiscount.SetFocus
        flxDiscount.Col = 2
        If flxDiscount.row <> flxDiscount.Rows - 1 Then
            flxDiscount.row = flxDiscount.row + 1
        End If
        
    End If
End Sub

Private Sub txtFlxDisType_LostFocus()
    txtFlxDisType.Visible = False
End Sub

Private Sub txtFlxInvPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFlxInvPrice, KeyAscii, NumbersWithDecimalPoint)
    If KeyAscii = 13 Then
        flxAPI.TextMatrix(flxAPI.row, 6) = Format(txtFlxInvPrice.Text, "#,##0.0000")
        txtFlxInvPrice.Visible = False
        
        flxAPI.SetFocus
        flxAPI.Col = 6
        If flxAPI.row <> flxAPI.Rows - 1 Then
            flxAPI.row = flxAPI.row + 1
        End If
        
    End If
End Sub

Private Sub txtFlxInvPrice_LostFocus()
    txtFlxInvPrice.Visible = False
End Sub

'Private Sub txtFlxRemarks_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
'
'    If KeyAscii = 13 Then
'        flxAPI.TextMatrix(flxAPI.Row, flxAPI.Col) = txtFlxRemarks.Text
'        flxAPI.SetFocus
'        flxAPI.Col = flxAPI.Col
'        txtFlxRemarks.Visible = False
'    End If
'End Sub


Private Sub txtFlxValue_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFlxValue, KeyAscii, NumbersWithDecimalPoint)
    
    
    If KeyAscii = 13 Then
    
        If txtFlxValue.Text = "" Then
            txtFlxValue.Text = ""
            txtFlxValue.Visible = False
            Exit Sub
        End If
    
        If txtFlxValue.Text > 100 Then
            txtFlxValue.Text = ""
            txtFlxValue.Visible = False
            flxDiscount.SetFocus
            flxDiscount.Col = 2
            Exit Sub
        End If
        
        If txtFlxValue.Text < 1 Then
            flxDiscount.TextMatrix(flxDiscount.row, 2) = Format(txtFlxValue.Text, "#,##0.00") & "%"
            'Call Computations
        Else
            flxDiscount.TextMatrix(flxDiscount.row, 2) = Format(txtFlxValue.Text, "#,##0.00") & "%"
            'Call Computations
        End If
        
        If flxDiscount.TextMatrix(flxDiscount.row, 0) = "[X]" Then
            flxDiscount.TextMatrix(flxDiscount.row, 0) = "[  ]"
            
            'Call Computations

        End If
        
            flxDiscount.TextMatrix(flxDiscount.row, 2) = Format(txtFlxValue.Text, "#,##0.00") & "%"
            txtFlxValue.Visible = False
            flxDiscount.SetFocus
            flxDiscount.Col = 3
            
        
    End If
End Sub

Private Sub txtFlxValue_LostFocus()
    txtFlxValue.Visible = False
End Sub

Sub InvCostDis()
    Dim NewCost As Double
    Dim less1 As Double
    Dim less2 As Double
    Dim less3 As Double
    Dim disval As Double
    
    Call Computations
    
    
    DtlsRow = 0
    For DtlsRow = 1 To flxAPI.Rows - 1
    
            NewCost = Val(Replace(Trim(flxAPI.TextMatrix(DtlsRow, 10)), ",", ""))
            less1 = Left(Trim(txtLessINVPer1.Text), Len(txtLessINVPer1.Text) - 1)
            less2 = Left(Trim(txtLessINVPer2.Text), Len(txtLessINVPer2.Text) - 1)
            less3 = Left(Trim(txtLessINVPer3.Text), Len(txtLessINVPer3.Text) - 1)
            
            'Less% 1
                
                less1 = less1 / 100
                
                disval = NewCost * less1
                NewCost = NewCost - disval
                flxAPI.TextMatrix(DtlsRow, 10) = Format(NewCost, "#,##0.00")

            'Less% 2
                
                less2 = less2 / 100
                   
                disval = NewCost * less2
                NewCost = NewCost - disval
                flxAPI.TextMatrix(DtlsRow, 10) = Format(NewCost, "#,##0.00")
    
            'Less% 3
                
                less3 = less3 / 100
                
                disval = NewCost * less3
                NewCost = NewCost - disval
                flxAPI.TextMatrix(DtlsRow, 10) = Format(NewCost, "#,##0.00")
            
    Next
            
End Sub

Private Sub TotalPay()
    Dim grosstotinv As Double
    Dim payvalue As Double
    Dim totpayinv As Double
    Dim lessinvper1 As Double
    Dim lessinvper2 As Double
    Dim lessinvper3 As Double
    Dim lessinv1 As Double
    Dim lessinv2 As Double
    Dim lessinv3 As Double
    Dim totvat As Double
    Dim TotPay As Double

    
        lessinvper1 = Val(Replace(txtLessINVPer1.Text, " , ", ""))
        lessinvper2 = Val(Replace(txtLessINVPer2.Text, " , ", ""))
        lessinvper3 = Val(Replace(txtLessINVPer3.Text, " , ", ""))
  
 '------------------------------------------------------------------'
 
        'Get the total of invoice payable
        grosstotinv = Val(Replace(txtGrossTotINV.Text, ",", ""))
        
        lessinv1 = grosstotinv * (lessinvper1 / 100)
                
        If lessinv1 > grosstotinv Then
            MsgBox "Less is greater than gross total!", vbExclamation
            txtLessINVPer1.SetFocus
            Exit Sub
        Else
                If cboInvType.Text = "VAT INCLUSIVE" Then
                    totpayinv = grosstotinv - lessinv1
                    txtLessINV.Text = Format(grosstotinv - totpayinv, "#,##0.00")
                Else
                    totpayinv = grosstotinv + lessinv1
                    txtLessINV.Text = Format(totpayinv - grosstotinv, "#,##0.00")
                End If
                
                txtTotPayINV.Text = Format(totpayinv, "#,##0.00")
                txtLessINVPer1.Text = lessinvper1 & "%"
        End If
        
        
        If txtLessINVPer2.Text <> "0%" Then
                'Get the total of new cost total payable
                payvalue = Val(Replace(txtTotPayINV.Text, ",", ""))
                lessinv2 = payvalue * (lessinvper2 / 100)
                        
                If lessinv2 > grosstotinv Then
                    MsgBox "Less is greater than gross total!", vbExclamation
                    txtLessINVPer2.SetFocus
                    Exit Sub
                Else
                        If cboInvType.Text = "VAT INCLUSIVE" Then
                            totpayinv = payvalue - lessinv2
                            txtLessINV.Text = Format(grosstotinv - totpayinv, "#,##0.00")
                        Else
                            totpayinv = payvalue + lessinv2
                            txtLessINV.Text = Format(totpayinv - grosstotinv, "#,##0.00")
                        End If
                            
                        txtTotPayINV.Text = Format(totpayinv, "#,##0.00")
                        txtLessINVPer2.Text = lessinvper2 & "%"
                End If
        Else
                txtLessINVPer2.Text = "0%"
        End If
        
        
        If txtLessINVPer3.Text <> "0%" Then
                'Get the total of new cost total payable
                payvalue = Val(Replace(txtTotPayINV.Text, ",", ""))
                lessinv3 = payvalue * (lessinvper3 / 100)
                        
                If lessinv3 > grosstotinv Then
                    MsgBox "Less is greater than gross total!", vbExclamation
                    txtLessINVPer3.SetFocus
                    Exit Sub
                Else
                        If cboInvType.Text = "VAT INCLUSIVE" Then
                            totpayinv = payvalue - lessinv3
                            txtLessINV.Text = Format(grosstotinv - totpayinv, "#,##0.00")
                        Else
                            totpayinv = payvalue + lessinv3
                            txtLessINV.Text = Format(totpayinv - grosstotinv, "#,##0.00")
                        End If
                        
                        txtTotPayINV.Text = Format(totpayinv, "#,##0.00")
                        txtLessINVPer3.Text = lessinvper3 & "%"
                End If
        Else
                txtLessINVPer3.Text = "0%"
        End If
        
        
        
        totvat = Val(Replace(txtTotVat, ",", ""))
        payvalue = Val(Replace(txtTotPayINV, ",", ""))
        
        
        If cboInvType.Text = "VAT INCLUSIVE" Then
             TotPay = payvalue - totvat
        Else
             TotPay = payvalue + totvat
        End If
        
        
        txtTotPayINV.Text = Format(TotPay, "#,##0.00")
        txtOtherAmnt.Text = "0.00"
        
End Sub

Private Sub txtFlxVat_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFlxVat, KeyAscii, NumbersWithDecimalPoint)
    
    Dim Gross As Double
    Dim vat As Double
    Dim totvat As Double
    Dim product As Double
    
    If KeyAscii = 13 Then
    
        DoEvents
        Call Computations
        
        Gross = Val(Replace(Trim(txtGrossTotINV.Text), ",", ""))
        vat = Val(Replace(Trim(txtFlxVat.Text), ",", ""))
        
        vat = vat / 100
        totvat = Gross * vat
        
        
        txtTotVat.Text = Format(totvat, "#,##0.00")
        
        flxDiscount.TextMatrix(flxDiscount.row, 2) = txtFlxVat.Text & "%"
        txtFlxVat.Visible = False
        
        DoEvents
        Call TotalPay
        
    End If
    
End Sub

Private Sub txtFlxVat_LostFocus()
    txtFlxVat.Visible = False
End Sub

Private Sub txtLessINVPer1_Click()
    'Highlight the txtLessINVPer1
    txtLessINVPer1.SelStart = 0
    txtLessINVPer1.SelLength = Len(txtLessINVPer1.Text)
    txtLessINVPer1.SetFocus
End Sub

Private Sub txtLessINVPer1_KeyPress(KeyAscii As Integer)
    Dim percnt As Double
    
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtLessINVPer1, KeyAscii, NumbersWithDecimalPoint)
    
    If KeyAscii = 13 Then
        
        percnt = Val(Replace(txtLessINVPer1.Text, " , ", ""))
        
        If percnt > 100 Then
            MsgBox "Percentage is greater than 100%", vbExclamation, "VALIDATION"
            'Highlight the txtLessINVPer1
            txtLessINVPer1.SelStart = 0
            txtLessINVPer1.SelLength = Len(txtLessINVPer1.Text)
            txtLessINVPer1.SetFocus
            Exit Sub
        End If
        
        'Call InvCostDis
        Call TotalPay
            
            'Highlight the txtLessINVPer1
            txtLessINVPer1.SelStart = 0
            txtLessINVPer1.SelLength = Len(txtLessINVPer1.Text)
            txtLessINVPer1.SetFocus
        
    End If
End Sub

Private Sub txtLessINVPer1_LostFocus()
'        Dim percnt As Double
'
'        percnt = Val(Replace(txtLessINVPer1.Text, " , ", ""))
'
'        If percnt > 100 Then
'            MsgBox "Percentage is greater than 100%", vbExclamation, "VALIDATION"
'            'Highlight the txtLessINVPer1
'            txtLessINVPer1.SelStart = 0
'            txtLessINVPer1.SelLength = Len(txtLessINVPer1.Text)
'            txtLessINVPer1.SetFocus
'            Exit Sub
'        End If
'
'        Call InvCostDis
'        Call TotalPay

        If txtLessINVPer1.Text = "" Then
            txtLessINVPer1.Text = "0%"
            Exit Sub
        End If

End Sub

Private Sub txtLessINVPer2_Click()
    'Highlight the txtLessINVPer2
     txtLessINVPer2.SelStart = 0
     txtLessINVPer2.SelLength = Len(txtLessINVPer2.Text)
     txtLessINVPer2.SetFocus
End Sub

Private Sub txtLessINVPer2_KeyPress(KeyAscii As Integer)
Dim percnt As Double
    
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtLessINVPer2, KeyAscii, NumbersWithDecimalPoint)
    
    If KeyAscii = 13 Then
    
        
        percnt = Val(Replace(txtLessINVPer2.Text, " , ", ""))
        
        If percnt > 100 Then
            MsgBox "Percentage is greater than 100%", vbExclamation, "VALIDATION"
            'Highlight the txtLessINVPer2
            txtLessINVPer2.SelStart = 0
            txtLessINVPer2.SelLength = Len(txtLessINVPer2.Text)
            txtLessINVPer2.SetFocus
            Exit Sub
        End If
        
        'Call InvCostDis
        Call TotalPay
            
            'Highlight the txtLessINVPer2
            txtLessINVPer2.SelStart = 0
            txtLessINVPer2.SelLength = Len(txtLessINVPer2.Text)
            txtLessINVPer2.SetFocus
    End If
End Sub

Private Sub txtLessINVPer2_LostFocus()
'        Dim percnt As Double
'
'        percnt = Val(Replace(txtLessINVPer2.Text, " , ", ""))
'
'        If percnt > 100 Then
'            MsgBox "Percentage is greater than 100%", vbExclamation, "VALIDATION"
'            'Highlight the txtLessINVPer2
'            txtLessINVPer2.SelStart = 0
'            txtLessINVPer2.SelLength = Len(txtLessINVPer2.Text)
'            txtLessINVPer2.SetFocus
'            Exit Sub
'        End If
'
'        Call InvCostDis
'        Call TotalPay

        If txtLessINVPer2.Text = "" Then
            txtLessINVPer2.Text = "0%"
            Exit Sub
        End If
            
End Sub

Private Sub txtLessINVPer3_Click()
    'Highlight the txtLessINVPer2
     txtLessINVPer3.SelStart = 0
     txtLessINVPer3.SelLength = Len(txtLessINVPer3.Text)
     txtLessINVPer3.SetFocus
End Sub

Private Sub txtLessINVPer3_KeyPress(KeyAscii As Integer)
    Dim percnt As Double
    
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtLessINVPer3, KeyAscii, NumbersWithDecimalPoint)
    
    If KeyAscii = 13 Then
    
        percnt = Val(Replace(txtLessINVPer3.Text, ",", ""))
    
        If percnt > 100 Then
            MsgBox "Percentage is greater than 100%", vbExclamation, "VALIDATION"
            'Highlight the txtLessNCPer3
            txtLessINVPer3.SelStart = 0
            txtLessINVPer3.SelLength = Len(txtLessINVPer3.Text)
            txtLessINVPer3.SetFocus
            Exit Sub
        End If
        
        'Call InvCostDis
        Call TotalPay
        
            'Highlight the txtLessNCPer3
            txtLessINVPer3.SelStart = 0
            txtLessINVPer3.SelLength = Len(txtLessINVPer3.Text)
            txtLessINVPer3.SetFocus
        
    End If
End Sub

Private Sub txtLessINVPer3_LostFocus()
'        Dim percnt As Double
'
'        percnt = Val(Replace(txtLessINVPer3.Text, ",", ""))
'
'        If percnt > 100 Then
'            MsgBox "Percentage is greater than 100%", vbExclamation, "VALIDATION"
'            'Highlight the txtLessNCPer3
'            txtLessINVPer3.SelStart = 0
'            txtLessINVPer3.SelLength = Len(txtLessINVPer3.Text)
'            txtLessINVPer3.SetFocus
'            Exit Sub
'        End If

'
'        Call InvCostDis
'        Call TotalPay

        If txtLessINVPer3.Text = "" Then
            txtLessINVPer3.Text = "0%"
            Exit Sub
        End If
        
End Sub

Private Sub txtNotes_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub

Private Sub txtOtherAmnt_Click()
    'Highlight the txtOtherAmnt
    txtOtherAmnt.SelStart = 0
    txtOtherAmnt.SelLength = Len(txtOtherAmnt.Text)
    txtOtherAmnt.SetFocus
End Sub

Private Sub txtOtherAmnt_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtOtherAmnt, KeyAscii, NumbersWithDecimalPoint)
    
    If KeyAscii = 13 Then
        Dim payable As Double
        Dim TotPay As Double
        
        If txtOtherAmnt.Text = "" Then Exit Sub
        
        payable = Val(Replace(txtTotPayINV.Text, ",", ""))
        TotPay = payable + txtOtherAmnt.Text
        
        txtTotPayINV.Text = Format(TotPay, "#,##0.00")
    
    End If
    
End Sub

Private Sub txtOtherAmnt_LostFocus()
        If txtOtherAmnt.Text = "" Then
            txtOtherAmnt.Text = "0.00"
            Exit Sub
        End If
End Sub


Private Sub txtPckngNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        
        Call GetRecord(txtPckngNo.Text)
        
    End If
End Sub

Private Sub txtTerms_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtTerms, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then
        
        If txtTerms.Text = "" Then Exit Sub
        txtDueDate.Text = DateAdd("d", txtTerms.Text, txtPckngDate.Text)
        txtDueDate.SetFocus
    
    End If
End Sub

Private Sub txtVcherNo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtVcherNo, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtChckDate.SetFocus
End Sub
