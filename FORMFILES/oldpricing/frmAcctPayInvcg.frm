VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAcctPayInvcg 
   Caption         =   "ACCOUNTS PAYABLE INVOICING"
   ClientHeight    =   10185
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15255
   ControlBox      =   0   'False
   Icon            =   "frmAcctPayInvcg.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10185
   ScaleWidth      =   15255
   WindowState     =   2  'Maximized
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   15960
      TabIndex        =   46
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
      Image           =   "frmAcctPayInvcg.frx":57E2
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   15960
      TabIndex        =   47
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
      Image           =   "frmAcctPayInvcg.frx":5D7C
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
      Image           =   "frmAcctPayInvcg.frx":6316
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   10
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
      Image           =   "frmAcctPayInvcg.frx":68B0
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   11
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
      Image           =   "frmAcctPayInvcg.frx":6E4A
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   3720
      TabIndex        =   12
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
      Image           =   "frmAcctPayInvcg.frx":C63C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   5160
      TabIndex        =   13
      Top             =   120
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   661
      Caption         =   "&EXIT"
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
      Image           =   "frmAcctPayInvcg.frx":CBD6
      cBack           =   -2147483633
   End
   Begin VB.Frame fraHoldTop 
      Height          =   1815
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   15015
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
         TabIndex        =   23
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
      Begin VB.TextBox txtPONo 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Tag             =   "XYZ"
         Top             =   1365
         Width           =   3855
      End
      Begin VB.TextBox txtWayBillNo 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Tag             =   "XYZ"
         Top             =   990
         Width           =   3855
      End
      Begin VB.TextBox txtDlvrdBy 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Tag             =   "XYZ"
         Top             =   615
         Width           =   3855
      End
      Begin VB.TextBox txtSupName 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Tag             =   "XYZ"
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label lblDays 
         BackStyle       =   0  'Transparent
         Caption         =   "DAYS"
         Height          =   255
         Left            =   14160
         TabIndex        =   50
         Tag             =   "XYZ"
         Top             =   1050
         Width           =   495
      End
      Begin VB.Label lblDueDate 
         BackStyle       =   0  'Transparent
         Caption         =   "DUE DATE:"
         Height          =   255
         Left            =   11280
         TabIndex        =   27
         Tag             =   "XYZ"
         Top             =   1425
         Width           =   1215
      End
      Begin VB.Label lblTerms 
         BackStyle       =   0  'Transparent
         Caption         =   "TERMS:"
         Height          =   255
         Left            =   11280
         TabIndex        =   26
         Tag             =   "XYZ"
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label lblPckngDate 
         BackStyle       =   0  'Transparent
         Caption         =   "PACKING DATE:"
         Height          =   255
         Left            =   11280
         TabIndex        =   25
         Tag             =   "XYZ"
         Top             =   675
         Width           =   1335
      End
      Begin VB.Label lblPckngNo 
         BackStyle       =   0  'Transparent
         Caption         =   "PACKING NO.:"
         Height          =   255
         Left            =   11280
         TabIndex        =   24
         Tag             =   "XYZ"
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "P.O. NUMBER:"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Tag             =   "XYZ"
         Top             =   1425
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "WAY BILL NO.:"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Tag             =   "XYZ"
         Top             =   1050
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DELIVERED BY:"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Tag             =   "XYZ"
         Top             =   675
         Width           =   1455
      End
      Begin VB.Label lblStartBookNo 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER NAME:"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Tag             =   "XYZ"
         Top             =   300
         Width           =   1455
      End
   End
   Begin VB.TextBox txtFlxDisType 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   15960
      MaxLength       =   50
      TabIndex        =   53
      Tag             =   "XYF"
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
      TabIndex        =   52
      Tag             =   "XYF"
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
      TabIndex        =   51
      Tag             =   "XYF"
      Text            =   "txtFlxInvPrice"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   15960
      TabIndex        =   49
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
      TabIndex        =   48
      Text            =   "txtHoldGrayBackColor"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraHoldBottom 
      Height          =   7815
      Left            =   120
      TabIndex        =   28
      Top             =   2280
      Width           =   15015
      Begin VB.TextBox txtTotPayNC 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   11880
         Locked          =   -1  'True
         MaxLength       =   11
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Tag             =   "XYZ"
         Top             =   5160
         Width           =   1455
      End
      Begin VB.TextBox txtTotPayINV 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   13335
         Locked          =   -1  'True
         MaxLength       =   11
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Tag             =   "XYZ"
         Top             =   5160
         Width           =   1455
      End
      Begin VB.TextBox txtLessNC 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   11880
         Locked          =   -1  'True
         MaxLength       =   16
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Tag             =   "XYZ"
         Top             =   4800
         Width           =   1455
      End
      Begin VB.TextBox txtLessINV 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   13335
         Locked          =   -1  'True
         MaxLength       =   16
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Tag             =   "XYZ"
         Top             =   4800
         Width           =   1455
      End
      Begin VB.TextBox txtLessINVPer3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   13335
         Locked          =   -1  'True
         MaxLength       =   7
         ScrollBars      =   2  'Vertical
         TabIndex        =   62
         Tag             =   "XYZ"
         Top             =   4440
         Width           =   1455
      End
      Begin VB.TextBox txtLessNCPer3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   11880
         Locked          =   -1  'True
         MaxLength       =   7
         ScrollBars      =   2  'Vertical
         TabIndex        =   61
         Tag             =   "XYZ"
         Top             =   4440
         Width           =   1455
      End
      Begin VB.TextBox txtLessINVPer2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   13335
         Locked          =   -1  'True
         MaxLength       =   7
         ScrollBars      =   2  'Vertical
         TabIndex        =   58
         Tag             =   "XYZ"
         Top             =   4080
         Width           =   1455
      End
      Begin VB.TextBox txtLessNCPer2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   11880
         Locked          =   -1  'True
         MaxLength       =   7
         ScrollBars      =   2  'Vertical
         TabIndex        =   57
         Tag             =   "XYZ"
         Top             =   4080
         Width           =   1455
      End
      Begin VB.TextBox txtLessINVPer1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   13335
         Locked          =   -1  'True
         MaxLength       =   7
         ScrollBars      =   2  'Vertical
         TabIndex        =   55
         Tag             =   "XYZ"
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox txtLessNCPer1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   11880
         Locked          =   -1  'True
         MaxLength       =   7
         ScrollBars      =   2  'Vertical
         TabIndex        =   54
         Tag             =   "XYZ"
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox txtGrossTotINV 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   13335
         Locked          =   -1  'True
         MaxLength       =   11
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Tag             =   "XYZ"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox txtGrossTotNC 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   11880
         Locked          =   -1  'True
         MaxLength       =   11
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Tag             =   "XYZ"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.ComboBox cboStat 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         ItemData        =   "frmAcctPayInvcg.frx":123C8
         Left            =   7680
         List            =   "frmAcctPayInvcg.frx":123D2
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Tag             =   "XYZ"
         Top             =   5640
         Width           =   3375
      End
      Begin VB.TextBox txtAppBy 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   12360
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         Tag             =   "XYZ"
         Text            =   "LORENZO C. RUBINOS JR."
         Top             =   7320
         Width           =   2535
      End
      Begin VB.TextBox txtDateRecv 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   12360
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Tag             =   "XYZ"
         Top             =   6840
         Width           =   2535
      End
      Begin VB.TextBox txtPrepBy 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Tag             =   "XYZ"
         Top             =   7320
         Width           =   2535
      End
      Begin VB.TextBox txtRecBy 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Tag             =   "XYZ"
         Top             =   6840
         Width           =   2535
      End
      Begin VB.TextBox txtNotes 
         BackColor       =   &H00C0C0C0&
         Height          =   735
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Tag             =   "XYZ"
         Top             =   6000
         Width           =   7215
      End
      Begin MSFlexGridLib.MSFlexGrid flxAPI 
         Height          =   3135
         Left            =   120
         TabIndex        =   4
         Tag             =   "XYF"
         Top             =   120
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   13
         RowHeightMin    =   315
         BackColor       =   16777215
         BackColorBkg    =   12632256
         HighLight       =   2
         GridLines       =   2
         AllowUserResizing=   3
         FormatString    =   $"frmAcctPayInvcg.frx":123E6
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
         Height          =   3015
         Left            =   120
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   3480
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   5318
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         FixedCols       =   0
         RowHeightMin    =   315
         BackColorBkg    =   14737632
         HighLight       =   2
         GridLines       =   2
         AllowUserResizing=   3
         FormatString    =   "       |^INVOICE DISCOUNT  |^VALUE     |^DISCOUNT TYPE           "
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
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "LESS % 3:"
         Height          =   255
         Left            =   10320
         TabIndex        =   60
         Tag             =   "XYZ"
         Top             =   4500
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "LESS % 2:"
         Height          =   255
         Left            =   10320
         TabIndex        =   59
         Tag             =   "XYZ"
         Top             =   4140
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "LESS AMOUNT  :"
         Height          =   255
         Left            =   10320
         TabIndex        =   56
         Tag             =   "XYZ"
         Top             =   4860
         Width           =   1455
      End
      Begin VB.Label lblAppBy 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "APPROVED BY:"
         Height          =   255
         Left            =   11040
         TabIndex        =   45
         Tag             =   "XYZ"
         Top             =   7440
         Width           =   1215
      End
      Begin VB.Label lblDateRecv 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DATE RECEIVED:"
         Height          =   255
         Left            =   10800
         TabIndex        =   44
         Tag             =   "XYZ"
         Top             =   6960
         Width           =   1455
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PREPARED BY:"
         Height          =   255
         Left            =   6000
         TabIndex        =   43
         Tag             =   "XYZ"
         Top             =   7440
         Width           =   1455
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "RECEIVED BY:"
         Height          =   255
         Left            =   6000
         TabIndex        =   42
         Tag             =   "XYZ"
         Top             =   6960
         Width           =   1455
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NOTES:"
         Height          =   255
         Left            =   6840
         TabIndex        =   41
         Tag             =   "XYZ"
         Top             =   6000
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "STATUS:"
         Height          =   255
         Left            =   6720
         TabIndex        =   40
         Tag             =   "XYZ"
         Top             =   5640
         Width           =   855
      End
      Begin VB.Label llblTotPay 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL PAYABLE:"
         Height          =   255
         Left            =   10320
         TabIndex        =   39
         Tag             =   "XYZ"
         Top             =   5220
         Width           =   1455
      End
      Begin VB.Label lblLess 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "LESS % 1:"
         Height          =   255
         Left            =   10320
         TabIndex        =   38
         Tag             =   "XYZ"
         Top             =   3780
         Width           =   1455
      End
      Begin VB.Label lblGrossTot 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "GROSS TOTAL:"
         Height          =   255
         Left            =   10320
         TabIndex        =   37
         Tag             =   "XYZ"
         Top             =   3420
         Width           =   1455
      End
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   7560
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   15375
      m_MDI_Height    =   10695
   End
   Begin VB.Menu mnuInvFrm 
      Caption         =   "Invoice"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmAcctPayInvcg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim didRecordEntry As Boolean
Dim formAction As String
Dim DtlsCol As Integer
Dim DtlsRow As Integer

Dim sqlPckno As String
Dim sqlPONo As String

Public Sub ShowRecord(ByVal pPACKINGNO As String)
On Error GoTo er
    Dim tmpRs As New ADODB.Recordset
    Dim i As Integer
    
    Call modAPI.sp_load_packing_info(pPACKINGNO, tmpRs)
    If tmpRs.State = 0 Then Exit Sub
    
    If tmpRs.RecordCount <= 0 Then
        MsgBox "Packing No. " & pPACKINGNO & " is now ready for Costing!", vbInformation, "Supplier Invoice"
        tmpRs.Close
        Exit Sub
    End If
    
    
    'packingno, sprrno, TERMS, DUEDATE, DISC1VAL, DISC1TYPE, DISC2VAL, DISC2TYPE, DISC3VAL, DISC3TYPE, DISC4VAL,
    'DISC4TYPE, DISC5VAL, DISC5TYPE, VAT, PREPAREDBY, APPROVEDBY, INVSTATUS, INVNOTES, NCLESSPER, NCLESS,
    'INVLESPER, INVOICELESS, NCTTLPAY, INVOICETTLPAY, RECVDBY, DATERCVD, sparepartsid, QTY, MRCCOST, INVOICEPRICE,
    'TTLDISCOUNT, VATAMOUNT, NEWCOST, NCTOTAL, INVOICETOTAL
    
    tmpRs.MoveFirst
    '                  0      1             2         3         4          5            6       7       8           9          10
    '.FormatString = " |^PART NUMBER |^DESCRIPTION |^QTY. |^MRC COST |^INV. PRICE |^DISCOUNT |^VAT |^NEW COST |^NC TOTAL |^INV TOTAL"
        DtlsRow = 1
        While Not tmpRs.EOF
                
            flxAPI.TextMatrix(DtlsRow, 5) = Format(tmpRs!INVOICEPRICE, "#,##0.00")
            flxAPI.TextMatrix(DtlsRow, 6) = Format(tmpRs!TTLDISCOUNT, "#,##0.00")
            flxAPI.TextMatrix(DtlsRow, 7) = Format(tmpRs!VATAMOUNT, "#,##0.00")
            flxAPI.TextMatrix(DtlsRow, 8) = Format(tmpRs!newcost, "#,##0.00")
            flxAPI.TextMatrix(DtlsRow, 9) = Format(tmpRs!nctotal, "#,##0.00")
            flxAPI.TextMatrix(DtlsRow, 10) = Format(tmpRs!INVOICETOTAL, "#,##0.00")
    
        
            txtTerms.Text = tmpRs!Terms
            txtDueDate.Text = tmpRs!DueDate
            txtNotes.Text = tmpRs!INVNOTES
            txtLessNCPer1.Text = tmpRs!NCLESSPER
            txtLessNC.Text = tmpRs!NCLESS
            txtLessINVPer1.Text = tmpRs!INVLESPER
            txtLessINV.Text = tmpRs!INVOICELESS
            txtAppBy.Text = tmpRs!ApprovedBy
            
            If tmpRs!INVSTATUS = "PENDING" Then
                cboStat.ListIndex = 0
            Else
                cboStat.ListIndex = 1
            End If
    
            flxDiscount.TextMatrix(1, 2) = tmpRs!DISC1VAL
            flxDiscount.TextMatrix(2, 2) = tmpRs!DISC2VAL
            flxDiscount.TextMatrix(3, 2) = tmpRs!DISC3VAL
            flxDiscount.TextMatrix(4, 2) = tmpRs!DISC4VAL
            flxDiscount.TextMatrix(5, 2) = tmpRs!DISC5VAL
            flxDiscount.TextMatrix(6, 2) = tmpRs!Vat
            
            flxDiscount.TextMatrix(1, 3) = tmpRs!DISC1TYPE
            flxDiscount.TextMatrix(2, 3) = tmpRs!DISC2TYPE
            flxDiscount.TextMatrix(3, 3) = tmpRs!DISC3TYPE
            flxDiscount.TextMatrix(4, 3) = tmpRs!DISC4TYPE
            flxDiscount.TextMatrix(5, 3) = tmpRs!DISC5TYPE
    
            
            For i = 1 To 5
                If flxDiscount.TextMatrix(i, 2) = "" Then
                    flxDiscount.TextMatrix(i, 0) = "[  ]"
                Else
                    flxDiscount.TextMatrix(i, 0) = "[X]"
                End If
            Next
            
        DtlsRow = DtlsRow + 1
            
        tmpRs.MoveNext
    Wend
    
    DoEvents
    If flxAPI.Rows = 1 Then
        flxAPI.Rows = 2
    End If
    
    If tmpRs.State <> 0 Then tmpRs.Close
    
    DoEvents
    Call TotalGross
    Call TotalPay
    
    DoEvents
    Call EnabledSelectedButton

Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->ShowRecord()")
End Sub


Public Sub GetRecord(ByVal pPACKINGNO As String)
On Error GoTo er
    Dim tmpRs As New ADODB.Recordset

    Call modAPI.sp_get_packing_info(pPACKINGNO, tmpRs)
    If tmpRs.State = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass

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
    
    'supplier, shipper, waybillno, packinglist, packingno, packdate, partnumber, description, receive,
    'receivedby, cost, sparepartid, sprrno, daterec
    
    tmpRs.MoveFirst
    
            txtSupName.Text = tmpRs!Supplier
            txtDlvrdBy.Text = tmpRs!Shipper
            txtWayBillNo.Text = tmpRs!WayBillno
            txtPONO.Text = tmpRs!PACKINGLIST
            txtPckngNo.Text = tmpRs!PACKINGNO
            txtPckngDate.Text = tmpRs!PACKDATE
            txtRecBy.Text = tmpRs!receivedby
            txtDateRecv.Text = tmpRs!DATEREC
    
        While Not tmpRs.EOF
            If flxAPI.TextMatrix(1, 1) = "" Then
                flxAPI.TextMatrix(1, 1) = tmpRs!PartNumber
                flxAPI.TextMatrix(1, 2) = tmpRs!Description
                flxAPI.TextMatrix(1, 3) = tmpRs!RECEIVE
                flxAPI.TextMatrix(1, 4) = Format(tmpRs!Cost, "#,##0.00")
                flxAPI.TextMatrix(1, 11) = tmpRs!SPAREPARTID
                flxAPI.TextMatrix(1, 12) = tmpRs!SPRRNO
            Else
                flxAPI.AddItem _
                vbTab & tmpRs!PartNumber & _
                vbTab & tmpRs!Description & _
                vbTab & tmpRs!RECEIVE & _
                vbTab & Format(tmpRs!Cost, "#,##0.00") & _
                vbTab & "" & _
                vbTab & "" & _
                vbTab & "" & _
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
    Call ShowRecord(pPACKINGNO)
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
'4 ->EXIT
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

Sub LockedEntryObject(DataEntry As Boolean)

    DataEntry = Not DataEntry
        
        txtPckngNo.Locked = DataEntry
        txtTerms.Locked = DataEntry
        txtDueDate.Locked = DataEntry
        txtLessNCPer1.Locked = DataEntry
        txtLessINVPer1.Locked = DataEntry
        txtLessNCPer2.Locked = DataEntry
        txtLessINVPer2.Locked = DataEntry
        txtLessNCPer3.Locked = DataEntry
        txtLessINVPer3.Locked = DataEntry
        txtNotes.Locked = DataEntry
        txtAppBy.Locked = DataEntry
                                                                                                                                                                                                                                 
    If txtPckngNo.Locked Then
   
        txtPckngNo.BackColor = txtHoldGrayBackColor.BackColor
        txtTerms.BackColor = txtHoldGrayBackColor.BackColor
        txtDueDate.BackColor = txtHoldGrayBackColor.BackColor
        txtLessNCPer1.BackColor = txtHoldGrayBackColor.BackColor
        txtLessINVPer1.BackColor = txtHoldGrayBackColor.BackColor
        txtLessNCPer2.BackColor = txtHoldGrayBackColor.BackColor
        txtLessINVPer2.BackColor = txtHoldGrayBackColor.BackColor
        txtLessNCPer3.BackColor = txtHoldGrayBackColor.BackColor
        txtLessINVPer3.BackColor = txtHoldGrayBackColor.BackColor
        txtNotes.BackColor = txtHoldGrayBackColor.BackColor
        txtAppBy.BackColor = txtHoldGrayBackColor.BackColor
    Else
        
        txtPckngNo.BackColor = txtHoldWhiteBackColor.BackColor
        txtTerms.BackColor = txtHoldWhiteBackColor.BackColor
        txtDueDate.BackColor = txtHoldWhiteBackColor.BackColor
        txtLessNCPer1.BackColor = txtHoldWhiteBackColor.BackColor
        txtLessINVPer1.BackColor = txtHoldWhiteBackColor.BackColor
        txtLessNCPer2.BackColor = txtHoldWhiteBackColor.BackColor
        txtLessINVPer2.BackColor = txtHoldWhiteBackColor.BackColor
        txtLessNCPer3.BackColor = txtHoldWhiteBackColor.BackColor
        txtLessINVPer3.BackColor = txtHoldWhiteBackColor.BackColor
        txtNotes.BackColor = txtHoldWhiteBackColor.BackColor
        txtAppBy.BackColor = txtHoldWhiteBackColor.BackColor
        
    End If

End Sub

Private Sub CreatTempTable()
On Error GoTo er
        Dim Qry As String
        

        Qry = "Drop Temporary Table IF Exists sp_acctpayinvcng.tmp_suppinvdtls;"
        Call LYRcon.Execute(Qry)
        DoEvents
        Qry = "Create Temporary Table IF Not Exists sp_acctpayinvcng.tmp_suppinvdtls( " & _
        "   PACKINGNO           varchar(20), " & _
        "   SPAREPARTSID        INT(4),           " & _
        "   QTY                 INT(10),                   " & _
        "   MRCCOST             DECIMAL(16,4),        " & _
        "   INVOICEPRICE        DECIMAL(16,4),   " & _
        "   TTLDISCOUNT         DECIMAL(16,4),     " & _
        "   VATAMOUNT           DECIMAL(16,4),     " & _
        "   NEWCOST             DECIMAL(16,4),         " & _
        "   NCTOTAL             DECIMAL(16,4),         " & _
        "   INVOICETOTAL        DECIMAL(16,4)  " & _
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
        
        
        Qry = "Insert Into sp_acctpayinvcng.tmp_suppinvdtls(PACKINGNO, SPAREPARTSID, QTY, MRCCOST, INVOICEPRICE, TTLDISCOUNT, VATAMOUNT, NEWCOST, NCTOTAL, INVOICETOTAL) Values"

              '          0      1              2         3         4         5            6        7        8         9          10             11          12
       '.FormatString = " |^PART NUMBER |^DESCRIPTION |^QTY. |^MRC COST |^INV. PRICE |^DISCOUNT |^VAT |^NEW COST |^NC TOTAL |^INV TOTAL |^SPAREPARTSID |^SPRRNO"
        For i = 1 To flxAPI.Rows - 1
        
            DoEvents

            If i > 1 Then Qry = Qry & ","
        
        Qry = Qry & "('" & txtPckngNo.Text & "','" & Trim(flxAPI.TextMatrix(i, 11)) & _
                        "','" & Trim(flxAPI.TextMatrix(i, 3)) & "','" & Val(Replace(Trim(flxAPI.TextMatrix(i, 4)), ",", "")) & "','" & Val(Replace(Trim(flxAPI.TextMatrix(i, 5)), ",", "")) & _
                        "','" & Val(Replace(Trim(flxAPI.TextMatrix(i, 6)), ",", "")) & "','" & Val(Replace(Trim(flxAPI.TextMatrix(i, 7)), ",", "")) & _
                        "','" & Val(Replace(Trim(flxAPI.TextMatrix(i, 8)), ",", "")) & "','" & Val(Replace(Trim(flxAPI.TextMatrix(i, 9)), ",", "")) & "','" & Val(Replace(Trim(flxAPI.TextMatrix(i, 10)), ",", "")) & _
                        "')"

        Next
        'Text1.Text = Qry
        Call LYRcon.Execute(Qry)
        'InputBox Qry, Qry, Qry
        
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error ->InsertTempTable()")
End Sub

Function ValidateEntry() As Boolean

    ValidateEntry = False
            
        If txtSupName.Text = "" Then
            MsgBox "No record of supplier name!", vbExclamation, "VALIDATION"
            txtSupName.SetFocus
            Exit Function
        End If
        
        If txtDlvrdBy.Text = "" Then
            MsgBox "No record of delivered by!", vbExclamation, "VALIDATION"
            txtDlvrdBy.SetFocus
            Exit Function
        End If

        If txtWayBillNo.Text = "" Then
            MsgBox "No record of way bill no.!", vbExclamation, "VALIDATION"
            txtWayBillNo.SetFocus
            Exit Function
        End If
        
        If txtPONO.Text = "" Then
            MsgBox "No record of P.O. no.!", vbExclamation, "VALIDATION"
            txtPONO.SetFocus
            Exit Function
        End If
    
        If txtPckngNo.Text = "" Then
            MsgBox "No record of packing no.!", vbExclamation, "VALIDATION"
            txtPckngNo.SetFocus
            Exit Function
        End If
        
        If txtPckngDate.Text = "" Then
            MsgBox "No record of packing date!", vbExclamation, "VALIDATION"
            txtPckngDate.SetFocus
            Exit Function
        End If
        
        If txtTerms.Text = "" Then
            MsgBox "No record of terms!", vbExclamation, "VALIDATION"
            txtTerms.SetFocus
            Exit Function
        End If
        
        If Not IsDate(txtDueDate.Text) Then
             MsgBox "INVALID DATE VALUE!", vbExclamation, "Validation"
             txtDueDate.Text = Format(txtDueDate, "mm / dd / YYYY")
             txtTerms.Text = ""
             txtTerms.SetFocus
             Exit Function
         End If
        
        If txtDueDate.Text = "" Then
            MsgBox "No record of due date!", vbExclamation, "VALIDATION"
            txtDueDate.SetFocus
            Exit Function
        End If
        
    ValidateEntry = True
    
End Function

Public Sub EnabledSelectedButton()
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->PRINT
'3 ->SUMMRY
'4 ->EXIT
'5 ->SAVE
'6 ->CANCEL

    
    cmdTrans(1).Enabled = False 'Edit Button
    
    
    If cboStat.Text = "FINAL" Then
        cmdTrans(1).Enabled = False
        cmdTrans(2).Enabled = True 'PRINT Button
    ElseIf cboStat.Text = "PENDING" Then
        cmdTrans(1).Enabled = True  'Edit Button
        cmdTrans(2).Enabled = True 'PRINT Button
    Else
        cmdTrans(1).Enabled = False
        cmdTrans(2).Enabled = False
    End If
 
End Sub

Private Sub cmdTrans_Click(Index As Integer)
On Error GoTo er
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->PRINT
'3 ->SUMMRY
'4 ->EXIT
'5 ->SAVE
'6 ->CANCEL

    Dim Ans As VbMsgBoxResult
    Dim Level As String, Qry As String
    Dim Message As String
    
    Dim DisValue1 As String
    Dim DisValue2 As String
    Dim DisValue3 As String
    Dim DisValue4 As String
    Dim DisValue5 As String
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
        Qry = "SELECT concat(employees.FIRSTNAME,' ',employees.LASTNAME) as PrepBy From employees Where employees.EmpId = " & UserID & ";"
        txtPrepBy.Text = GetFieldValue(Qry)
    
    ElseIf Index = 1 Then 'EDIT
        formAction = "EDIT"
        Call LockedEntryObject(True)
        Call SetButtonVisibility(True)
        
        cboStat.Locked = False
        cboStat.BackColor = txtHoldWhiteBackColor.BackColor
        
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
        .Database.SetDataSource tmpRs1
        .ReadRecords
        End With
       
        'Subreport Cash Attachment
        With rpt.Subreport3.OpenSubreport
        .Database.SetDataSource tmpRs2
        .ReadRecords
        End With
    
       frmPrint.CRViewer1.ReportSource = rpt
       frmPrint.CRViewer1.ViewReport
       frmPrint.CRViewer1.Zoom 150
       Screen.MousePointer = vbDefault
       
       DoEvents
       frmPrint.Caption = "PRINT" & " " & "-" & " " & "[" & "ACCOUNTS PAYABLE INVOICING" & "]"
       frmPrint.WindowState = vbMaximized
       frmPrint.MySource = "Supplier Invoice"
       frmPrint.Show vbModal
        
    ElseIf Index = 3 Then 'SUMMARY
        FrmSummary.mySrc = "Supplier Invoice"
        FrmSummary.Show
        formAction = "EXIT"
        Unload Me
    ElseIf Index = 4 Then 'EXIT
        formAction = "EXIT"
        
        Ans = MsgBox("ARE YOU SURE YOU WANT TO EXIT?", vbQuestion + vbYesNo, Me.Caption)
        If Ans = vbNo Then
            formAction = ""
            Exit Sub
        End If
        
        Unload Me
        
    ElseIf Index = 5 Then  'SAVE
        
        If ValidateEntry() = False Then Exit Sub
        
        DoEvents
        If DisabledTestConnection = False Then
            If modConnection.TestConnection() = False Then Exit Sub
        End If
        
        
            'filter the discount value
            If flxDiscount.TextMatrix(1, 2) <> "" And flxDiscount.TextMatrix(1, 0) = "[X]" Then
                DisValue1 = Trim(flxDiscount.TextMatrix(1, 2))
            End If
            If flxDiscount.TextMatrix(2, 2) <> "" And flxDiscount.TextMatrix(2, 0) = "[X]" Then
                DisValue2 = Trim(flxDiscount.TextMatrix(2, 2))
            End If
            If flxDiscount.TextMatrix(3, 2) <> "" And flxDiscount.TextMatrix(3, 0) = "[X]" Then
                DisValue3 = Trim(flxDiscount.TextMatrix(3, 2))
            End If
            If flxDiscount.TextMatrix(4, 2) <> "" And flxDiscount.TextMatrix(4, 0) = "[X]" Then
                DisValue4 = Trim(flxDiscount.TextMatrix(4, 2))
            End If
            If flxDiscount.TextMatrix(5, 2) <> "" And flxDiscount.TextMatrix(5, 0) = "[X]" Then
                DisValue5 = Trim(flxDiscount.TextMatrix(5, 2))
            End If
 
        Call CreatTempTable
        
        DoEvents
        If formAction = "EDIT" Then
        
            Ans = MsgBox("DO YOU WANT TO SAVE YOUR CHANGES?", vbQuestion + vbYesNo, Me.Caption)
            If Ans = vbNo Then Exit Sub
            
            Screen.MousePointer = vbHourglass
            
            Call modAPI.sp_acctpayinvcng_update(Trim(flxAPI.TextMatrix(1, 12)), txtSupName.Text, txtPckngNo.Text, txtTerms.Text, Format(txtDueDate.Text, "yyyy-mm-dd") _
                        , Trim(flxDiscount.TextMatrix(1, 1)), DisValue1, DisType1 _
                        , Trim(flxDiscount.TextMatrix(2, 1)), DisValue2, DisType2 _
                        , Trim(flxDiscount.TextMatrix(3, 1)), DisValue3, DisType3 _
                        , Trim(flxDiscount.TextMatrix(4, 1)), DisValue4, DisType4 _
                        , Trim(flxDiscount.TextMatrix(5, 1)), DisValue5, DisType5 _
                        , Trim(flxDiscount.TextMatrix(6, 2)), txtPrepBy.Text, txtAppBy.Text, cboStat.Text, txtNotes.Text, txtLessNCPer1.Text, txtLessNCPer2.Text, txtLessNCPer3.Text, Val(Replace(txtLessNC.Text, ",", "")) _
                        , txtLessINVPer1.Text, txtLessINVPer2.Text, txtLessINVPer3.Text, Val(Replace(txtLessINV.Text, ",", "")), Val(Replace(txtTotPayNC.Text, ",", "")), Val(Replace(txtTotPayINV.Text, ",", "")) _
                        , txtRecBy.Text, Format(txtDateRecv.Text, "yyyy-mm-dd"), Level, Message)
            
        ElseIf formAction = "NEW" Then
        
            Ans = MsgBox("PLEASE MAKE SURE THAT THE DATA IS CORRECT!" & _
                vbCrLf & _
                vbCrLf & "DO YOU WISH TO CONTINUE?", vbYesNo + vbQuestion, Me.Caption)
            If Ans = vbNo Then Exit Sub
            
            Screen.MousePointer = vbHourglass

            Call modAPI.sp_acctpayinvcng_create(Trim(flxAPI.TextMatrix(1, 12)), txtPckngNo.Text, txtTerms.Text, Format(txtDueDate.Text, "yyyy-mm-dd") _
                        , Trim(flxDiscount.TextMatrix(1, 1)), DisValue1, DisType1 _
                        , Trim(flxDiscount.TextMatrix(2, 1)), DisValue2, DisType2 _
                        , Trim(flxDiscount.TextMatrix(3, 1)), DisValue3, DisType3 _
                        , Trim(flxDiscount.TextMatrix(4, 1)), DisValue4, DisType4 _
                        , Trim(flxDiscount.TextMatrix(5, 1)), DisValue5, DisType5 _
                        , Trim(flxDiscount.TextMatrix(6, 2)), txtPrepBy.Text, txtAppBy.Text, cboStat.Text, txtNotes.Text, txtLessNCPer1.Text, txtLessNCPer2.Text, txtLessNCPer3.Text, Val(Replace(txtLessNC.Text, ",", "")) _
                        , txtLessINVPer1.Text, txtLessINVPer2.Text, txtLessINVPer3.Text, Val(Replace(txtLessINV.Text, ",", "")), Val(Replace(txtTotPayNC.Text, ",", "")), Val(Replace(txtTotPayINV.Text, ",", "")) _
                        , txtRecBy.Text, Format(txtDateRecv.Text, "yyyy-mm-dd"), Level, Message)
            
        End If

        
        Screen.MousePointer = vbDefault
        
        If Trim(Level) = "" Then
            Exit Sub
        ElseIf Level = "Granted" Or Level = "Warning" Then
            If Level = "Granted" Then
                MsgBox Message, vbInformation, "SAVE."
            Else
                MsgBox Message, vbExclamation, "SAVE!"
                Exit Sub
            End If
        Else
            MsgBox Message, vbCritical, "Error!"
            Exit Sub
        End If

        Call LockedEntryObject(False)
        
        cboStat.Locked = True
        cboStat.BackColor = txtHoldGrayBackColor.BackColor
        
        Call SetButtonVisibility(False)
        
        Call EnabledSelectedButton
        
        didRecordEntry = False

        formAction = ""
        
    ElseIf Index = 6 Then 'CANCEL Button
        Ans = MsgBox("Cancel the currrent operation?", vbYesNo + vbQuestion, Me.Caption)
        If Ans = vbNo Then Exit Sub

        Call LockedEntryObject(False)
        Call SetButtonVisibility(False)
        Call ClearEntryForm(True)
        Call flxDscntRowValue
        
        cmdTrans(1).Enabled = False
        cmdTrans(2).Enabled = False
        
        cboStat.Locked = True
        cboStat.BackColor = txtHoldGrayBackColor.BackColor
        
        didRecordEntry = False

        formAction = ""

    End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error ->cmdTrans_Click()")
End Sub

Sub flxAPIColsChanger()
On Error Resume Next
    With flxAPI
        .Clear
        .Rows = 2
        .Cols = 13
        
        .FormatString = " |^PART NUMBER |^DESCRIPTION |^QTY. |^MRC COST |^INV. PRICE |^DISCOUNT |^VAT |^NEW COST |^NC TOTAL |^INV TOTAL"
        
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.1  'PART NUMBER
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.2   'DESCRIPTION
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.07  'QTY
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.07 'MRC COST
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.09  'INV. PRICE
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.09  'DISCOUNT
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.09 'VAT
        .ColWidth(8) = (.Width - .ColWidth(0)) * 0.09 'NEW COST
        .ColWidth(9) = (.Width - .ColWidth(0)) * 0.09 'NC TOTAL
        .ColWidth(10) = (.Width - .ColWidth(0)) * 0.09 'INV TOTAL
        .ColWidth(11) = 0 'SPAREPARTS ID
        .ColWidth(12) = 0 'SPRR NO
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
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.27    'INVOICE DISCOUN
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.2    'VALUE
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.5     'DISCOUNT TYPE
        
        'Merge Rows
        .TextMatrix(0, 0) = "INVOICE DISCOUNT"
        .TextMatrix(0, 1) = "INVOICE DISCOUNT"
        
        .MergeCells = 1
        .MergeRow(0) = True
        .Height = .RowHeightMin * 8
    End With
    
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
    flxDiscount.TextMatrix(6, 2) = "12%"
    
End Sub

Sub ClearEntryForm(Focus1stObject As Boolean)
   
    txtSupName.Text = ""
    txtDlvrdBy.Text = ""
    txtWayBillNo.Text = ""
    txtPONO.Text = ""
    
    txtPckngNo.Text = ""
    txtTerms.Text = ""
    txtDueDate.Text = ""
    txtPckngDate.Text = ""
    txtGrossTotNC.Text = ""
    txtGrossTotINV.Text = ""
    txtLessNC.Text = ""
    txtLessNCPer1.Text = ""
    txtLessNCPer2.Text = ""
    txtLessNCPer3.Text = ""
    txtLessINV.Text = ""
    txtLessINVPer1.Text = ""
    txtLessINVPer2.Text = ""
    txtLessINVPer3.Text = ""
    txtTotPayNC.Text = ""
    txtTotPayINV.Text = ""
    If cboStat.ListIndex >= 0 Then cboStat.ListIndex = 0
    txtNotes.Text = ""
    txtRecBy.Text = ""
    txtDateRecv.Text = ""
    txtPrepBy.Text = ""
    txtAppBy.Text = "LORENZO C. RUBINOS JR."
    
    flxAPI.Rows = 1
    flxAPI.Rows = 2
    
    DtlsRow = 0
    For DtlsRow = 1 To flxDiscount.Rows - 2
        flxDiscount.TextMatrix(DtlsRow, 2) = ""
        flxDiscount.TextMatrix(DtlsRow, 3) = ""
    Next
                              
    If didRecordEntry = False Then
        If Focus1stObject = True Then cmdTrans(0).SetFocus
    Else
        If Focus1stObject = True Then txtPckngNo.SetFocus
    End If
End Sub

Private Sub flxAPI_DblClick()
On Error GoTo er

    If didRecordEntry = False Then Exit Sub
   
    '                 0      1            2           3        4           5           6        7        8          9       10
    '.FormatString = " |^PART NUMBER |^DESCRIPTION |^QTY. |^MRC COST |^INV. PRICE |^DISCOUNT |^VAT |^NEW COST |^NC TOTAL |^INV TOTAL"
    DtlsCol = flxAPI.Col
    If DtlsCol = 5 Then ' INV. PRICE
       With txtFlxInvPrice
        .Top = fraHoldBottom.Top + flxAPI.CellTop + flxAPI.Top
        .Left = fraHoldBottom.Left + flxAPI.CellLeft + flxAPI.Left
        .Width = flxAPI.CellWidth
        .Height = flxAPI.CellHeight
        .Text = Val(Replace(Trim(flxAPI.TextMatrix(flxAPI.Row, 5)), ",", ""))  'INV. PRICE
        .Visible = True
        .BorderStyle = 1
        .SelStart = 0
        .SelLength = Len(.Text)
        .ZOrder
        .SetFocus
      End With
        
    Else
    
    txtFlxInvPrice.Visible = False

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
    Dim newcost As Double
    
    Dim DisValue1 As Double
    Dim DisValue2 As Double
    Dim DisValue3 As Double
    Dim DisValue4 As Double
    Dim DisValue5 As Double
    
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
        If flxDiscount.TextMatrix(6, 2) <> "" Then
            DisVat = Left(Trim(flxDiscount.TextMatrix(6, 2)), Len(Trim(flxDiscount.TextMatrix(6, 2))) - 1)
            DisVat = DisVat / 100
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

        invprice = Val(Replace(Trim(flxAPI.TextMatrix(DtlsRow, 5)), ",", ""))
        
        If Value1 <> 0 Then 'Computation for Invoice Discount 1
            Discount1 = invprice * Value1
            DisValue1 = invprice - Discount1
            VatValue = invprice * DisVat
        Else
            VatValue = invprice * DisVat
        End If
        
        If Value2 <> 0 Then 'Computation for Invoice Discount 2
            If Discount1 = 0 Then
                Discount2 = invprice * Value2
                DisValue2 = invprice - Discount2
                VatValue = invprice * DisVat
            Else
                Discount2 = DisValue1 * Value2
                DisValue2 = DisValue1 - Discount2
                VatValue = DisValue2 * DisVat
            End If
        Else
            VatValue = invprice * DisVat
        End If
        
        If Value3 <> 0 Then 'Computation for Invoice Discount 3
            If Discount2 = 0 Then
                Discount3 = invprice * Value3
                DisValue3 = invprice - Discount3
                VatValue = invprice * DisVat
            Else
                Discount3 = Discount2 * Value3
                DisValue3 = DisValue2 - Discount3
                VatValue = DisValue3 * DisVat
            End If
        Else
            VatValue = invprice * DisVat
        End If
        
        If Value4 <> 0 Then 'Computation for Invoice Discount 4
            If Discount3 = 0 Then
                Discount4 = invprice * Value4
                DisValue4 = invprice - Discount4
                VatValue = invprice * DisVat
            Else
                Discount4 = Discount3 * Value4
                DisValue4 = DisValue3 - Discount4
                VatValue = DisValue4 * DisVat
            End If
        Else
            VatValue = invprice * DisVat
        End If
        
        If Value5 <> 0 Then 'Computation for Invoice Discount 5
            If Discount4 = 0 Then
                Discount5 = invprice * Value5
                DisValue5 = invprice - Discount5
                VatValue = invprice * DisVat
            Else
                Discount5 = Discount4 * Value5
                DisValue5 = DisValue4 - Discount5
                VatValue = DisValue5 * DisVat
            End If
        Else
            VatValue = invprice * DisVat
        End If
        
        'Sum of all discount
        TotalDis = Discount1 + Discount2 + Discount3 + Discount4 + Discount5
        
        'Computation of new cost
        
        newcost = invprice - (TotalDis + VatValue)
        
        flxAPI.TextMatrix(DtlsRow, 6) = Format(TotalDis, "#,##0.00")
        flxAPI.TextMatrix(DtlsRow, 7) = Format(VatValue, "#,##0.00")
        flxAPI.TextMatrix(DtlsRow, 8) = Format(newcost, "#,##0.00")
        
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
    Dim newcost As Double
    Dim qty As Integer
    Dim nctotal As Double
    Dim invtotal As Double
    Dim invprice As Double
    Dim i As Integer
    Dim disvalue As Double
    Dim totdisvalue As Double
    
    
    DtlsRow = 0
    For DtlsRow = 1 To flxAPI.Rows - 1
     
        'Computation for NC TOTAL
        qty = Val(Replace(Trim(flxAPI.TextMatrix(DtlsRow, 3)), ",", ""))
        newcost = Val(Replace(Trim(flxAPI.TextMatrix(DtlsRow, 8)), ",", ""))
        
        nctotal = qty * newcost
        
        flxAPI.TextMatrix(DtlsRow, 9) = Format(nctotal, "#,##0.00")
        
        
        'Computation for INV TOTAL
        invprice = Val(Replace(Trim(flxAPI.TextMatrix(DtlsRow, 5)), ",", ""))
        
        For i = 1 To flxDiscount.Rows - 1
        
            If flxDiscount.TextMatrix(i, 3) = "INVOICE DISCOUNT" And flxDiscount.TextMatrix(i, 0) = "[X]" Then  'Get all discount value with discount type invoice discount
                disvalue = Left(Trim(flxDiscount.TextMatrix(i, 2)), Len((flxDiscount.TextMatrix(i, 2))) - 1)
                disvalue = disvalue / 100 'Convert to decimal
                totdisvalue = totdisvalue + disvalue
            End If
        Next
        
        disvalue = invprice * totdisvalue
        invtotal = qty * (invprice - disvalue)
        
        flxAPI.TextMatrix(DtlsRow, 10) = Format(invtotal, "#,##0.00")
        
        
        
    Next
        
        Call TotalGross
        Call TotalPay
        
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error ->Total()")
End Sub

Sub TotalGross()
    Dim grosstotnc As Double
    Dim grosstotinv As Double
    Dim nctotal As Double
    Dim invtotal As Double
    Dim i As Integer
    
    For i = 1 To flxAPI.Rows - 1
    
        nctotal = Val(Replace(Trim(flxAPI.TextMatrix(i, 9)), ",", ""))
        invtotal = Val(Replace(Trim(flxAPI.TextMatrix(i, 10)), ",", ""))
        
        grosstotnc = grosstotnc + nctotal
        grosstotinv = grosstotinv + invtotal
    
    Next
    
        txtGrossTotNC.Text = Format(grosstotnc, "#,##0.00")
        txtGrossTotINV.Text = Format(grosstotinv, "#,##0.00")
        
End Sub

Private Sub flxDiscount_DblClick()
On Error GoTo er
    
    If didRecordEntry = False Then Exit Sub
    
    '                          0                 1             2           3
    '.FormatString = "<INVOICE DISCOUNT |<INVOICE DISCOUNT |^VALUE  |^DISCOUNT TYPE"
    DtlsCol = flxDiscount.Col
    
    If DtlsCol = 0 Then 'INVOICE DISCOUNT
    
        
        If flxDiscount.TextMatrix(flxDiscount.Row, 2) = "" Or flxDiscount.TextMatrix(flxDiscount.Row, 0) = "" Then Exit Sub
    
        If flxDiscount.TextMatrix(flxDiscount.Row, 0) = "[  ]" Then
            flxDiscount.TextMatrix(flxDiscount.Row, 0) = "[X]"
            
            DoEvents
            Call Computations
            
        Else
            flxDiscount.TextMatrix(flxDiscount.Row, 0) = "[  ]"
            
            DoEvents
            Call Computations
            
        End If
    
    ElseIf DtlsCol = 2 Then 'VALUE
       With txtFlxValue
        .Top = fraHoldBottom.Top + flxDiscount.CellTop + flxDiscount.Top
        .Left = fraHoldBottom.Left + flxDiscount.CellLeft + flxDiscount.Left
        .Width = flxDiscount.CellWidth
        .Height = flxDiscount.CellHeight
        .Visible = True
        
        If flxDiscount.TextMatrix(flxDiscount.Row, 2) <> "" Then
            .Text = Left(Trim(flxDiscount.TextMatrix(flxDiscount.Row, 2)), Len((flxDiscount.TextMatrix(flxDiscount.Row, 2))) - 1)
        Else
            .Text = ""
        End If
        
        .BorderStyle = 1
        .SelStart = 0
        .SelLength = Len(.Text)
        .ZOrder
        .SetFocus
      End With
        
    ElseIf DtlsCol = 3 Then 'DISCOUNT TYPE
    
       If flxDiscount.Row = 6 Then Exit Sub
       
       With txtFlxDisType
        .Top = fraHoldBottom.Top + flxDiscount.CellTop + flxDiscount.Top
        .Left = fraHoldBottom.Left + flxDiscount.CellLeft + flxDiscount.Left
        .Width = flxDiscount.CellWidth
        .Height = flxDiscount.CellHeight
        .Text = flxDiscount.TextMatrix(flxDiscount.Row, 3)  'DISCOUNT TYPE
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
    
     'Call modConnection.DBConnection
     
     'frmAcctPayInvcg.Show
     
     cboStat.ListIndex = 0
     
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

Private Sub RXResizer1_ResizeComplete()
'On Error Resume Next

    'Align Save Button
    cmdTrans(5).Left = cmdTrans(0).Left
    cmdTrans(5).Top = cmdTrans(0).Top

    'Align Cancel Button
    cmdTrans(6).Left = Me.Width - 1600
    cmdTrans(6).Top = cmdTrans(0).Top
    
    fraHoldTop.Width = Me.Width - 510
    fraHoldBottom.Width = Me.Width - 510
    
    'fraHoldTop.Height = Me.Height - 9345
    fraHoldBottom.Height = Me.Height - 3345
    
    If flxAPI.TextMatrix(1, 1) = "" Then
        Call flxAPIColsChanger
    End If
    
    DtlsRow = 0
    For DtlsRow = 1 To flxDiscount.Rows - 1
        If flxDiscount.TextMatrix(DtlsRow, 2) = "" Then
            Call flxDscntColsChanger
            Call flxDscntRowValue
        End If
    Next
    
    
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
    
        flxAPI.SetFocus
        flxAPI.Col = 5
    
    End If
End Sub


Private Sub cboStat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNotes.SetFocus
End Sub


Private Sub txtFlxDisType_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFlxDisType, KeyAscii, ValidCodeCharsWithSpace)
    If KeyAscii = 13 Then
        flxDiscount.TextMatrix(flxDiscount.Row, 3) = txtFlxDisType.Text
        txtFlxDisType.Visible = False
        
        flxDiscount.SetFocus
        flxDiscount.Col = 2
        If flxDiscount.Row <> flxDiscount.Rows - 1 Then
            flxDiscount.Row = flxDiscount.Row + 1
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
        flxAPI.TextMatrix(flxAPI.Row, 5) = Format(txtFlxInvPrice.Text, "#,##0.00")
        txtFlxInvPrice.Visible = False
        
        flxAPI.SetFocus
        flxAPI.Col = 5
        If flxAPI.Row <> flxAPI.Rows - 1 Then
            flxAPI.Row = flxAPI.Row + 1
        End If
        
        Call Computations
        
    End If
End Sub

Private Sub txtFlxInvPrice_LostFocus()
    txtFlxInvPrice.Visible = False
End Sub

Private Sub txtFlxValue_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFlxValue, KeyAscii, NumbersWithDecimalPoint)
    
    If KeyAscii = 13 Then
    
        If txtFlxValue.Text = "" Then
            txtFlxValue.Visible = False
            Exit Sub
        End If
    
        If txtFlxValue.Text > 100 Then
            txtFlxValue.Text = ""
            txtFlxValue.Visible = False
            flxDiscount.SetFocus
            flxDiscount.Col = 2
            Exit Sub
        ElseIf txtFlxValue.Text <= 0 Then
            txtFlxValue.Text = ""
            txtFlxValue.Visible = False
            flxDiscount.SetFocus
            flxDiscount.Col = 2
            Exit Sub
        End If
        
        If flxDiscount.TextMatrix(flxDiscount.Row, 0) = "[X]" Then
            flxDiscount.TextMatrix(flxDiscount.Row, 0) = "[  ]"
            
            Call Computations

        End If
        
            flxDiscount.TextMatrix(flxDiscount.Row, 2) = Format(txtFlxValue.Text, "#,##0.00") & "%"
            txtFlxValue.Visible = False
            flxDiscount.SetFocus
            flxDiscount.Col = 3
            
        
    End If
End Sub

Private Sub txtFlxValue_LostFocus()
    txtFlxValue.Visible = False
End Sub

Private Sub TotalPay()
    Dim grosstotnc As Double
    Dim grosstotinv As Double
    Dim payvalue As Double
    Dim totpaync As Double
    Dim totpayinv As Double
    Dim lessinvper1 As Double
    Dim lessncper1 As Double
    Dim lessinvper2 As Double
    Dim lessncper2 As Double
    Dim lessinvper3 As Double
    Dim lessncper3 As Double
    Dim lessnc1 As Double
    Dim lessnc2 As Double
    Dim lessnc3 As Double
    Dim lessinv1 As Double
    Dim lessinv2 As Double
    Dim lessinv3 As Double
   
        lessncper1 = Val(Replace(txtLessNCPer1.Text, " , ", ""))
        lessinvper1 = Val(Replace(txtLessINVPer1.Text, " , ", ""))
        
        lessncper2 = Val(Replace(txtLessNCPer2.Text, " , ", ""))
        lessinvper2 = Val(Replace(txtLessINVPer2.Text, " , ", ""))
        
        lessncper3 = Val(Replace(txtLessNCPer3.Text, " , ", ""))
        lessinvper3 = Val(Replace(txtLessINVPer3.Text, " , ", ""))
        
  
        'Get the total of New Cost total gross
        grosstotnc = Val(Replace(txtGrossTotNC.Text, ",", ""))
        lessnc1 = grosstotnc * (lessncper1 / 100)
        
        If lessnc1 > grosstotnc Then
            MsgBox "Less is greater than gross total!", vbExclamation
            txtLessNCPer1.SetFocus
            Exit Sub
        Else
                totpaync = grosstotnc - lessnc1
                txtLessNC.Text = Format(grosstotnc - totpaync, "#,##0.00")
                txtTotPayNC.Text = Format(totpaync, "#,##0.00")
                txtLessNCPer1.Text = lessncper1 & "%"
        End If
  
        
        If txtLessNCPer2.Text <> "0%" Then
                'Get the total of new cost total payable
                payvalue = Val(Replace(txtTotPayNC.Text, ",", ""))
                lessnc2 = payvalue * (lessncper2 / 100)
                
                If lessnc2 > payvalue Then
                    MsgBox "Less is greater than gross total!", vbExclamation
                    txtLessNCPer2.SetFocus
                    Exit Sub
                Else
                        totpaync = payvalue - lessnc2
                        txtLessNC.Text = Format(grosstotnc - totpaync, "#,##0.00")
                        txtTotPayNC.Text = Format(totpaync, "#,##0.00")
                        txtLessNCPer2.Text = lessncper2 & "%"
                End If
        Else
                txtLessNCPer2.Text = "0%"
        End If
        
        If txtLessNCPer3.Text <> "0%" Then
                'Get the total of new cost total payable
                payvalue = Val(Replace(txtTotPayNC.Text, ",", ""))
                lessnc3 = payvalue * (lessncper3 / 100)
                
                If lessnc3 > payvalue Then
                    MsgBox "Less is greater than gross total!", vbExclamation
                    txtLessNCPer3.SetFocus
                    Exit Sub
                Else
                        totpaync = payvalue - lessnc3
                        txtLessNC.Text = Format(grosstotnc - totpaync, "#,##0.00")
                        txtTotPayNC.Text = Format(totpaync, "#,##0.00")
                        txtLessNCPer3.Text = lessncper3 & "%"
                End If
        Else
                txtLessNCPer3.Text = "0%"
        End If
 '------------------------------------------------------------------'
 
        'Get the total of invoice payable
        grosstotinv = Val(Replace(txtGrossTotINV.Text, ",", ""))
        lessinv1 = grosstotinv * (lessinvper1 / 100)
                
        If lessinv1 > grosstotinv Then
            MsgBox "Less is greater than gross total!", vbExclamation
            txtLessINVPer1.SetFocus
            Exit Sub
        Else
                totpayinv = grosstotinv - lessinv1
                txtLessINV.Text = Format(grosstotinv - totpayinv, "#,##0.00")
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
                        totpayinv = payvalue - lessinv2
                        txtLessINV.Text = Format(grosstotinv - totpayinv, "#,##0.00")
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
                        totpayinv = payvalue - lessinv3
                        txtLessINV.Text = Format(grosstotinv - totpayinv, "#,##0.00")
                        txtTotPayINV.Text = Format(totpayinv, "#,##0.00")
                        txtLessINVPer3.Text = lessinvper3 & "%"
                End If
        Else
                txtLessINVPer3.Text = "0%"
        End If
        
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
    Call ObjKeyTrapping.Validation(txtLessINVPer1, KeyAscii, NumbersOnly)
    
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
    
        Call TotalPay
            
            'Highlight the txtLessNCPer2
            txtLessNCPer2.SelStart = 0
            txtLessNCPer2.SelLength = Len(txtLessNCPer2.Text)
            txtLessNCPer2.SetFocus
        
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
    Call ObjKeyTrapping.Validation(txtLessINVPer2, KeyAscii, NumbersOnly)
    
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
    
        Call TotalPay
            
            'Highlight the txtLessNCPer3
             txtLessNCPer3.SelStart = 0
             txtLessNCPer3.SelLength = Len(txtLessNCPer3.Text)
             txtLessNCPer3.SetFocus

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
    Call ObjKeyTrapping.Validation(txtLessNCPer3, KeyAscii, NumbersOnly)
    
    If KeyAscii = 13 Then
    
        percnt = Val(Replace(txtLessNCPer3.Text, ",", ""))
    
        If percnt > 100 Then
            MsgBox "Percentage is greater than 100%", vbExclamation, "VALIDATION"
            'Highlight the txtLessNCPer3
            txtLessNCPer3.SelStart = 0
            txtLessNCPer3.SelLength = Len(txtLessNCPer3.Text)
            txtLessNCPer3.SetFocus
            Exit Sub
        End If
    
        Call TotalPay
        cboStat.SetFocus
        
    End If
End Sub

Private Sub txtLessNCPer1_Click()
    'Highlight the txtLessNCPer1
    txtLessNCPer1.SelStart = 0
    txtLessNCPer1.SelLength = Len(txtLessNCPer1.Text)
    txtLessNCPer1.SetFocus
End Sub

Private Sub txtLessNCPer1_KeyPress(KeyAscii As Integer)
    Dim percnt As Double
    
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtLessNCPer1, KeyAscii, NumbersOnly)
    
    If KeyAscii = 13 Then
    
        percnt = Val(Replace(txtLessNCPer1.Text, ",", ""))
    
        If percnt > 100 Then
            MsgBox "Percentage is greater than 100%", vbExclamation, "VALIDATION"
            'Highlight the txtLessNCPer1
            txtLessNCPer1.SelStart = 0
            txtLessNCPer1.SelLength = Len(txtLessNCPer1.Text)
            txtLessNCPer1.SetFocus
            Exit Sub
        End If
    
        Call TotalPay
        
        'Highlight the txtLessINVPer1
        txtLessINVPer1.SelStart = 0
        txtLessINVPer1.SelLength = Len(txtLessINVPer1.Text)
        txtLessINVPer1.SetFocus
        
    End If
End Sub

Private Sub txtLessNCPer2_Click()
    'Highlight the txtLessNCPer2
    txtLessNCPer2.SelStart = 0
    txtLessNCPer2.SelLength = Len(txtLessNCPer2.Text)
    txtLessNCPer2.SetFocus
End Sub

Private Sub txtLessNCPer2_KeyPress(KeyAscii As Integer)
    Dim percnt As Double
    
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtLessNCPer2, KeyAscii, NumbersOnly)
    
    If KeyAscii = 13 Then
    
        percnt = Val(Replace(txtLessNCPer2.Text, ",", ""))
    
        If percnt > 100 Then
            MsgBox "Percentage is greater than 100%", vbExclamation, "VALIDATION"
            'Highlight the txtLessNCPer2
            txtLessNCPer2.SelStart = 0
            txtLessNCPer2.SelLength = Len(txtLessNCPer2.Text)
            txtLessNCPer2.SetFocus
            Exit Sub
        End If
    
        Call TotalPay
        
        'Highlight the txtLessINVPer2
        txtLessINVPer2.SelStart = 0
        txtLessINVPer2.SelLength = Len(txtLessINVPer2.Text)
        txtLessINVPer2.SetFocus
        
    End If
End Sub

Private Sub txtLessNCPer3_Click()
    'Highlight the txtLessNCPer3
    txtLessNCPer3.SelStart = 0
    txtLessNCPer3.SelLength = Len(txtLessNCPer3.Text)
    txtLessNCPer3.SetFocus
End Sub

Private Sub txtLessNCPer3_KeyPress(KeyAscii As Integer)
    Dim percnt As Double
    
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtLessNCPer3, KeyAscii, NumbersOnly)
    
    If KeyAscii = 13 Then
    
        percnt = Val(Replace(txtLessNCPer3.Text, ",", ""))
    
        If percnt > 100 Then
            MsgBox "Percentage is greater than 100%", vbExclamation, "VALIDATION"
            'Highlight the txtLessNCPer3
            txtLessNCPer3.SelStart = 0
            txtLessNCPer3.SelLength = Len(txtLessNCPer3.Text)
            txtLessNCPer3.SetFocus
            Exit Sub
        End If
    
        Call TotalPay
        
        'Highlight the txtLessINVPer3
        txtLessINVPer3.SelStart = 0
        txtLessINVPer3.SelLength = Len(txtLessINVPer3.Text)
        txtLessINVPer3.SetFocus
        
    End If
End Sub

Private Sub txtPckngNo_KeyPress(KeyAscii As Integer)
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

