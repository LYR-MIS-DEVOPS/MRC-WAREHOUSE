VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSRPUpdater 
   Caption         =   "SRP UPDATER"
   ClientHeight    =   9105
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15225
   ControlBox      =   0   'False
   Icon            =   "frmSRPUpdater.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9105
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtFlxWDiscount 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   15600
      MaxLength       =   3
      TabIndex        =   50
      Tag             =   "XYZ"
      Text            =   "txtFlxDiscount"
      Top             =   7800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin lvButton.lvButtons_H cmdSaveExcel 
      Height          =   375
      Left            =   5160
      TabIndex        =   49
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
      Image           =   "frmSRPUpdater.frx":000C
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtFlxPartNo 
      Height          =   375
      Left            =   15600
      TabIndex        =   48
      Text            =   "txtFlxPartNo"
      Top             =   6960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtTag 
      Height          =   285
      Left            =   15720
      TabIndex        =   47
      Text            =   "Tag"
      Top             =   6480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer tmAutoLoad 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9120
      Top             =   0
   End
   Begin VB.ComboBox cboflxPartNo 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmSRPUpdater.frx":05A6
      Left            =   15600
      List            =   "frmSRPUpdater.frx":05A8
      TabIndex        =   45
      Tag             =   "XYZ"
      Text            =   "cboflxPartNo"
      Top             =   5760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer tmrShowRecord 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10440
      Top             =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1095
      Left            =   6600
      TabIndex        =   38
      Top             =   1320
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1931
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "FIND PART NUMBER"
      TabPicture(0)   =   "frmSRPUpdater.frx":05AA
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
         TabIndex        =   39
         Top             =   525
         Width           =   4095
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   4320
         TabIndex        =   40
         Top             =   510
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
         Image           =   "frmSRPUpdater.frx":05C6
         cBack           =   -2147483633
      End
      Begin VB.Label lblClose 
         Alignment       =   2  'Center
         Caption         =   " ( X )"
         Height          =   255
         Left            =   4920
         MousePointer    =   99  'Custom
         TabIndex        =   41
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.TextBox txtFlxDiscount 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   15600
      MaxLength       =   3
      TabIndex        =   37
      Tag             =   "XYZ"
      Text            =   "txtFlxDiscount"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFlxNewWSP 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   15600
      MaxLength       =   16
      TabIndex        =   36
      Tag             =   "XYZ"
      Text            =   "txtFlxNewWSP"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFlxNewSRP 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   15600
      MaxLength       =   16
      TabIndex        =   35
      Tag             =   "XYZ"
      Text            =   "txtFlxNewSRP"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFlxValuePer 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   15600
      MaxLength       =   16
      TabIndex        =   34
      Tag             =   "XYZ"
      Text            =   "txtFlxValuePer"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFlxProfMar 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   15600
      MaxLength       =   16
      TabIndex        =   33
      Tag             =   "XYZ"
      Text            =   "txtFlxProfMar"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   15600
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "txtHoldGrayBackColor"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   15600
      TabIndex        =   30
      Text            =   "HoldBackColor"
      Top             =   1320
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
      Image           =   "frmSRPUpdater.frx":5DB8
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   1
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
      Image           =   "frmSRPUpdater.frx":6352
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   2
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
      Image           =   "frmSRPUpdater.frx":68EC
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   3720
      TabIndex        =   3
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
      Image           =   "frmSRPUpdater.frx":C0DE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   4
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
      Image           =   "frmSRPUpdater.frx":C678
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   15600
      TabIndex        =   28
      Top             =   840
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
      Image           =   "frmSRPUpdater.frx":11E6A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   15600
      TabIndex        =   29
      Top             =   480
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
      Image           =   "frmSRPUpdater.frx":12404
      cBack           =   -2147483633
   End
   Begin VB.Frame fraHoldBottom 
      Height          =   6495
      Left            =   120
      TabIndex        =   16
      Tag             =   "XYZ"
      Top             =   2520
      Width           =   15015
      Begin MSFlexGridLib.MSFlexGrid flxSRP 
         Height          =   4575
         Left            =   120
         TabIndex        =   25
         Tag             =   "XYZ"
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8070
         _Version        =   393216
         Cols            =   15
         RowHeightMin    =   315
         BackColor       =   16777215
         BackColorBkg    =   12632256
         HighLight       =   2
         GridLines       =   2
         AllowUserResizing=   1
         FormatString    =   $"frmSRPUpdater.frx":1299E
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
      Begin VB.TextBox txtLastUpdate 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Tag             =   "XYZ"
         Top             =   6015
         Width           =   3015
      End
      Begin VB.TextBox txtEncodedBy 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Tag             =   "XYZ"
         Top             =   5640
         Width           =   3015
      End
      Begin VB.TextBox txtRemarks 
         BackColor       =   &H00C0C0C0&
         Height          =   615
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Tag             =   "XYZ"
         Top             =   4920
         Width           =   6015
      End
      Begin VB.TextBox txtAppBy 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   11160
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Tag             =   "XYZ"
         Text            =   "LORENZO C. RUBINOS JR."
         Top             =   5760
         Width           =   3255
      End
      Begin VB.ComboBox cboStat 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         ItemData        =   "frmSRPUpdater.frx":12A31
         Left            =   11160
         List            =   "frmSRPUpdater.frx":12A3B
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Tag             =   "XYZ"
         Top             =   5040
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "LAST UPDATE:"
         Height          =   255
         Left            =   480
         TabIndex        =   27
         Tag             =   "XYZ"
         Top             =   6000
         Width           =   1215
      End
      Begin VB.Label lblStat 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "STATUS:"
         Height          =   255
         Left            =   10200
         TabIndex        =   24
         Tag             =   "XYZ"
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "REMARKS:"
         Height          =   255
         Left            =   720
         TabIndex        =   23
         Tag             =   "XYZ"
         Top             =   4920
         Width           =   855
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "ENCODED BY:"
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Tag             =   "XYZ"
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label lblAppBy 
         BackStyle       =   0  'Transparent
         Caption         =   "APPROVED BY:"
         Height          =   255
         Left            =   9720
         TabIndex        =   21
         Tag             =   "XYZ"
         Top             =   5760
         Width           =   1215
      End
   End
   Begin VB.Frame fraHoldTop 
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Tag             =   "XYZ"
      Top             =   480
      Width           =   15015
      Begin VB.TextBox txtPckngDate 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   12720
         Locked          =   -1  'True
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   46
         Tag             =   "XYZ"
         Top             =   630
         Width           =   1935
      End
      Begin VB.TextBox txtSRPRefNo 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Tag             =   "XYZ"
         Top             =   990
         Width           =   3855
      End
      Begin VB.TextBox txtDateFinal 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   11
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Tag             =   "XYZ"
         Top             =   1360
         Width           =   2055
      End
      Begin VB.CheckBox chkManual 
         Caption         =   "MANUAL ENTRY"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   12720
         TabIndex        =   44
         Tag             =   "XYZ"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.ComboBox cboFilter 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         ItemData        =   "frmSRPUpdater.frx":12A4F
         Left            =   12720
         List            =   "frmSRPUpdater.frx":12A62
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Tag             =   "XYZ"
         Top             =   1000
         Width           =   1935
      End
      Begin VB.TextBox txtPONo 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Tag             =   "XYZ"
         Top             =   615
         Width           =   3855
      End
      Begin VB.TextBox txtSupName 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   1000
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Tag             =   "XYZ"
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox txtPckngNo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   12720
         Locked          =   -1  'True
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Tag             =   "XYZ"
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "FILTER BY:"
         Height          =   255
         Left            =   11160
         TabIndex        =   42
         Tag             =   "XYZ"
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "P.O. NUMBER:"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Tag             =   "XYZ"
         Top             =   705
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "SRP REF. NUMBER:"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Tag             =   "XYZ"
         Top             =   1050
         Width           =   1695
      End
      Begin VB.Label lblPckngNo 
         BackStyle       =   0  'Transparent
         Caption         =   "PACKING NO.:"
         Height          =   255
         Left            =   11160
         TabIndex        =   12
         Tag             =   "XYZ"
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label lblPckngDate 
         BackStyle       =   0  'Transparent
         Caption         =   "PACKING DATE:"
         Height          =   255
         Left            =   11160
         TabIndex        =   11
         Tag             =   "XYZ"
         Top             =   675
         Width           =   1335
      End
      Begin VB.Label lblTerms 
         BackStyle       =   0  'Transparent
         Caption         =   "SRP DATE FINAL:"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Tag             =   "XYZ"
         Top             =   1410
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog dlgDEMO 
      Left            =   11400
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   9720
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   15465
      m_MDI_Height    =   9675
   End
   Begin VB.Menu mnuSRP 
      Caption         =   "SRP"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuPopUpFormat 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuSTFindRow 
         Caption         =   "FIND PART NUMBER"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSTAddRow 
         Caption         =   "&ADD"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSTRemoveRow 
         Caption         =   "RE&MOVE"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmSRPUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmpRs As New ADODB.Recordset
Dim rsPart As New ADODB.Recordset

Public filter As String
Public formAction As String
Public SupCode As String

Dim didRecordEntry As Boolean

Dim flxRowNo As Integer
Dim flxColNo As Integer
Dim flxTTLCol As Integer

Dim DtlsCol As Integer
Dim DtlsRow As Integer

Dim NoDisSRP As Double

Dim AllArrPartNo() As String
Dim arrPartNo() As String

Dim SearchcboPartNo As Boolean
Dim PrtSrch As Boolean
Dim ans As String


Private Sub GetRefNo()
On Error GoTo er

    Call modSRP.sp_get_refno(tmpRs)
    
    If tmpRs.State = 0 Then Exit Sub

    If tmpRs.RecordCount <= 0 Then
        tmpRs.Close
        Exit Sub
    End If
      
    tmpRs.MoveFirst
        txtSRPRefNo.Text = tmpRs!RefNo
    
    tmpRs.MoveNext
    
    
Exit Sub
er:
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->GetRefNo()")
End Sub

Private Sub CreatTempTable()
On Error GoTo er
        Dim Qry As String

        Qry = "Drop Temporary Table IF Exists sp_srppricing.tmp_srppricingdtls;"
        Call LYRcon.Execute(Qry)
'        DoEvents
'        Qry = "Create Temporary Table IF Not Exists sp_srppricing.tmp_srppricingdtls( " & _
'        "   PACKINGNO           VARCHAR(20),        " & _
'        "   SPAREPARTID         INT(10),            " & _
'        "   PARTNUMBER          VARCHAR(50),        " & _
'        "   SUPPLIER            LONGTEXT,           " & _
'        "   OLDSRP              DECIMAL(16,2),      " & _
'        "   NEWSRP              DECIMAL(16,2),      " & _
'        "   SRPPMARGIN          DECIMAL(16,2),      " & _
'        "   SRPRATIO            VARCHAR(10),        " & _
'        "   DISCOUNT            VARCHAR(10),        " & _
'        "   OLDWSP              DECIMAL(16,2),      " & _
'        "   NEWWSP              DECIMAL(16,2),      " & _
'        "   WSPROFMAR           DECIMAL(16,2),      " & _
'        "   WSRATIO             VARCHAR(10)         " & _
'        ");"

        DoEvents
        Qry = "Create Temporary Table IF Not Exists sp_srppricing.tmp_srppricingdtls( " & _
        "   PACKINGNO           VARCHAR(20),        " & _
        "   SPAREPARTID         INT(10),            " & _
        "   PARTNUMBER          VARCHAR(50),        " & _
        "   SUPPLIER            LONGTEXT,           " & _
        "   OLDSRP              DECIMAL(16,2),      " & _
        "   NEWSRP              DECIMAL(16,2),      " & _
        "   SRPPMARGIN          VARCHAR(10),        " & _
        "   RDISCOUNT           VARCHAR(10),        " & _
        "   WDISCOUNT           VARCHAR(10),        " & _
        "   SRPMARKUP           VARCHAR(10),        " & _
        "   SRPMARKUPAMNT       DECIMAL(16,2)       " & _
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
        Dim discount As String, wdiscount As String
        Dim SQNO As Integer
        Dim i As Integer
        
'0       1             2            3          4            5        6          7              8           9           10          11               12            13           14
'|<PART NUMBER  |<DESCRIPTION |^OLD COST  |^NEW COST  |^OLD SRP |^NEW SRP |^SRP P-MARGIN |^SRP Ratio |^DISCOUNT |^OLD WS PRICE |^NEW WS PRICE |^WS P-MARGIN |^WS Ratio |^SPAREPARTSID


        Qry = "Insert Into sp_srppricing.tmp_srppricingdtls(PackingNo, PartNumber, Supplier, OldSRP, NewSRP, SRPPMargin, RDiscount, WDiscount, SparepartID, SRPMarkUp, SRPMarkUpAmnt) Values"

        For i = 1 To flxSRP.Rows - 1
            DoEvents
            If i > 1 Then Qry = Qry & ","
        
            If Trim(flxSRP.TextMatrix(i, 11)) = "" Then
                discount = Trim(flxSRP.TextMatrix(i, 11)) & "%"
            Else
                If Trim(flxSRP.TextMatrix(i, 11)) = "0%" Then
                    discount = Trim(flxSRP.TextMatrix(i, 11))
                Else
                    If Right(Trim(flxSRP.TextMatrix(i, 11)), 1) = "%" Then
                        discount = Trim(flxSRP.TextMatrix(i, 11))
                    Else
                        discount = Trim(flxSRP.TextMatrix(i, 11)) & "%"
                    End If
                End If
            End If
            
            If Trim(flxSRP.TextMatrix(i, 12)) = "" Then
                wdiscount = Trim(flxSRP.TextMatrix(i, 12)) & "%"
            Else
                If Trim(flxSRP.TextMatrix(i, 12)) = "0%" Then
                    wdiscount = Trim(flxSRP.TextMatrix(i, 12))
                Else
                    If Right(Trim(flxSRP.TextMatrix(i, 12)), 1) = "%" Then
                        wdiscount = Trim(flxSRP.TextMatrix(i, 12))
                    Else
                        wdiscount = Trim(flxSRP.TextMatrix(i, 12)) & "%"
                    End If
                End If
            End If
            '
        Qry = Qry & "('" & txtPckngNo.Text & "','" & Trim(flxSRP.TextMatrix(i, 1)) & "','" & txtSupName.Text & _
                        "','" & Val(Replace(Trim(flxSRP.TextMatrix(i, 6)), ",", "")) & "','" & Val(Replace(Trim(flxSRP.TextMatrix(i, 7)), ",", "")) & "','" & Format(Replace(flxSRP.TextMatrix(i, 9), "%", ""), "#,###.00") & _
                        "','" & discount & "','" & wdiscount & "','" & Trim(flxSRP.TextMatrix(i, 13)) & "','" & Replace(Trim(flxSRP.TextMatrix(i, 10)), ",", "") & "','" & Val(Replace(Trim(flxSRP.TextMatrix(i, 8)), ",", "")) & "')"

        Next
        'Text1.Text = Qry
        'InputBox Qry, Qry, Qry
        Call LYRcon.Execute(Qry)
        
        
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error ->InsertTempTable()")
End Sub

Private Sub SRPComputation()
On Error GoTo er
Dim CostVat As Double
Dim NewSRP As Double
Dim MarkUpAmnt As Double
Dim ProfMar As Double
Dim MarkUp As Double
    
    DtlsCol = flxSRP.Col
    
         CostVat = Val(Replace(Trim(flxSRP.TextMatrix(flxSRP.row, 4)), ",", ""))
         CostVat = CostVat * 1.12
         CostVat = Format(CostVat, "#,#00.00")
         
         
        'COMPUTE PROFIT MARGI & MARK UP
        If DtlsCol = 7 Then
        
            If txtFlxNewSRP.Text < CostVat Then
                MsgBox "New SRP is less than new cost!", vbExclamation, Me.Caption
                Exit Sub
            Else
                MarkUpAmnt = txtFlxNewSRP.Text - CostVat
            End If
            
            ProfMar = (MarkUpAmnt / txtFlxNewSRP.Text) * 100
            
            MarkUp = (((CostVat + MarkUpAmnt) - CostVat) / CostVat) * 100
            
            flxSRP.TextMatrix(flxSRP.row, 7) = Format(txtFlxNewSRP.Text, "#,###,#00")
            flxSRP.TextMatrix(flxSRP.row, 8) = Format(Round(MarkUpAmnt), "#,#00")
            flxSRP.TextMatrix(flxSRP.row, 9) = Format(Round(ProfMar), "#,#00") & "%"
            flxSRP.TextMatrix(flxSRP.row, 10) = Format(MarkUp, "#,###.00") & "%"
        
        'COMPUTE PROFIT MARGIN & NEW SRP
        ElseIf DtlsCol = 10 Then
        
            MarkUpAmnt = Val(Replace(txtFlxValuePer.Text, ",", "")) / 100
            MarkUpAmnt = CostVat * MarkUpAmnt
            NewSRP = CostVat + MarkUpAmnt
            ProfMar = (MarkUpAmnt / NewSRP) * 100

            flxSRP.TextMatrix(flxSRP.row, 7) = Format(Round(NewSRP), "#,###,#00")
            flxSRP.TextMatrix(flxSRP.row, 8) = Format(Round(MarkUpAmnt), "#,#00")
            flxSRP.TextMatrix(flxSRP.row, 9) = Format(Round(ProfMar), "#,#00") & "%"
            flxSRP.TextMatrix(flxSRP.row, 10) = Format(txtFlxValuePer.Text, "#,#00.00") & "%"
        End If
        
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error ->SRPComputation()")
End Sub


Private Sub WSComputation()
On Error GoTo er
Dim CostVat As Double
Dim NewWS As Double
Dim ProfMar As Double
Dim Ratio As Double
    
    DtlsCol = flxSRP.Col
    
         CostVat = Val(Replace(Trim(flxSRP.TextMatrix(flxSRP.row, 4)), ",", ""))
         CostVat = CostVat * 1.12
         CostVat = Format(CostVat, "#,#00")
         
        'COMPUTE PROFIT MARGI & Ratio
        If DtlsCol = 11 Then
        
            If txtFlxNewWSP.Text < CostVat Then
                MsgBox "New wholesale price is less than new cost!", vbExclamation, Me.Caption
                Exit Sub
            Else
                ProfMar = txtFlxNewWSP.Text - CostVat
                'ProfMar = Format(ProfMar, "#,#00")
            End If
            
            'Ratio = ProfMar / txtFlxNewWSP.Text
            Ratio = ProfMar / CostVat
          
            Ratio = Ratio * 100
            flxSRP.TextMatrix(flxSRP.row, 11) = Format(txtFlxNewWSP.Text, "#,#00")
            flxSRP.TextMatrix(flxSRP.row, 12) = Format(ProfMar, "#,#00")
            
'            If Ratio > 100 Then
'                Ratio = 100
'            End If
            
            flxSRP.TextMatrix(flxSRP.row, 13) = Format(Ratio, "#,#00.00") & "%"
            
        End If
        
        'COMPUTE NEW SRP & RATIO
        If DtlsCol = 12 Then
            NewWS = CostVat + txtFlxProfMar.Text
            NewWS = Format(NewWS, "#,#00")
            
            Ratio = txtFlxProfMar.Text / NewWS
            Ratio = txtFlxProfMar.Text / CostVat
            Ratio = Ratio * 100
            
            
            flxSRP.TextMatrix(flxSRP.row, 11) = Format(NewWS, "#,#00")
            flxSRP.TextMatrix(flxSRP.row, 12) = Format(txtFlxProfMar.Text, "#,#00")
            
'            If Ratio > 100 Then
'                Ratio = 100
'            End If
            
            flxSRP.TextMatrix(flxSRP.row, 13) = Format(Ratio, "#,#00.00") & "%"
            
        End If
        
        'COMPUTE PROFIT MARGIN & NEW SRP
        If DtlsCol = 13 Then
        
            Ratio = txtFlxValuePer.Text
            
'            If Ratio > 100 Then
                NewWS = CostVat + (CostVat * (Ratio / 100))
                ProfMar = NewWS - CostVat
                flxSRP.TextMatrix(flxSRP.row, 11) = Format(NewWS, "#,#00")
                flxSRP.TextMatrix(flxSRP.row, 12) = Format(ProfMar, "#,#00")
                flxSRP.TextMatrix(flxSRP.row, 13) = Format(txtFlxValuePer.Text, "#,#00.00") & "%"
'                Exit Sub
'            End If
'
'            Ratio = (100 - Ratio) / 100
'
'            If Ratio < 1 Then Ratio = 1
'
'            NewWS = CostVat / Ratio
'            ProfMar = NewWS - CostVat
'
'            flxSRP.TextMatrix(flxSRP.Row, 11) = Format(NewWS, "#,#00")
'            flxSRP.TextMatrix(flxSRP.Row, 12) = Format(ProfMar, "#,#00")
'            flxSRP.TextMatrix(flxSRP.Row, 13) = Format(txtFlxValuePer.Text, "#,#00.00") & "%"
            
        End If
        
    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error ->WSComputation()")
End Sub

Public Sub GetRecord(ByVal pPACKINGNO As String)
On Error GoTo er
    Dim level As String
    Dim message As String
    
    Screen.MousePointer = vbHourglass
    

    Call modSRP.sp_get_packing_info(pPACKINGNO, tmpRs, level, message)
    If tmpRs.State = 0 Then Exit Sub
    
    If level = "Warning" Then
        MsgBox message, vbInformation, "GET PACKING INFO"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


    If tmpRs.RecordCount <= 0 Then
        MsgBox "NO RECORD OF NEW COST!", vbInformation, "PACKING NO."
        
        Screen.MousePointer = vbDefault
        'Highlight the txtRefNo
        txtPckngNo.SelStart = 0
        txtPckngNo.SelLength = Len(txtPckngNo.Text)
        txtPckngNo.SetFocus
    
        tmpRs.Close
        Exit Sub
    End If
    
    flxSRP.Rows = 1
    flxSRP.Rows = 2
    
'       1              2            3          4           5          6           7            8           9           10             11             12           13          14
'|<PART NUMBER  |<DESCRIPTION |^OLD COST  |^NEW COST  |^OLD SRP |^NEW SRP |^SRP P-MARGIN |^SRP RATIO |^DISCOUNT |^OLD WS PRICE |^NEW WS PRICE |^WS P-MARGIN |^WS RATIO |^SPAREPARTSID "
    
    tmpRs.MoveFirst
    
            txtSupName.Text = tmpRs!Supplier
            txtPONO.Text = tmpRs!PONO
            'txtSRPRefNo.Text = tmpRs!RefNo
            txtPckngNo.Text = tmpRs!PACKINGNO
            txtPckngDate.Text = tmpRs!PACKDATE

        While Not tmpRs.EOF
            If flxSRP.TextMatrix(1, 1) = "" Then
                flxSRP.TextMatrix(1, 1) = tmpRs!PartNumber                              'PN
                flxSRP.TextMatrix(1, 2) = tmpRs!Description                             'DESC
                flxSRP.TextMatrix(1, 3) = Format(tmpRs!OldCost, "#,###,###.00")         'OLD CWOV
                flxSRP.TextMatrix(1, 4) = Format(tmpRs!NewCost, "#,###,###.00")         'NEW CWOV
                flxSRP.TextMatrix(1, 5) = Format((tmpRs!OldCost * 1.12), "#,###,###.00") 'CWV
                flxSRP.TextMatrix(1, 6) = Format(tmpRs!OldSRP, "#,###,###.00")          'CURRENT SRP
                flxSRP.TextMatrix(1, 7) = ""                                            'NEW SRP
                flxSRP.TextMatrix(1, 8) = ""                                            'MARK UP AMOUNT
                flxSRP.TextMatrix(1, 9) = "0%"                                          'GPM(%)
                flxSRP.TextMatrix(1, 10) = "0%"                                         'MARK UP (%)
                flxSRP.TextMatrix(1, 11) = tmpRs!MAXDISCOUNT                            'R-DISCOUNT
                flxSRP.TextMatrix(1, 12) = tmpRs!WSMAXDISCOUNT                          'W-DISCOUNT
                flxSRP.TextMatrix(1, 13) = tmpRs!SPAREPARTID                            'SPAREPARTID
                flxSRP.TextMatrix(1, 14) = IIf(IsNull(tmpRs!Supplier), "", tmpRs!Supplier) 'SUPPLIER
            Else
                flxSRP.AddItem _
                vbTab & tmpRs!PartNumber & _
                vbTab & tmpRs!Description & _
                vbTab & Format(tmpRs!OldCost, "#,###,###.00") & _
                vbTab & Format(tmpRs!NewCost, "#,###,###.00") & _
                vbTab & Format((tmpRs!OldCost * 1.12), "#,###,###.00") & _
                vbTab & Format(tmpRs!OldSRP, "#,###,###.00") & _
                vbTab & "" & _
                vbTab & "" & _
                vbTab & "0%" & _
                vbTab & "0%" & _
                vbTab & tmpRs!MAXDISCOUNT & _
                vbTab & tmpRs!WSMAXDISCOUNT & _
                vbTab & tmpRs!SPAREPARTID & _
                vbTab & IIf(IsNull(tmpRs!Supplier), "", tmpRs!Supplier)
            End If
        
    tmpRs.MoveNext
    Wend
    
    DoEvents
    Call GetRefNo
    
    If flxSRP.Rows = 1 Then
        flxSRP.Rows = 2
    End If
    
    If tmpRs.State <> 0 Then tmpRs.Close
    
    Screen.MousePointer = vbDefault
    
    MsgBox message, vbInformation, "PACKING NO."
    
    flxSRP.SetFocus
    flxSRP.Col = 7
    
Exit Sub
er:
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->GetRecord()")
End Sub

Public Sub GetSupInfo(code As String)
On Error GoTo er
    Dim flxrow As Integer
    
    Screen.MousePointer = vbHourglass

    Call modSRP.sp_partno_info(code, tmpRs)
    If tmpRs.State = 0 Then Exit Sub
    
    
    If tmpRs.RecordCount <= 0 Then
        tmpRs.Close
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
       
    tmpRs.MoveFirst
    
'        If InStr(txtSupName.Text, tmpRs!SUPPLIERNAME) = False And flxSRP.Rows > 2 Then
'            'MsgBox "Part number" & " " & cboflxPartNo.Text & " " & "has different supplier!", vbExclamation, "Different Supplier"
'            If txtSupName.Text <> "" Then
'                txtSupName.Text = txtSupName.Text & " " & "," & " " & tmpRs!SUPPLIERNAME
'            End If
'            Screen.MousePointer = vbDefault
'            'flxSRP.TextMatrix(flxSRP.Row, 1) = ""
'            'Exit Sub
'        ElseIf txtSupName.Text = "" Then
'                txtSupName.Text = tmpRs!SUPPLIERNAME
'        End If


        While Not tmpRs.EOF
  
            flxSRP.TextMatrix(flxSRP.row, 14) = tmpRs!SUPPLIERNAME

        tmpRs.MoveNext
        Wend
        
        'If txtSupName.Text <> "" Then txtSupName.Text = ""
        
        For flxrow = 1 To flxSRP.Rows - 1
            
            If InStr(txtSupName.Text, flxSRP.TextMatrix(flxrow, 14)) = False Then
                
                If txtSupName.Text <> "" Then
                    txtSupName.Text = txtSupName.Text & " " & "," & " " & flxSRP.TextMatrix(flxrow, 14)
                Else
                    txtSupName.Text = flxSRP.TextMatrix(flxrow, 14)
                End If
            End If
        Next
    
    If tmpRs.State <> 0 Then tmpRs.Close
    
    Screen.MousePointer = vbDefault
    
Exit Sub
er:
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->GetSupInfo()")
End Sub


Public Sub ShowRecord(ByVal pPACKINGNO As String)
On Error GoTo er
    
    Screen.MousePointer = vbHourglass
    
    If filter = "ME" Then
        chkManual.Value = 1
        chkManual.Enabled = False
        mnuSTAddRow.Visible = True
        mnuSTRemoveRow.Visible = True
    End If
    
    If cboFilter.ListIndex = 0 Then
        filter = ""
    ElseIf cboFilter.ListIndex = 1 Then
        filter = "NI"
    ElseIf cboFilter.ListIndex = 2 Then
        filter = "IC"
    ElseIf cboFilter.ListIndex = 3 Then
        filter = "DC"
    ElseIf cboFilter.ListIndex = 4 Then
        filter = "AI"
    End If
    
    If chkManual.Value = 1 Then filter = "ME"

    Call modSRP.sp_show_record(pPACKINGNO, filter, tmpRs)
    If tmpRs.State = 0 Then Exit Sub

    If tmpRs.RecordCount <= 0 Then
        flxSRP.Rows = 1
        flxSRP.Rows = 2
        Screen.MousePointer = vbDefault
        tmpRs.Close
        Exit Sub
    End If
    
    Call flxSRPColsChanger
    
    flxSRP.Rows = 1
    flxSRP.Rows = 2
    
'       1              2            3          4           5          6           7            8           9           10             11             12           13          14
'|<PART NUMBER  |<DESCRIPTION |^OLD COST  |^NEW COST  |^OLD SRP |^NEW SRP |^SRP P-MARGIN |^SRP RATIO |^DISCOUNT |^OLD WS PRICE |^NEW WS PRICE |^WS P-MARGIN |^WS RATIO |^SPAREPARTSID "
    
    tmpRs.MoveFirst

            txtSupName.Text = tmpRs!Supplier
            txtPONO.Text = tmpRs!PONO
            txtSRPRefNo.Text = tmpRs!RefNo
            txtPckngNo.Text = tmpRs!PACKINGNO
            txtPckngDate.Text = tmpRs!PACKDATE
            txtDateFinal.Text = tmpRs!DATEFINAL
            txtRemarks.Text = tmpRs!Remarks
            txtEncodedBy.Text = tmpRs!ENCODEDBY
            Call modObjTools.changeListIndex(cboStat, tmpRs!Status)
            txtLastUpdate.Text = tmpRs!LASTUPDATE
            txtAppBy.Text = tmpRs!ApprovedBy
            
        While Not tmpRs.EOF
            If flxSRP.TextMatrix(1, 1) = "" Then
                flxSRP.TextMatrix(1, 1) = tmpRs!PartNumber
                flxSRP.TextMatrix(1, 2) = tmpRs!Description
                flxSRP.TextMatrix(1, 3) = Format(tmpRs!OldCost, "#,##0.00")
                flxSRP.TextMatrix(1, 4) = Format(tmpRs!NewCost, "#,##0.00")
                flxSRP.TextMatrix(1, 5) = Format(tmpRs!NewCost * 1.12, "#,##0.00")
                flxSRP.TextMatrix(1, 6) = Format(tmpRs!OldSRP, "#,##0.00")
                flxSRP.TextMatrix(1, 7) = Format(tmpRs!NewSRP, "#,##0")
                flxSRP.TextMatrix(1, 8) = Format(tmpRs!SRPMarkUpAmnt, "#,##0.00")
                flxSRP.TextMatrix(1, 9) = Format(tmpRs!SRPPMargin, "#,##0") & "%"
                flxSRP.TextMatrix(1, 10) = tmpRs!SRPMarkUp & "%"
                flxSRP.TextMatrix(1, 11) = tmpRs!discount
                flxSRP.TextMatrix(1, 12) = tmpRs!wdiscount
                flxSRP.TextMatrix(1, 13) = tmpRs!SPAREPARTID
                flxSRP.TextMatrix(1, 14) = IIf(IsNull(tmpRs!Supplier), "", tmpRs!Supplier)
            Else
                flxSRP.AddItem _
                vbTab & tmpRs!PartNumber & _
                vbTab & tmpRs!Description & _
                vbTab & Format(tmpRs!OldCost, "#,##0.00") & _
                vbTab & Format(tmpRs!NewCost, "#,##0.00") & _
                vbTab & Format(tmpRs!NewCost * 1.12, "#,##0.00") & _
                vbTab & Format(tmpRs!OldSRP, "#,##0.00") & _
                vbTab & Format(tmpRs!NewSRP, "#,##0") & _
                vbTab & Format(tmpRs!SRPMarkUpAmnt, "#,##0.00") & _
                vbTab & Format(tmpRs!SRPPMargin, "#,##0") & "%" & _
                vbTab & tmpRs!SRPMarkUp & "%" & _
                vbTab & tmpRs!discount & _
                vbTab & tmpRs!wdiscount & _
                vbTab & tmpRs!SPAREPARTID & _
                vbTab & IIf(IsNull(tmpRs!Supplier), "", tmpRs!Supplier)
            End If
        
    tmpRs.MoveNext
    Wend
    
    If flxSRP.Rows = 1 Then
        flxSRP.Rows = 2
    End If
    
    If tmpRs.State <> 0 Then tmpRs.Close
    
    
    Screen.MousePointer = vbDefault
    
    Call EnabledSelectedButton
    
Exit Sub
er:
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->ShowRecord()")
End Sub

Function ValidateEntry() As String
    Dim vMsg As String, Qry As String
    Dim i As Integer, Cntr As Integer
    
    vMsg = ""
    
    If txtPckngNo.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field packing no."
    If txtEncodedBy.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field encoded by."
'    If txtSRPRefNo.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field ref. no."
    If txtAppBy.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field approved by."
    
    If txtPckngDate.Text = "" Then
        vMsg = vMsg & vbCrLf & "Incomplete entry on packing date."
    ElseIf Not IsDate(txtPckngDate.Text) Then
        vMsg = vMsg & vbCrLf & "Invalid packing date."
    End If
    
'    If txtDateFinal.Text = "" Then
'        vMsg = vMsg & vbCrLf & "Incomplete entry on SRP final date."
'    ElseIf Not IsDate(txtDateFinal.Text) Then
'        vMsg = vMsg & vbCrLf & "Invalid SRP final date."
'    End If
    

'    If Not IsDate(txtChckDate.Text) Or txtChckDate.Mask <> "## / ## / ####" Then
'        vMsg = vMsg & vbCrLf & "Invalid check date."
'    End If

    
   For i = 1 To flxSRP.Rows - 1
   
        If Val(flxSRP.TextMatrix(i, 3)) > 0 And Val(flxSRP.TextMatrix(i, 7)) = 0 Then vMsg = vMsg & vbCrLf & "There is No Price @ row " & i
        If flxSRP.TextMatrix(i, 1) = "" And flxSRP.TextMatrix(i, 2) = "" Then vMsg = vMsg & vbCrLf & "Please delete blank row!"
   Next
   
   ValidateEntry = vMsg
    
End Function



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
        txtRemarks.Locked = DataEntry
        txtAppBy.Locked = DataEntry
                                                                                                                                                                                                                                 
    If txtPckngNo.Locked Then
   
        txtPckngNo.BackColor = txtHoldGrayBackColor.BackColor
        txtRemarks.BackColor = txtHoldGrayBackColor.BackColor
        txtAppBy.BackColor = txtHoldGrayBackColor.BackColor

    Else
        
        txtPckngNo.BackColor = txtHoldWhiteBackColor.BackColor
        txtRemarks.BackColor = txtHoldWhiteBackColor.BackColor
        txtAppBy.BackColor = txtHoldWhiteBackColor.BackColor

    End If

End Sub

Sub flxSRPColsChanger()
On Error Resume Next
Dim i As Integer
    With flxSRP
        flxTTLCol = 15
        .Clear
        .Rows = 2
        .Cols = flxTTLCol

        '               0       1             2           3          4        5         6           7              8         9         10          11             12
        .FormatString = "|<PARTNUMBER |<DESCRIPTION |^OLD CwoV |^NEW CwoV |^CwV |^CURRENT SRP |^NEW SRP |^MARK-UP AMNT |^GPM (%) |^MARK-UP(%) |^R-DISCOUNT |^W-DISCOUNT "
        
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.15   'PART NUMBER
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.2  'DESCRIPTION
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.06  'OLD CwoV
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.06  'NEW CwoV
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.06  'CwV
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.07  'CURRENT SRP
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.06  'NEW SRP
        .ColWidth(8) = (.Width - .ColWidth(0)) * 0.06  'MARK-UP AMNT
        .ColWidth(9) = (.Width - .ColWidth(0)) * 0.05  'GPM (%)
        .ColWidth(10) = (.Width - .ColWidth(0)) * 0.06 'MARK-UP(%)
        .ColWidth(11) = (.Width - .ColWidth(0)) * 0.07 'R-DISCOUNT
        .ColWidth(12) = (.Width - .ColWidth(0)) * 0.06 'W-DISCOUNT
        .ColWidth(13) = 0 'SPAREPARTSID
        .ColWidth(14) = 0 'SUPPLIER
     End With
    
    
End Sub


Sub ClearEntryForm(Focus1stObject As Boolean)
   
    txtSupName.Text = ""
    txtPONO.Text = ""
    txtSRPRefNo.Text = ""
    txtPckngNo.Text = ""
    txtPckngDate.Text = ""
    txtDateFinal.Text = ""
    
    If cboStat.ListIndex >= 0 Then cboStat.ListIndex = 0
    
    txtRemarks.Text = ""
    txtAppBy.Text = "LORENZO C. RUBINOS JR."
    
    
    
    flxSRP.Rows = 1
    flxSRP.Rows = 2
                              
    If didRecordEntry = False Then
        If Focus1stObject = True Then cmdTrans(0).SetFocus
    Else
        If Focus1stObject = True Then txtPckngNo.SetFocus
    End If
    
End Sub

Private Sub cboFilter_Click()

    Call ShowRecord(txtPckngNo.Text)

End Sub


Private Sub cboflxPartNo_LostFocus()
    cboflxPartNo.Visible = False
End Sub

Private Sub cboStat_Click()
    If cboStat.ListIndex = 1 Then
        txtDateFinal.Text = Format(Now(), "mm/dd/YYYY")
    Else
        txtDateFinal.Text = ""
    End If
End Sub

Private Sub chkManual_Click()

    If didRecordEntry = False Then
        chkManual.Value = 1
        Exit Sub
    End If
    
    'If formAction = "SUMMARY" Or formAction = "EDIT" Then Exit Sub
    If formAction = "SUMMARY" Then Exit Sub
    
    If chkManual.Value = 1 Then
        txtPckngDate.Text = Format(Now(), "YYYY-mm-dd")
        mnuSTAddRow.Visible = True
        mnuSTRemoveRow.Visible = True
    ElseIf txtSupName.Text <> "" Then
        ans = MsgBox("Exit manual entry?", vbYesNo + vbQuestion, Me.Caption)
        If ans = vbNo Then
            chkManual.Value = 1
            Exit Sub
        End If
        
        Call ClearEntryForm(True)
        
        mnuSTAddRow.Visible = False
        mnuSTRemoveRow.Visible = False
        
    End If
    
End Sub

Private Sub cmdFind_Click()
On Error GoTo er
 Dim i As Integer

    If Trim(txtFind.Text) = "" Then Exit Sub
    With flxSRP
            For i = 1 To flxSRP.Rows - 1
                If InStr(Trim(flxSRP.TextMatrix(i, 1)), Trim(txtFind.Text)) Then
                    .Col = 0: .row = i
                    .ColSel = 12: .RowSel = i
                    .TopRow = i
                    .Highlight = flexHighlightAlways
                    Exit Sub
                End If
            Next i
    End With
    MsgBox "No match found!", vbInformation + vbOKOnly, "Find"

Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error ->cmdFind_Click()")
End Sub

Private Sub cmdSaveExcel_Click()
On Error GoTo er
    Dim XLapp As Excel.Application
    Dim wbkMVR As Excel.Workbook
    Dim wksLineA1 As Excel.Worksheet
    Dim sSaveAsFileName
    Dim i As Integer
    Dim j As Integer
    
    
    With dlgDEMO
        .CancelError = True
        .Flags = cdlOFNHideReadOnly
        .filename = ""
        .filter = "Microsoft Excel (*.xls)|*.xls|All Files(*.*)|*.*"
        .ShowOpen
        If dlgDEMO.filename = "" Then Exit Sub
        sSaveAsFileName = .filename
    End With
    
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Set XLapp = CreateObject("Excel.Application")
    XLapp.Visible = True
    XLapp.Workbooks.Add
    Set wbkMVR = XLapp.ActiveWorkbook
    Set wksLineA1 = wbkMVR.Sheets(1)
    
    wksLineA1.Name = "Sales"
    wksLineA1.Cells(1, 1).Value = "PARTNUMBER"
    wksLineA1.Cells(1, 2).Value = "DESCRIPTION"
    wksLineA1.Cells(1, 3).Value = "OLD COST"
    wksLineA1.Cells(1, 4).Value = "NEW COST"
    wksLineA1.Cells(1, 5).Value = "CWV"
    wksLineA1.Cells(1, 6).Value = "CURR SRP"
    wksLineA1.Cells(1, 7).Value = "NEW SRP"
    wksLineA1.Cells(1, 8).Value = "M.U AMOUNT"
    wksLineA1.Cells(1, 9).Value = "GPM(%)"
    wksLineA1.Cells(1, 10).Value = "M.U(%)"
    wksLineA1.Cells(1, 11).Value = "R-DISCOUNT"
    wksLineA1.Cells(1, 12).Value = "W-DISCOUNT"
    With flxSRP
        For i = 1 To .Rows - 1
            wksLineA1.Cells(i + 1, 1).Value = .TextMatrix(i, 1)
            wksLineA1.Cells(i + 1, 2).Value = .TextMatrix(i, 2)
            wksLineA1.Cells(i + 1, 3).Value = .TextMatrix(i, 3)
            wksLineA1.Cells(i + 1, 4).Value = .TextMatrix(i, 4)
            wksLineA1.Cells(i + 1, 5).Value = .TextMatrix(i, 5)
            wksLineA1.Cells(i + 1, 6).Value = .TextMatrix(i, 6)
            wksLineA1.Cells(i + 1, 7).Value = .TextMatrix(i, 7)
            wksLineA1.Cells(i + 1, 8).Value = .TextMatrix(i, 8)
            wksLineA1.Cells(i + 1, 9).Value = .TextMatrix(i, 9)
            wksLineA1.Cells(i + 1, 10).Value = .TextMatrix(i, 10)
            wksLineA1.Cells(i + 1, 11).Value = .TextMatrix(i, 11)
            wksLineA1.Cells(i + 1, 12).Value = .TextMatrix(i, 12)
        DoEvents
        Next i
    End With
    wksLineA1.Range("A1:L" & (i - 1)).Columns.AutoFit
    wksLineA1.Range("A1:L1").Font.Bold = True
    XLapp.DisplayAlerts = False

    
    wbkMVR.SaveAs filename:=sSaveAsFileName, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
    
    wbkMVR.Close SaveChanges:=False
    XLapp.Quit
    
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    
    Set XLapp = Nothing
    Set wbkMVR = Nothing
    Set wksLineA1 = Nothing
    MsgBox "Finish Saving to Excel!", vbInformation + vbOKOnly, Me.Caption
    
    Exit Sub
er:
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    MsgBox Err.Description, vbCritical, "Error!"
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

    Dim ans As VbMsgBoxResult
    Dim level As String
    Dim message As String
    

    If Index = 0 Then 'ADD
        formAction = "NEW"
        Call SetButtonVisibility(True)
        Call LockedEntryObject(True)
        
        didRecordEntry = True
        
        Call ClearEntryForm(True)
        
        chkManual.Enabled = True
        
        cboFilter.Locked = True
        cboFilter.BackColor = txtHoldGrayBackColor.BackColor
        
        cmdSaveExcel.Visible = False
    
    ElseIf Index = 1 Then 'EDIT
        formAction = "EDIT"
        
'        If filter = "ME" Then
'            Screen.MousePointer = vbHourglass
'            tmAutoLoad.Enabled = True
'            DoEvents
'            Call Load(frm_OnProcess2)
'            frm_OnProcess2.Show vbModal
'
'        End If
        
        Call LockedEntryObject(True)
        Call SetButtonVisibility(True)
        
        If UserID = 2 Then ' Or SupCode = "AA" Or SupCode = "CA" Or SupCode = "BZ" Or SupCode = "SP" Then
            cboStat.Locked = False
            cboStat.BackColor = txtHoldWhiteBackColor.BackColor
        End If
        
        cboStat.Locked = False
        cboStat.BackColor = txtHoldWhiteBackColor.BackColor
        cboFilter.Locked = True
        cboFilter.BackColor = txtHoldGrayBackColor.BackColor
        
        chkManual.Enabled = False
        
        didRecordEntry = True
    
        cmdSaveExcel.Visible = False
    ElseIf Index = 2 Then 'PRINT
       Dim Qry As String
       
       Screen.MousePointer = vbHourglass
       Dim rpt As New rptSRPPricing

        'PROCEDURE sp_srppricing.sp_show_record(
        '  pPACKINGNO VarChar(20)
        '  )
        
        If cboFilter.ListIndex = 0 Then
            filter = ""
        ElseIf cboFilter.ListIndex = 1 Then
            filter = "NI"
        ElseIf cboFilter.ListIndex = 1 Then
            filter = "IC"
        ElseIf cboFilter.ListIndex = 1 Then
            filter = "DC"
        End If
        
        If chkManual.Value = 1 Then filter = "ME"


        Qry = "Call sp_srppricing.sp_show_record('" & txtPckngNo.Text & "','" & filter & "');"
        DoEvents
        If tmpRs.State <> 0 Then tmpRs.Close
        Set tmpRs = LYRcon.Execute(Qry)


        DoEvents
        With rpt
            .DiscardSavedData
             DoEvents
            .Database.SetDataSource tmpRs
            .ReadRecords
        End With


       frmPrint.CRViewer1.ReportSource = rpt
       frmPrint.CRViewer1.ViewReport
       frmPrint.CRViewer1.Zoom 150
       Screen.MousePointer = vbDefault

       DoEvents
       frmPrint.Caption = "PRINT" & " " & "-" & " " & "[" & "SUGGESTED RETAIL PRICE" & "]"
       frmPrint.WindowState = vbMaximized
       frmPrint.MySource = "SRP PRICING"
       frmPrint.Show vbModal
        
    ElseIf Index = 3 Then 'SUMMARY
        FrmSummary.mySrc = "SRP Pricing"
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
        Dim ChckDate As String
        Dim EntryType As String
        
        If ValidateEntry() <> "" Then
            MsgBox ValidateEntry, vbExclamation, Me.Caption
            Exit Sub
        End If
        
'        DoEvents
'        If DisabledTestConnection = False Then
'            If modConnection.TestConnection() = False Then Exit Sub
'        End If
        
        Screen.MousePointer = vbHourglass
 
        Call CreatTempTable
        
        DoEvents
        If formAction = "EDIT" Then
        
            ans = MsgBox("DO YOU WANT TO SAVE YOUR CHANGES?", vbQuestion + vbYesNo, Me.Caption)
            If ans = vbNo Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
            Screen.MousePointer = vbHourglass
            Call modSRP.sp_srppricing_update(txtPckngNo.Text, txtDateFinal.Text, cboStat.Text, txtRemarks.Text, txtEncodedBy.Text, txtAppBy.Text, level, message)
            
            
        ElseIf formAction = "NEW" Then
        
            ans = MsgBox("PLEASE MAKE SURE THAT THE DATA IS CORRECT!" & _
                vbCrLf & _
                vbCrLf & "DO YOU WISH TO CONTINUE?", vbYesNo + vbQuestion, Me.Caption)
            If ans = vbNo Then Exit Sub
            
            If chkManual.Value = 1 Then
                EntryType = "MANUAL"
            Else
                EntryType = "AUTO"
            End If
            
            Screen.MousePointer = vbHourglass
            Call modSRP.sp_srppricing_create(txtPckngNo.Text, txtDateFinal.Text, cboStat.Text, txtRemarks.Text, txtEncodedBy.Text, txtAppBy.Text, EntryType, level, message)

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
        
        cboStat.Locked = True
        cboStat.BackColor = txtHoldGrayBackColor.BackColor
        
        cmdTrans(0).SetFocus
        
        cmdSaveExcel.Visible = True
        
        formAction = ""
        
    ElseIf Index = 6 Then 'CANCEL Button
        ans = MsgBox("Cancel the currrent operation?", vbYesNo + vbQuestion, Me.Caption)
        If ans = vbNo Then Exit Sub
        
        Call LockedEntryObject(False)
        Call SetButtonVisibility(False)
        Call ClearEntryForm(True)
        Call EnabledSelectedButton
        
        cboStat.Locked = True
        cboStat.BackColor = txtHoldGrayBackColor.BackColor
        
        chkManual.Enabled = True
        chkManual.Value = 0
        chkManual.Enabled = False

        didRecordEntry = False
        formAction = ""
        
        cmdSaveExcel.Visible = True
    End If

Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error ->cmdTrans_Click()")
End Sub

Private Sub flxSRP_Click()
    
    If didRecordEntry = False Then Exit Sub
    
    
    If chkManual.Value = 1 Then
        DtlsCol = flxSRP.Col
        DtlsRow = flxSRP.row
        If DtlsCol = 1 Then 'PART NUMBER
'           With cboflxPartNo
'            .Top = fraHoldBottom.Top + flxSRP.CellTop + flxSRP.Top
'            .Left = fraHoldBottom.Left + flxSRP.CellLeft + flxSRP.Left
'            .Width = flxSRP.CellWidth
'            .Text = flxSRP.TextMatrix(flxSRP.Row, DtlsCol)
'            .Visible = True
'            .SelStart = 0
'            .SelLength = Len(.Text)
'            .ZOrder
'            .SetFocus
'          End With
           With txtFlxPartNo
            .Top = fraHoldBottom.Top + flxSRP.CellTop + flxSRP.Top
            .Left = fraHoldBottom.Left + flxSRP.CellLeft + flxSRP.Left
            .Width = flxSRP.CellWidth
            .Height = flxSRP.CellHeight
            .Text = flxSRP.TextMatrix(flxSRP.row, DtlsCol)
            .Visible = True
            .BorderStyle = 1
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
          End With
        End If
        
        If DtlsCol = 7 Then 'NEW SRP
           With txtFlxNewSRP
            .Top = fraHoldBottom.Top + flxSRP.CellTop + flxSRP.Top
            .Left = fraHoldBottom.Left + flxSRP.CellLeft + flxSRP.Left
            .Width = flxSRP.CellWidth
            .Height = flxSRP.CellHeight
            .Text = Val(Replace(Trim(flxSRP.TextMatrix(flxSRP.row, 7)), ",", ""))  'NEW SRP
            .Visible = True
            .BorderStyle = 1
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
          End With
    
'        ElseIf DtlsCol = 7 Then 'PROFIT MARGIN
'           With txtFlxProfMar
'            .Top = fraHoldBottom.Top + flxSRP.CellTop + flxSRP.Top
'            .Left = fraHoldBottom.Left + flxSRP.CellLeft + flxSRP.Left
'            .Width = flxSRP.CellWidth
'            .Height = flxSRP.CellHeight
'            .Text = Val(Replace(Trim(flxSRP.TextMatrix(flxSRP.row, 7)), ",", ""))  'PROFIT MARGIN
'            .Visible = True
'            .BorderStyle = 1
'            .SelStart = 0
'            .SelLength = Len(.Text)
'            .ZOrder
'            .SetFocus
'          End With
          
        ElseIf DtlsCol = 10 Then 'MARK-UP(%)
           With txtFlxValuePer
            .Top = fraHoldBottom.Top + flxSRP.CellTop + flxSRP.Top
            .Left = fraHoldBottom.Left + flxSRP.CellLeft + flxSRP.Left
            .Width = flxSRP.CellWidth
            .Height = flxSRP.CellHeight
    
            If flxSRP.TextMatrix(flxSRP.row, 10) <> "" Then
                .Text = Left(Trim(flxSRP.TextMatrix(flxSRP.row, 10)), Len((flxSRP.TextMatrix(flxSRP.row, 10))) - 1)
            Else
                .Text = ""
            End If
    
            .Visible = True
            .BorderStyle = 1
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
          End With
          
        ElseIf DtlsCol = 11 Then 'DISCOUNT
           With txtFlxDiscount
            .Top = fraHoldBottom.Top + flxSRP.CellTop + flxSRP.Top
            .Left = fraHoldBottom.Left + flxSRP.CellLeft + flxSRP.Left
            .Width = flxSRP.CellWidth
            .Height = flxSRP.CellHeight

            If flxSRP.TextMatrix(flxSRP.row, 11) <> "" Then
                .Text = Left(Trim(flxSRP.TextMatrix(flxSRP.row, 11)), Len((flxSRP.TextMatrix(flxSRP.row, 11))) - 1)
            Else
                .Text = ""
            End If

            .Visible = True
            .BorderStyle = 1
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
          End With
          
'        ElseIf DtlsCol = 11 Then 'NEW WS PRICE
'           With txtFlxNewWSP
'            .Top = fraHoldBottom.Top + flxSRP.CellTop + flxSRP.Top
'            .Left = fraHoldBottom.Left + flxSRP.CellLeft + flxSRP.Left
'            .Width = flxSRP.CellWidth
'            .Height = flxSRP.CellHeight
'            .Text = Val(Replace(Trim(flxSRP.TextMatrix(flxSRP.row, 11)), ",", ""))  'NEW WS PRICE
'            .Visible = True
'            .BorderStyle = 1
'            .SelStart = 0
'            .SelLength = Len(.Text)
'            .ZOrder
'            .SetFocus
'          End With
          
'        ElseIf DtlsCol = 12 Then 'NEW WS P-MARGIN
'           With txtFlxProfMar
'            .Top = fraHoldBottom.Top + flxSRP.CellTop + flxSRP.Top
'            .Left = fraHoldBottom.Left + flxSRP.CellLeft + flxSRP.Left
'            .Width = flxSRP.CellWidth
'            .Height = flxSRP.CellHeight
'            .Text = Val(Replace(Trim(flxSRP.TextMatrix(flxSRP.row, 12)), ",", ""))  'NEW WS P-MARGIN
'            .Visible = True
'            .BorderStyle = 1
'            .SelStart = 0
'            .SelLength = Len(.Text)
'            .ZOrder
'            .SetFocus
'          End With
          
'        ElseIf DtlsCol = 13 Then 'WS Ratio
'           With txtFlxValuePer
'            .Top = fraHoldBottom.Top + flxSRP.CellTop + flxSRP.Top
'            .Left = fraHoldBottom.Left + flxSRP.CellLeft + flxSRP.Left
'            .Width = flxSRP.CellWidth
'            .Height = flxSRP.CellHeight
'
'            If flxSRP.TextMatrix(flxSRP.row, 13) <> "" Then
'                .Text = Left(Trim(flxSRP.TextMatrix(flxSRP.row, 13)), Len((flxSRP.TextMatrix(flxSRP.row, 13))) - 1)
'            Else
'                .Text = ""
'            End If
'
'            .Visible = True
'            .BorderStyle = 1
'            .SelStart = 0
'            .SelLength = Len(.Text)
'            .ZOrder
'            .SetFocus
'          End With
'        ElseIf DtlsCol = 16 Then 'WDISCOUNT
'           With txtFlxWDiscount
'            .Top = fraHoldBottom.Top + flxSRP.CellTop + flxSRP.Top
'            .Left = fraHoldBottom.Left + flxSRP.CellLeft + flxSRP.Left
'            .Width = flxSRP.CellWidth
'            .Height = flxSRP.CellHeight
'
'            If flxSRP.TextMatrix(flxSRP.row, 16) <> "" Then
'                .Text = Left(Trim(flxSRP.TextMatrix(flxSRP.row, 16)), Len((flxSRP.TextMatrix(flxSRP.row, 16))) - 1)
'            Else
'                .Text = ""
'            End If
'
'            .Visible = True
'            .BorderStyle = 1
'            .SelStart = 0
'            .SelLength = Len(.Text)
'            .ZOrder
'            .SetFocus
'          End With
        End If

    End If
    
End Sub

Private Sub flxSRP_DblClick()
On Error GoTo er
   
    DtlsCol = flxSRP.Col
    If DtlsCol = 1 Then 'PART NUMBER
       With txtFlxPartNo
        .Top = fraHoldBottom.Top + flxSRP.CellTop + flxSRP.Top
        .Left = fraHoldBottom.Left + flxSRP.CellLeft + flxSRP.Left
        .Width = flxSRP.CellWidth
        .Height = flxSRP.CellHeight
        .Text = flxSRP.TextMatrix(flxSRP.row, 1)  'PART NUMBER
        .Visible = True
        .BorderStyle = 1
        .SelStart = 0
        .SelLength = Len(.Text)
        .ZOrder
        .SetFocus
      End With
    End If
    
    If didRecordEntry = False Then Exit Sub
    
    If DtlsCol = 7 Then 'NEW SRP
       With txtFlxNewSRP
        .Top = fraHoldBottom.Top + flxSRP.CellTop + flxSRP.Top
        .Left = fraHoldBottom.Left + flxSRP.CellLeft + flxSRP.Left
        .Width = flxSRP.CellWidth
        .Height = flxSRP.CellHeight
        .Text = Val(Replace(Trim(flxSRP.TextMatrix(flxSRP.row, 7)), ",", ""))  'NEW SRP
        .Visible = True
        .BorderStyle = 1
        .SelStart = 0
        .SelLength = Len(.Text)
        .ZOrder
        .SetFocus
      End With

'    ElseIf DtlsCol = 7 Then 'PROFIT MARGIN
'       With txtFlxProfMar
'        .Top = fraHoldBottom.Top + flxSRP.CellTop + flxSRP.Top
'        .Left = fraHoldBottom.Left + flxSRP.CellLeft + flxSRP.Left
'        .Width = flxSRP.CellWidth
'        .Height = flxSRP.CellHeight
'        .Text = Val(Replace(Trim(flxSRP.TextMatrix(flxSRP.row, 7)), ",", ""))  'PROFIT MARGIN
'        .Visible = True
'        .BorderStyle = 1
'        .SelStart = 0
'        .SelLength = Len(.Text)
'        .ZOrder
'        .SetFocus
'      End With
      
    ElseIf DtlsCol = 10 Then 'MARK-UP(%)
       With txtFlxValuePer
        .Top = fraHoldBottom.Top + flxSRP.CellTop + flxSRP.Top
        .Left = fraHoldBottom.Left + flxSRP.CellLeft + flxSRP.Left
        .Width = flxSRP.CellWidth
        .Height = flxSRP.CellHeight

        If flxSRP.TextMatrix(flxSRP.row, 10) <> "" Then
            .Text = Left(Trim(flxSRP.TextMatrix(flxSRP.row, 10)), Len((flxSRP.TextMatrix(flxSRP.row, 10))) - 1)
        Else
            .Text = ""
        End If

        .Visible = True
        .BorderStyle = 1
        .SelStart = 0
        .SelLength = Len(.Text)
        .ZOrder
        .SetFocus
      End With
      
    ElseIf DtlsCol = 11 Then 'DISCOUNT
       With txtFlxDiscount
        .Top = fraHoldBottom.Top + flxSRP.CellTop + flxSRP.Top
        .Left = fraHoldBottom.Left + flxSRP.CellLeft + flxSRP.Left
        .Width = flxSRP.CellWidth
        .Height = flxSRP.CellHeight

        If flxSRP.TextMatrix(flxSRP.row, 11) <> "" Then
            .Text = Left(Trim(flxSRP.TextMatrix(flxSRP.row, 11)), Len((flxSRP.TextMatrix(flxSRP.row, 11))) - 1)
        Else
            .Text = ""
        End If

        .Visible = True
        .BorderStyle = 1
        .SelStart = 0
        .SelLength = Len(.Text)
        .ZOrder
        .SetFocus
      End With
      
'    ElseIf DtlsCol = 11 Then 'NEW WS PRICE
'       With txtFlxNewWSP
'        .Top = fraHoldBottom.Top + flxSRP.CellTop + flxSRP.Top
'        .Left = fraHoldBottom.Left + flxSRP.CellLeft + flxSRP.Left
'        .Width = flxSRP.CellWidth
'        .Height = flxSRP.CellHeight
'        .Text = Val(Replace(Trim(flxSRP.TextMatrix(flxSRP.row, 11)), ",", ""))  'NEW WS PRICE
'        .Visible = True
'        .BorderStyle = 1
'        .SelStart = 0
'        .SelLength = Len(.Text)
'        .ZOrder
'        .SetFocus
'      End With
      
'    ElseIf DtlsCol = 12 Then 'NEW WS P-MARGIN
'       With txtFlxProfMar
'        .Top = fraHoldBottom.Top + flxSRP.CellTop + flxSRP.Top
'        .Left = fraHoldBottom.Left + flxSRP.CellLeft + flxSRP.Left
'        .Width = flxSRP.CellWidth
'        .Height = flxSRP.CellHeight
'        .Text = Val(Replace(Trim(flxSRP.TextMatrix(flxSRP.row, 12)), ",", ""))  'NEW WS P-MARGIN
'        .Visible = True
'        .BorderStyle = 1
'        .SelStart = 0
'        .SelLength = Len(.Text)
'        .ZOrder
'        .SetFocus
'      End With
      
'    ElseIf DtlsCol = 13 Then 'WS Ratio
'       With txtFlxValuePer
'        .Top = fraHoldBottom.Top + flxSRP.CellTop + flxSRP.Top
'        .Left = fraHoldBottom.Left + flxSRP.CellLeft + flxSRP.Left
'        .Width = flxSRP.CellWidth
'        .Height = flxSRP.CellHeight
'
'        If flxSRP.TextMatrix(flxSRP.row, 13) <> "" Then
'            .Text = Left(Trim(flxSRP.TextMatrix(flxSRP.row, 13)), Len((flxSRP.TextMatrix(flxSRP.row, 13))) - 1)
'        Else
'            .Text = ""
'        End If
'
'        .Visible = True
'        .BorderStyle = 1
'        .SelStart = 0
'        .SelLength = Len(.Text)
'        .ZOrder
'        .SetFocus
'      End With

'    ElseIf DtlsCol = 16 Then 'WDISCOUNT
'           With txtFlxWDiscount
'            .Top = fraHoldBottom.Top + flxSRP.CellTop + flxSRP.Top
'            .Left = fraHoldBottom.Left + flxSRP.CellLeft + flxSRP.Left
'            .Width = flxSRP.CellWidth
'            .Height = flxSRP.CellHeight
'
'            If flxSRP.TextMatrix(flxSRP.row, 16) <> "" Then
'                .Text = Left(Trim(flxSRP.TextMatrix(flxSRP.row, 16)), Len((flxSRP.TextMatrix(flxSRP.row, 16))) - 1)
'            Else
'                .Text = ""
'            End If
'
'            .Visible = True
'            .BorderStyle = 1
'            .SelStart = 0
'            .SelLength = Len(.Text)
'            .ZOrder
'            .SetFocus
'          End With
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error ->flxSRP_DblClick()")
End Sub

Private Sub flxSRP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
        If chkManual.Value = 1 Then
            Call flxSRP_Click
        Else
            Call flxSRP_DblClick
        End If
        
    End If
End Sub

Private Sub flxSRP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If didRecordEntry = False Then Exit Sub
    If Button = vbRightButton Then
        PopupMenu mnuPopUpFormat, vbPopupMenuRightButton
    End If
End Sub

Private Sub Form_Load()
On Error GoTo er
    
    Warehouse.sBar.Panels(6).Text = "SRP Updater"
    
    If formAction = "SUMMARY" Then Exit Sub
    
    
'    Screen.MousePointer = vbHourglass
'    tmAutoLoad.Enabled = True
'    DoEvents
'    Call Load(frm_OnProcess2)
'    frm_OnProcess2.Show vbModal
    
    cboStat.ListIndex = 0
    txtEncodedBy.Text = EmployeesName
    txtLastUpdate.Text = Now()
    
    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error ->Form_Load()")
End Sub

Sub setcbolist()
Dim tmpRs As ADODB.Recordset

    DoEvents
    Call modSRP.sp_partno_list(tmpRs)
    If tmpRs.State = 1 Then
        If tmpRs.RecordCount > 0 Then
            Call modObjTools.fillArray(AllArrPartNo, tmpRs)
        End If
    End If
    If tmpRs.State <> 0 Then tmpRs.Close
    
    DoEvents
    Call modObjTools.addToListObjectUsingArrayWithCols(cboflxPartNo, arrPartNo, AllArrPartNo, 1, 1)
End Sub

Private Sub lblClose_Click()
    SSTab1.Visible = False
    txtFind.Text = ""
End Sub

Private Sub mnuSTAddRow_Click()
On Error GoTo er

    If didRecordEntry = False Then Exit Sub

     flxSRP.AddItem "" & _
     vbTab & "" & _
     vbTab & "" & _
     vbTab & "" & _
     vbTab & "" & _
     vbTab & "" & _
     vbTab & "" & _
     vbTab & "" & _
     vbTab & "" & _
     vbTab & "" & _
     vbTab & "" & _
     vbTab & "" & _
     vbTab & "" & _
     vbTab & ""

     
     flxSRP.SetFocus
     flxSRP.Col = 1
     flxSRP.row = flxSRP.Rows - 1

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->mnuSTAddRow_Click()")
End Sub

Private Sub mnuSTFindRow_Click()
    SSTab1.Visible = True
    txtFind.SetFocus
End Sub

Private Sub mnuSTRemoveRow_Click()
On Error GoTo er
    Dim ans As String
    Dim flxrow As Integer
    
    If didRecordEntry = False Then Exit Sub
    
    ans = MsgBox("Delete row?", vbYesNo + vbQuestion, "Delete row")
    If ans = vbNo Then Exit Sub
    
    txtSupName.Text = ""
    
    If flxSRP.Rows = 1 Then
        mnuSTRemoveRow.Enabled = False
        Exit Sub
    End If
    
    If flxSRP.Rows = 2 Then
        
        flxSRP.Rows = 1
        flxSRP.Rows = 2
       
    Else
        
        flxSRP.RemoveItem (flxSRP.RowSel)
        
        For flxrow = 1 To flxSRP.Rows - 1 'Reset Supplier Name
            
            If InStr(txtSupName.Text, flxSRP.TextMatrix(flxrow, 14)) = False Then
                
                If txtSupName.Text <> "" Then
                    txtSupName.Text = txtSupName.Text & " " & "," & " " & flxSRP.TextMatrix(flxrow, 14)
                Else
                    txtSupName.Text = flxSRP.TextMatrix(flxrow, 14)
                End If
            End If
        Next
        
    End If
    
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error ->mnuSTRemoveRow_Click()"
End Sub

Private Sub RXResizer1_ResizeComplete()
    'Align Save Button
    cmdTrans(5).Left = cmdTrans(0).Left
    cmdTrans(5).Top = cmdTrans(0).Top

    'Align Cancel Button
    cmdTrans(6).Left = Me.Width - 1600
    cmdTrans(6).Top = cmdTrans(0).Top
    
'    If flxSRP.TextMatrix(1, 1) = "" Then
'        Call flxSRPColsChanger
'    End If
    
    With flxSRP
        flxTTLCol = 15
        .Cols = flxTTLCol

        '               0       1             2           3          4        5         6           7              8         9         10          11             12
        .FormatString = "|<PARTNUMBER |<DESCRIPTION |^OLD CwoV |^NEW CwoV |^CwV |^CURRENT SRP |^NEW SRP |^MARK-UP AMNT |^GPM (%) |^MARK-UP(%) |^R-DISCOUNT |^W-DISCOUNT "
        
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.15   'PART NUMBER
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.2  'DESCRIPTION
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.06  'OLD CwoV
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.06  'NEW CwoV
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.06  'CwV
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.07  'CURRENT SRP
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.06  'NEW SRP
        .ColWidth(8) = (.Width - .ColWidth(0)) * 0.07  'MARK-UP AMNT
        .ColWidth(9) = (.Width - .ColWidth(0)) * 0.05  'GPM (%)
        .ColWidth(10) = (.Width - .ColWidth(0)) * 0.07 'MARK-UP(%)
        .ColWidth(11) = (.Width - .ColWidth(0)) * 0.07 'R-DISCOUNT
        .ColWidth(12) = (.Width - .ColWidth(0)) * 0.06 'W-DISCOUNT
        .ColWidth(13) = 0 'SPAREPARTSID
        .ColWidth(14) = 0 'SUPPLIER
        
     End With
     
End Sub

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
    
    If cboStat.Text = "FINAL" And txtSupName.Text <> "" Then
        cmdTrans(1).Enabled = False 'Edit Button
        cmdTrans(2).Enabled = True  'PRINT Button
        
    ElseIf cboStat.Text = "PENDING" And txtSupName.Text <> "" Then
        cmdTrans(1).Enabled = True  'Edit Button
        cmdTrans(2).Enabled = True  'PRINT Button
    
    Else
        cmdTrans(1).Enabled = False 'Edit Button
        cmdTrans(2).Enabled = False 'Print Button
    End If
    
    If filter <> "ME" Then
        cboFilter.Locked = False
        cboFilter.BackColor = txtHoldWhiteBackColor.BackColor
    End If
 
End Sub



Private Sub tmAutoLoad_Timer()
On Error GoTo er

    tmAutoLoad.Enabled = False
    Screen.MousePointer = vbHourglass
    frmSRPUpdater.Enabled = False
    
    DoEvents
    DoEvents
    Call setcbolist
    DoEvents
    DoEvents
    
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    
    DoEvents
    frmSRPUpdater.Enabled = True
    Screen.MousePointer = vbDefault

    
    didRecordEntry = False
    
Exit Sub
er:
    frmSRPUpdater.Enabled = True
    Screen.MousePointer = vbDefault
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    Call modObjTools.ErrorMessage(Err.Description, "Error->tmAutoLoad_Timer()")
End Sub

Public Sub tmrShowRecord_Timer()
On Error GoTo er
    If tmrShowRecord.Enabled = False Then Exit Sub
    Screen.MousePointer = vbHourglass

    tmrShowRecord.Enabled = False

    DoEvents
    Call ShowRecord(txtPckngNo.Text)

    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If

    Screen.MousePointer = Default

Exit Sub
er:
    tmrShowRecord.Enabled = False
    Screen.MousePointer = Default
End Sub

Private Sub txtFind_Change()
On Error GoTo er
    Dim Qry As String

    If PrtSrch Then
        PrtSrch = False
        Exit Sub
    End If

'    If rsPart.State = adStateOpen Then rsPart.Close
'    rsPart.Open
'
'    If txtFind.Text = "" Then Exit Sub
'    rsPart.Find "[PARTNUMBER] LIKE '" + txtFind.Text + "%'", , adSearchForward, 1
'    If Not rsPart.EOF Then
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
        
        txtFind.Text = tmpRs!PartNumber
        txtFind.SelStart = Val(txtTag.Text)
        txtFind.SelLength = Len(txtFind.Text)
    End If

Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtFind_Change()")
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

Private Sub txtFlxDiscount_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFlxDiscount, KeyAscii, NumbersOnly)
    
    
    If KeyAscii = 13 Then
        
        If txtFlxDiscount.Text = "" Or txtFlxDiscount.Text = "0" Then
            flxSRP.TextMatrix(flxSRP.row, flxSRP.Col) = "0%"
            txtFlxDiscount.Text = 0
            Call SRPComputation
            txtFlxDiscount.Visible = False
            flxSRP.Col = flxSRP.Col
            flxSRP.row = flxSRP.row
            flxSRP.SetFocus
            Exit Sub
        
        Else
        
            If txtFlxDiscount.Text > 100 Then
                txtFlxDiscount.Text = ""
                txtFlxDiscount.Visible = False
                Exit Sub
            End If
            
            flxSRP.TextMatrix(flxSRP.row, flxSRP.Col) = txtFlxDiscount.Text & "%"
            
        End If
        
        
        txtFlxDiscount.Visible = False
        flxSRP.Col = flxSRP.Col
        flxSRP.SetFocus
    

    
    End If
End Sub

Private Sub txtFlxDiscount_LostFocus()
    txtFlxDiscount.Visible = False
End Sub

Private Sub txtFlxNewSRP_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFlxNewSRP, KeyAscii, NumbersOnly)
    
    If KeyAscii = 13 Then

        If txtFlxNewSRP.Text = "" Or txtFlxNewSRP.Text = "0" Then
            flxSRP.TextMatrix(flxSRP.row, 8) = ""
            flxSRP.TextMatrix(flxSRP.row, 9) = "0%"
            flxSRP.TextMatrix(flxSRP.row, 10) = "0%"
            txtFlxNewSRP.Visible = False
            flxSRP.Col = 7
            flxSRP.row = flxSRP.row
            flxSRP.SetFocus
            Exit Sub
        End If
        
        Call SRPComputation
        
        txtFlxNewSRP.Visible = False
        flxSRP.Col = 7
        If flxSRP.row = flxSRP.Rows - 1 Then
            flxSRP.row = flxSRP.row
        Else
            flxSRP.row = flxSRP.row + 1
        End If
        flxSRP.SetFocus
    
    End If
End Sub

Private Sub txtFlxNewSRP_LostFocus()
    txtFlxNewSRP.Visible = False
End Sub

Private Sub txtFlxNewWSP_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFlxNewWSP, KeyAscii, NumbersOnly)
    
    If KeyAscii = 13 Then
    
        If txtFlxNewWSP.Text = "" Or txtFlxNewWSP.Text = "0" Then
            flxSRP.TextMatrix(flxSRP.row, 8) = ""
            flxSRP.TextMatrix(flxSRP.row, 9) = "0%"
            flxSRP.TextMatrix(flxSRP.row, 10) = "0%"
            txtFlxNewWSP.Visible = False
            flxSRP.Col = 11
            flxSRP.row = flxSRP.row
            flxSRP.SetFocus
            Exit Sub
        End If

        Call WSComputation

        txtFlxNewWSP.Visible = False
        flxSRP.Col = 11
        If flxSRP.row = flxSRP.Rows - 1 Then
            flxSRP.row = flxSRP.row
        Else
            flxSRP.row = flxSRP.row + 1
        End If
        flxSRP.SetFocus
    
    End If
End Sub

Private Sub txtFlxNewWSP_LostFocus()
    txtFlxNewWSP.Visible = False
End Sub

Private Sub txtFlxPartNo_Change()
On Error GoTo er
    Dim Qry As String
    
    If chkManual = 0 Then Exit Sub
    If txtFlxPartNo.Text = "" Then Exit Sub
    
    If SearchcboPartNo Then
        SearchcboPartNo = False
        Exit Sub
    End If
    
    If tmpRs.State = adStateOpen Then tmpRs.Close
    
    'Qry = "SELECT SPAREPARTID,PARTNUMBER FROM mrcwarehouse.spareparts WHERE PARTNUMBER LIKE '" + txtFlxPartNo.Text + "%';"
    'InputBox Qry, Qry, Qry
    
    Qry = "SELECT sp.SPAREPARTID,sp.PARTNUMBER,sp.DESCRIPTION,sp.COST,sp.STOREPRICE" & _
          ",sp.DLRPRICE,sup.SUPPLIERCODE,sup.SUPPLIERNAME,'' as PONO, sp.MAXDISCOUNT, sp.WSMAXDISCOUNT " & _
          " FROM mrcwarehouse.spareparts as sp" & _
          " INNER JOIN mrcwarehouse.supplier as sup ON sp.SUPPLIERCODE = sup.SUPPLIERCODE" & _
          " WHERE sp.PARTNUMBER LIKE '" + txtFlxPartNo.Text + "%';"
    
    tmpRs.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic

    If tmpRs.RecordCount > 0 Then
        If Not SearchcboPartNo Then
            txtTag.Text = Len(txtFlxPartNo.Text)
            SearchcboPartNo = True
        End If
        
        If txtFlxPartNo.Text <> "" Then
            txtFlxPartNo.Text = tmpRs!PartNumber
        End If
        
        SupCode = tmpRs!SUPPLIERCODE
        flxSRP.TextMatrix(flxSRP.row, 2) = tmpRs!Description
        flxSRP.TextMatrix(flxSRP.row, 3) = Format(tmpRs!Cost, "#,##0.00")
        flxSRP.TextMatrix(flxSRP.row, 4) = Format(tmpRs!Cost, "#,##0.00")
        flxSRP.TextMatrix(flxSRP.row, 5) = Format(tmpRs!Cost * 1.12, "#,##0.00")
        flxSRP.TextMatrix(flxSRP.row, 6) = Format(tmpRs!STOREPRICE, "#,##0")
        flxSRP.TextMatrix(flxSRP.row, 11) = Val(tmpRs!MAXDISCOUNT) & "%"
        flxSRP.TextMatrix(flxSRP.row, 12) = Val(tmpRs!WSMAXDISCOUNT) & "%"
        flxSRP.TextMatrix(flxSRP.row, 13) = tmpRs!SPAREPARTID
        flxSRP.TextMatrix(flxSRP.row, 14) = tmpRs!SUPPLIERNAME
        
        txtFlxPartNo.SelStart = Val(txtTag.Text)
        txtFlxPartNo.SelLength = Len(txtFlxPartNo.Text)
    End If

Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtFlxPartNo_Change()")
End Sub

Private Sub txtFlxPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If chkManual = 0 Then Exit Sub
    
    If KeyCode = 8 Or KeyCode = 46 Then
        SearchcboPartNo = True
    ElseIf KeyCode = 13 Then
        Call txtFlxPartNo_KeyPress(13)
    End If
End Sub


Private Sub txtFlxPartNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        
        If chkManual = 0 Then Exit Sub
        
        If txtFlxPartNo.Text = "" Then
            flxSRP.TextMatrix(flxSRP.row, 1) = ""
            flxSRP.TextMatrix(flxSRP.row, 2) = ""
            flxSRP.TextMatrix(flxSRP.row, 3) = ""
            flxSRP.TextMatrix(flxSRP.row, 4) = ""
            flxSRP.TextMatrix(flxSRP.row, 5) = ""
            flxSRP.TextMatrix(flxSRP.row, 6) = ""
            flxSRP.TextMatrix(flxSRP.row, 7) = ""
            flxSRP.TextMatrix(flxSRP.row, 8) = ""
            flxSRP.TextMatrix(flxSRP.row, 9) = "0%"
            flxSRP.TextMatrix(flxSRP.row, 10) = "0%"
            flxSRP.TextMatrix(flxSRP.row, 11) = "0%"
            flxSRP.TextMatrix(flxSRP.row, 12) = "0%"
            flxSRP.TextMatrix(flxSRP.row, 13) = ""
            flxSRP.TextMatrix(flxSRP.row, 14) = ""
        End If
        
        flxSRP.TextMatrix(flxSRP.row, flxSRP.Col) = txtFlxPartNo.Text
        txtFlxPartNo.Visible = False
        SearchcboPartNo = False
        flxSRP.row = flxSRP.row
        flxSRP.Col = 1
        flxSRP.SetFocus
        
        DoEvents
        Call GetSupInfo(SupCode)
        
    End If
End Sub


Private Sub txtFlxPartNo_LostFocus()
    txtFlxPartNo.Visible = False
    If flxSRP.TextMatrix(DtlsRow, 1) = "" Then
        flxSRP.TextMatrix(DtlsRow, 1) = ""
        flxSRP.TextMatrix(DtlsRow, 2) = ""
        flxSRP.TextMatrix(DtlsRow, 3) = ""
        flxSRP.TextMatrix(DtlsRow, 4) = ""
        flxSRP.TextMatrix(DtlsRow, 5) = ""
        flxSRP.TextMatrix(DtlsRow, 6) = ""
        flxSRP.TextMatrix(DtlsRow, 7) = ""
        flxSRP.TextMatrix(DtlsRow, 8) = ""
        flxSRP.TextMatrix(DtlsRow, 9) = "0%"
        flxSRP.TextMatrix(DtlsRow, 10) = "0%"
        flxSRP.TextMatrix(DtlsRow, 11) = "0%"
        flxSRP.TextMatrix(DtlsRow, 12) = "0%"
        flxSRP.TextMatrix(DtlsRow, 13) = ""
        flxSRP.TextMatrix(DtlsRow, 14) = ""
        
        SearchcboPartNo = False
    End If
    If flxSRP.TextMatrix(DtlsRow, 1) <> "" And chkManual.Value = 1 Then
        Call mnuSTAddRow_Click
    End If
End Sub


Private Sub txtFlxProfMar_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFlxProfMar, KeyAscii, NumbersWithDecimalPoint)
    
    If KeyAscii = 13 Then
        
       If flxSRP.Col = 7 Then
            If txtFlxProfMar.Text = "" Or txtFlxProfMar.Text = "0" Then
            flxSRP.TextMatrix(flxSRP.row, 8) = ""
            flxSRP.TextMatrix(flxSRP.row, 9) = "0%"
            flxSRP.TextMatrix(flxSRP.row, 10) = "0%"
                txtFlxProfMar.Visible = False
                flxSRP.Col = DtlsCol
                flxSRP.SetFocus
                Exit Sub
            End If
        ElseIf flxSRP.Col = 12 Then
            If txtFlxProfMar.Text = "" Or txtFlxProfMar.Text = "0" Then
                flxSRP.TextMatrix(flxSRP.row, 11) = ""
                flxSRP.TextMatrix(flxSRP.row, 12) = ""
                flxSRP.TextMatrix(flxSRP.row, 13) = ""
                txtFlxProfMar.Visible = False
                flxSRP.Col = DtlsCol
                flxSRP.SetFocus
                Exit Sub
            End If
        End If
    
        flxSRP.TextMatrix(flxSRP.row, flxSRP.Col) = Format(txtFlxProfMar.Text, "#,##0.00")
        
       If flxSRP.Col = 7 Then
                
                Call SRPComputation
            
       ElseIf flxSRP.Col = 12 Then

                Call WSComputation
       End If
        
        txtFlxProfMar.Visible = False
        flxSRP.Col = flxSRP.Col
        If flxSRP.row = flxSRP.Rows - 1 Then
            flxSRP.row = flxSRP.row
        Else
            flxSRP.row = flxSRP.row + 1
        End If
        flxSRP.SetFocus
    
    End If
End Sub

Private Sub txtFlxProfMar_LostFocus()
    txtFlxProfMar.Visible = False
    txtFlxProfMar.Text = ""
End Sub

Private Sub txtFlxValuePer_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFlxValuePer, KeyAscii, NumbersWithDecimalPoint)
    
    If KeyAscii = 13 Then
    
        If flxSRP.Col = 10 Then
            If txtFlxValuePer.Text = "" Or txtFlxValuePer.Text = "0" Then
                flxSRP.TextMatrix(flxSRP.row, 8) = ""
                flxSRP.TextMatrix(flxSRP.row, 9) = "0%"
                flxSRP.TextMatrix(flxSRP.row, 10) = "0%"
                txtFlxProfMar.Visible = False
                flxSRP.Col = DtlsCol
                flxSRP.SetFocus
                Exit Sub
            End If
'        ElseIf flxSRP.Col = 13 Then
'            If txtFlxValuePer.Text = "" Or txtFlxValuePer.Text = "0" Then
'                flxSRP.TextMatrix(flxSRP.row, 11) = ""
'                flxSRP.TextMatrix(flxSRP.row, 12) = ""
'                flxSRP.TextMatrix(flxSRP.row, 13) = ""
'                txtFlxProfMar.Visible = False
'                flxSRP.Col = DtlsCol
'                flxSRP.SetFocus
'                Exit Sub
'            End If
        End If
    
        If flxSRP.Col = 10 Then
            Call SRPComputation
'        ElseIf flxSRP.Col = 13 Then
'            Call WSComputation
        End If
        
        txtFlxValuePer.Visible = False
        flxSRP.Col = flxSRP.Col
        If flxSRP.row = flxSRP.Rows - 1 Then
            flxSRP.row = flxSRP.row
        Else
            flxSRP.row = flxSRP.row + 1
        End If
        flxSRP.SetFocus
       
    End If
End Sub

Private Sub txtFlxValuePer_LostFocus()
    txtFlxValuePer.Visible = False
    txtFlxValuePer.Text = ""
End Sub

Private Sub txtFlxWDiscount_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFlxWDiscount, KeyAscii, NumbersOnly)
    
    
    If KeyAscii = 13 Then
        
        If txtFlxWDiscount.Text = "" Or txtFlxWDiscount.Text = "0" Then
            flxSRP.TextMatrix(flxSRP.row, flxSRP.Col) = "0%"
            txtFlxWDiscount.Text = 0
            Call SRPComputation
            txtFlxWDiscount.Visible = False
            flxSRP.Col = flxSRP.Col
            flxSRP.row = flxSRP.row
            flxSRP.SetFocus
            Exit Sub
        
        Else
        
            If txtFlxWDiscount.Text > 100 Then
                txtFlxWDiscount.Text = ""
                txtFlxWDiscount.Visible = False
                Exit Sub
            End If
            
            flxSRP.TextMatrix(flxSRP.row, flxSRP.Col) = txtFlxWDiscount.Text & "%"
            
        End If
        
        txtFlxWDiscount.Visible = False
        flxSRP.Col = flxSRP.Col
        flxSRP.SetFocus
    
    End If
End Sub

Private Sub txtFlxWDiscount_LostFocus()
txtFlxWDiscount.Visible = False
End Sub

Private Sub txtPckngNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        
        If chkManual.Value = 1 Then
            flxSRP.SetFocus
        Else
            DoEvents
            Call GetRecord(txtPckngNo.Text)
        End If
    End If
End Sub


Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub


