VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmPriceFactor 
   Caption         =   "PRICE FACTOR"
   ClientHeight    =   10455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10455
   ScaleWidth      =   15255
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrSMSQTYLSI 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8880
      Top             =   0
   End
   Begin VB.TextBox txtTag 
      Height          =   285
      Left            =   15600
      TabIndex        =   59
      Text            =   "Text1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer tmr0SRP 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8280
      Top             =   0
   End
   Begin lvButton.lvButtons_H cmd0SRP 
      Height          =   375
      Left            =   3120
      TabIndex        =   54
      Top             =   120
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "Zero SRP"
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
      Image           =   "frmPriceFactor.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdRefresh 
      Height          =   375
      Left            =   1680
      TabIndex        =   51
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
      Image           =   "frmPriceFactor.frx":059A
      cBack           =   -2147483633
   End
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7080
      Top             =   0
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   15480
      Locked          =   -1  'True
      TabIndex        =   45
      Text            =   "txtHoldGrayBackColor"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   15480
      TabIndex        =   44
      Text            =   "HoldBackColor"
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   120
      TabIndex        =   40
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   " E&XIT FORM"
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
      Image           =   "frmPriceFactor.frx":0B34
      cBack           =   -2147483633
   End
   Begin VB.Frame SPpricing 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Tag             =   "XYF"
      Top             =   480
      Width           =   15015
      Begin VB.TextBox txtFindPart 
         Height          =   375
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   58
         Top             =   240
         Width           =   3615
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Left            =   11400
         TabIndex        =   55
         Tag             =   "XYF"
         Top             =   600
         Width           =   2055
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "All Suppliers"
         Height          =   315
         Left            =   13680
         TabIndex        =   53
         Tag             =   "XYF"
         Top             =   600
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   6240
         TabIndex        =   52
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
         Image           =   "frmPriceFactor.frx":6326
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cboSupplier 
         Height          =   315
         Left            =   11400
         TabIndex        =   38
         Tag             =   "XYF"
         Top             =   240
         Width           =   3495
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Left            =   2210
         TabIndex        =   2
         Top             =   270
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxFind 
         Height          =   4455
         Left            =   120
         TabIndex        =   57
         Tag             =   "XYF"
         Top             =   960
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   7858
         _Version        =   393216
         Cols            =   10
         RowHeightMin    =   300
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "           "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   10
      End
      Begin VB.Label Label1 
         Caption         =   "TYPE:"
         Height          =   255
         Left            =   10440
         TabIndex        =   56
         Tag             =   "XYF"
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "SUPPLIER:"
         Height          =   255
         Left            =   10440
         TabIndex        =   39
         Tag             =   "XYF"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblFindParts 
         Caption         =   "FIND PART DESCRIPTION:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   285
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4335
      Left            =   9000
      TabIndex        =   4
      Tag             =   "XYF"
      Top             =   6000
      Width           =   6135
      Begin MSFlexGridLib.MSFlexGrid flxSRPPriceHis 
         Height          =   1695
         Left            =   120
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   480
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2990
         _Version        =   393216
         Cols            =   7
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxWSPriceHis 
         Height          =   1575
         Left            =   120
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   2640
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2778
         _Version        =   393216
         Cols            =   7
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         AllowUserResizing=   1
      End
      Begin VB.Label Label3 
         Caption         =   "WHOLESALE PRICE HISTORY"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   2400
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "SRP PRICE HISTORY"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4335
      Left            =   120
      TabIndex        =   3
      Tag             =   "XYF"
      Top             =   6000
      Width           =   8775
      Begin VB.TextBox txtNetofVat 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   7080
         TabIndex        =   16
         Tag             =   "XYF"
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtRefDate 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   17
         Tag             =   "XYF"
         Top             =   540
         Width           =   1455
      End
      Begin lvButton.lvButtons_H cmdTrans 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   240
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
         Image           =   "frmPriceFactor.frx":BB18
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   2655
         Left            =   120
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   1560
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4683
         _Version        =   393216
         TabOrientation  =   1
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         TabCaption(0)   =   "     ."
         TabPicture(0)   =   "frmPriceFactor.frx":C0B2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame5"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         Begin VB.Frame Frame5 
            Height          =   2415
            Left            =   4200
            TabIndex        =   30
            Tag             =   "XYF"
            Top             =   120
            Width           =   4095
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "   PROFIT RATIO"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   50
               Tag             =   "XYF"
               Top             =   720
               Width           =   1575
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "PROFIT MARGIN"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   36
               Tag             =   "XYF"
               Top             =   1160
               Width           =   1575
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "           FIX PRICE"
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   35
               Tag             =   "XYF"
               Top             =   1600
               Width           =   1575
            End
            Begin VB.TextBox txtWS 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   1730
               Locked          =   -1  'True
               TabIndex        =   34
               Tag             =   "XYF"
               Text            =   "0"
               Top             =   240
               Width           =   1935
            End
            Begin VB.TextBox txtWSProfRatio 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   1730
               Locked          =   -1  'True
               TabIndex        =   33
               Tag             =   "XYF"
               Text            =   "0"
               Top             =   720
               Width           =   1935
            End
            Begin VB.TextBox txtWSProfMar 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   1730
               Locked          =   -1  'True
               TabIndex        =   32
               Tag             =   "XYF"
               Text            =   "0"
               Top             =   1160
               Width           =   1935
            End
            Begin VB.TextBox txtWSFixPrice 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   1730
               Locked          =   -1  'True
               TabIndex        =   31
               Tag             =   "XYF"
               Text            =   "0"
               Top             =   1600
               Width           =   1935
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               Caption         =   "%"
               Height          =   255
               Left            =   3680
               TabIndex        =   48
               Tag             =   "XYF"
               Top             =   780
               Width           =   135
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               Caption         =   "CURRENT WSP:"
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
               Left            =   240
               TabIndex        =   37
               Tag             =   "XYF"
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.Frame Frame4 
            Height          =   2415
            Left            =   120
            TabIndex        =   20
            Tag             =   "XYF"
            Top             =   120
            Width           =   3975
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "   PROFIT RATIO"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   49
               Tag             =   "XYF"
               Top             =   720
               Width           =   1575
            End
            Begin VB.TextBox txtSRPMaxDis 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   1725
               Locked          =   -1  'True
               TabIndex        =   29
               Tag             =   "XYF"
               Text            =   "0"
               Top             =   2000
               Width           =   1815
            End
            Begin VB.TextBox txtSRPFixPrice 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   1730
               Locked          =   -1  'True
               TabIndex        =   27
               Tag             =   "XYF"
               Text            =   "0"
               Top             =   1560
               Width           =   1815
            End
            Begin VB.TextBox txtSRPProfMar 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   1730
               Locked          =   -1  'True
               TabIndex        =   26
               Tag             =   "XYF"
               Text            =   "0"
               Top             =   1160
               Width           =   1815
            End
            Begin VB.TextBox txtSRPProfRatio 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   1730
               Locked          =   -1  'True
               TabIndex        =   25
               Tag             =   "XYF"
               Text            =   "0"
               Top             =   720
               Width           =   1815
            End
            Begin VB.TextBox txtSRP 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   1730
               Locked          =   -1  'True
               TabIndex        =   24
               Tag             =   "XYF"
               Text            =   "0"
               Top             =   240
               Width           =   1815
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "           FIX PRICE"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   23
               Tag             =   "XYF"
               Top             =   1600
               Width           =   1575
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "PROFIT MARGIN"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   22
               Tag             =   "XYF"
               Top             =   1160
               Width           =   1575
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               Caption         =   "%"
               Height          =   255
               Left            =   3550
               TabIndex        =   47
               Tag             =   "XYF"
               Top             =   780
               Width           =   135
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Caption         =   "%"
               Height          =   255
               Left            =   3550
               TabIndex        =   46
               Tag             =   "XYF"
               Top             =   2040
               Width           =   135
            End
            Begin VB.Label Label11 
               Caption         =   " MAX DISCOUNT:"
               Height          =   255
               Left            =   120
               TabIndex        =   28
               Tag             =   "XYF"
               Top             =   1995
               Width           =   1335
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               Caption         =   "CURRENT SRP:"
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
               TabIndex        =   21
               Tag             =   "XYF"
               Top             =   240
               Width           =   1575
            End
         End
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   15
         Tag             =   "XYF"
         Top             =   1200
         Width           =   7095
      End
      Begin VB.TextBox txtRefNo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   13
         Tag             =   "XYF"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtPartNo 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   840
         Width           =   3135
      End
      Begin lvButton.lvButtons_H cmdTrans 
         Height          =   375
         Index           =   2
         Left            =   1320
         TabIndex        =   42
         Top             =   240
         Visible         =   0   'False
         Width           =   1080
         _ExtentX        =   1905
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
         Image           =   "frmPriceFactor.frx":C0CE
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdTrans 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Visible         =   0   'False
         Width           =   1080
         _ExtentX        =   1905
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
         Image           =   "frmPriceFactor.frx":C668
         cBack           =   -2147483633
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "REF. DATE:"
         Height          =   255
         Left            =   5400
         TabIndex        =   18
         Tag             =   "XYF"
         Top             =   555
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "REF. NO.:"
         Height          =   255
         Left            =   5400
         TabIndex        =   14
         Tag             =   "XYF"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "NET OF VAT:"
         Height          =   255
         Left            =   5400
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   870
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "DESCRIPTION:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "PART NUMBER:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   870
         Width           =   1455
      End
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   7680
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   15375
      m_MDI_Height    =   10935
      OriginalBorder  =   2
      OriginalCaption =   25
   End
   Begin VB.Menu mnuPriceUpdater 
      Caption         =   "PRICEUPDATER"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmPriceFactor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ans As VbMsgBoxResult
Dim AllArrSupplier() As String
Dim AllArrClass() As String
Dim AllArrPartNo() As String
Dim AllArrDesc() As String
Dim arrSupplier() As String
Dim arrClass() As String
Dim arrPartNo() As String
Dim arrDesc() As String
Dim rsParts As New ADODB.Recordset

Dim ActRow As Integer, fRow As Integer, fCol As Integer, TCol As Integer, AccID As Integer, ColPrn As Integer

Dim SearchcboSup As Boolean
Dim SearchcboType As Boolean
Dim SearchPartNo As Boolean
Dim SupCode As String
Dim pSupCode As String, chkSup As Boolean, PrevSup As String, pClass As String, IsSelect As Boolean
Dim didRecordEntry As Boolean
Dim i As Integer, j As Integer
Dim SPid As Integer

Dim PartDesc As String
Dim flxrow As Integer
Dim bNone As Boolean
Sub getSRPProfMarginRatio()
On Error GoTo er

    Dim ProMar As Double
    Dim CostVat As Double
    Dim NewSRP As Double
    Dim Ratio As Double
    
    If txtNetofVat.Text = 0 Or txtSRP.Text = 0 Then Exit Sub
    
    NewSRP = Val(Replace(Trim(txtSRP.Text), ",", ""))
    'CostVat = Format(txtNetofVat.Text, "#,#00")
    CostVat = Val(Replace(txtNetofVat, ",", ""))
    CostVat = CostVat * 1.12
    CostVat = Format(CostVat, "#,#00")
    
    
    ProMar = NewSRP - CostVat
    
    'Ratio = ProMar / NewSRP
    Ratio = ProMar / CostVat
    Ratio = Ratio * 100
    
    txtSRPFixPrice.Text = Format(NewSRP, "#,#00")
    txtSRPProfMar.Text = Format(ProMar, "#,#00")
    
'    If Ratio > 100 Then
'        Ratio = 100
'    End If
    
    txtSRPProfRatio.Text = Format(Ratio, "#,#00.00")
        
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->getSRPProfMarginRatio()")
End Sub

Sub getWSProfMarRatio()
On Error GoTo er

    Dim ProMar As Double
    Dim CostVat As Double
    Dim NewWSP As Double
    Dim Ratio As Double
    
    If txtNetofVat.Text = 0 Or txtWS.Text = 0 Then Exit Sub
    
    NewWSP = Val(Replace(Trim(txtWS.Text), ",", ""))
    'CostVat = Format(txtNetofVat.Text, "#,#00")
    CostVat = Val(Replace(txtNetofVat, ",", ""))
    CostVat = CostVat * 1.12
    CostVat = Format(CostVat, "#,#00")
    
    ProMar = NewWSP - CostVat

    'Ratio = ProMar / NewWSP
    Ratio = ProMar / CostVat
    Ratio = Ratio * 100
    
    txtWSFixPrice.Text = Format(NewWSP, "#,#00")
    txtWSProfMar.Text = Format(ProMar, "#,#00")
    
'    If Ratio > 100 Then
'        Ratio = 100
'    End If
    
    txtWSProfRatio.Text = Format(Ratio, "#,#00.00")
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->getWSProfMarRatio()")
End Sub

Private Sub cbotype_Change()
On Error GoTo er
    Dim i As Integer, j As Integer
    Dim strPartial As String, strTotal As String
    Dim X As Integer


    If SearchcboType Then
        SearchcboType = False
        Exit Sub
    End If


    With cboType
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters

        If i <> -1 And i <= .ListCount Then
            strTotal = .List(i)

            DoEvents
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
               SearchcboType = True
               .SelText = Right$(strTotal, j)
               .SelStart = Len(strPartial)
               .SelLength = j
            End If

        End If

    End With

Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error ->cboType_Change()")
End Sub

Private Sub cboType_Click()
    
    Screen.MousePointer = vbHourglass
    tmrRefresh.Enabled = True
    DoEvents
    Call Load(frm_OnProcess2)
    frm_OnProcess2.Show vbModal

End Sub


Private Sub cboType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Or KeyCode = 46 Then
        SearchcboType = True
    End If
End Sub


Private Sub cboType_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboType, KeyAscii, ValidChars)
    
    If KeyAscii = 13 Then

        pClass = cboType.Text
        Screen.MousePointer = vbHourglass
        tmrRefresh.Enabled = True
        DoEvents
        Call Load(frm_OnProcess2)
        frm_OnProcess2.Show vbModal
        
    End If
End Sub

Private Sub cboSupplier_Change()
On Error GoTo er
    Dim i As Integer, j As Integer
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If SearchcboSup Then
        SearchcboSup = False
        Exit Sub
    End If

    With cboSupplier
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters

        If i <> -1 And i <= .ListCount Then
            strTotal = .List(i)
            
            SupCode = arrSupplier(i, 0)
    
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
    Call modObjTools.ErrorMessage(Err.Description, "Error ->cboSupplier_Change()")
End Sub

Private Sub cboSupplier_Click()
    
    SupCode = arrSupplier(cboSupplier.ListIndex, 0)

        
    chkAll.Value = 0
    
    If IsSelect = True Then: IsSelect = False: Exit Sub
    
    Screen.MousePointer = vbHourglass
    tmrRefresh.Enabled = True
    DoEvents
    Call Load(frm_OnProcess2)
    frm_OnProcess2.Show vbModal
        
    
End Sub

Private Sub cboSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Or KeyCode = 46 Then
        SearchcboSup = True
    End If
End Sub

Private Sub cboSupplier_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboSupplier, KeyAscii, ValidChars)
    
    If KeyAscii = 13 Then
        If cboSupplier.Text = "" Then Exit Sub
        
        chkAll.Value = 0
        
        If IsSelect = True Then: IsSelect = False: Exit Sub
        
        Screen.MousePointer = vbHourglass
        tmrRefresh.Enabled = True
        DoEvents
        frm_OnProcess2.OwnerForm = "SRP Updater"
        Call Load(frm_OnProcess2)
        frm_OnProcess2.Show vbModal
        
    End If
End Sub




Private Sub Check1_Click(Index As Integer)
On Error GoTo er
    
    If didRecordEntry = False Then
        Check1(0).Value = 0
        Check1(1).Value = 0
        Check1(2).Value = 0
        Check1(3).Value = 0
        Check1(4).Value = 0
        Check1(5).Value = 0
        Exit Sub
    End If

    If Index = 0 Then ' Check box SRP Prof Ratio
        
        If Check1(0).Value = 1 Then
            txtSRPProfRatio.Locked = False
            txtSRPProfRatio.BackColor = txtHoldWhiteBackColor.BackColor
            
            If txtSRPProfRatio.Text <> "" Then
                txtSRPProfRatio.SelStart = 0
                txtSRPProfRatio.SelLength = Len(txtSRPProfRatio.Text)
                txtSRPProfRatio.SetFocus
            End If
            
            Check1(1).Value = 0
            Check1(2).Value = 0
        Else
            txtSRPProfRatio.Locked = True
            txtSRPProfRatio.BackColor = txtHoldGrayBackColor.BackColor
        End If

    ElseIf Index = 1 Then ' Check box SRP Prof Margin
    
        If Check1(1).Value = 1 Then
            txtSRPProfMar.Locked = False
            txtSRPProfMar.BackColor = txtHoldWhiteBackColor.BackColor
            
            If txtSRPProfMar.Text <> "" Then
                txtSRPProfMar.SelStart = 0
                txtSRPProfMar.SelLength = Len(txtSRPProfMar.Text)
                txtSRPProfMar.SetFocus
            End If
            
            Check1(0).Value = 0
            Check1(2).Value = 0
        Else
            txtSRPProfMar.Locked = True
            txtSRPProfMar.BackColor = txtHoldGrayBackColor.BackColor
        End If
        
    ElseIf Index = 2 Then ' Check box SRP Fix Price
        
        If Check1(2).Value = 1 Then
            txtSRPFixPrice.Locked = False
            txtSRPFixPrice.BackColor = txtHoldWhiteBackColor.BackColor
            
            If txtSRPFixPrice.Text <> "" Then
                txtSRPFixPrice.SelStart = 0
                txtSRPFixPrice.SelLength = Len(txtSRPFixPrice.Text)
                txtSRPFixPrice.SetFocus
            End If
            
            Check1(0).Value = 0
            Check1(1).Value = 0
        Else
            txtSRPFixPrice.Locked = True
            txtSRPFixPrice.BackColor = txtHoldGrayBackColor.BackColor
        End If
    
    ElseIf Index = 3 Then ' Check box WS Prof Ratio
        
        If Check1(3).Value = 1 Then
            txtWSProfRatio.Locked = False
            txtWSProfRatio.BackColor = txtHoldWhiteBackColor.BackColor
            
            If txtWSProfRatio.Text <> "" Then
                txtWSProfRatio.SelStart = 0
                txtWSProfRatio.SelLength = Len(txtWSProfRatio.Text)
                txtWSProfRatio.SetFocus
            End If
            
            Check1(4).Value = 0
            Check1(5).Value = 0
        Else
            txtWSProfRatio.Locked = True
            txtWSProfRatio.BackColor = txtHoldGrayBackColor.BackColor
        End If

    ElseIf Index = 4 Then ' Check box WS Prof Margin
        
        If Check1(4).Value = 1 Then
            txtWSProfMar.Locked = False
            txtWSProfMar.BackColor = txtHoldWhiteBackColor.BackColor
            
            If txtWSProfMar.Text <> "" Then
                txtWSProfMar.SelStart = 0
                txtWSProfMar.SelLength = Len(txtWSProfMar.Text)
                txtWSProfMar.SetFocus
            End If
            
            Check1(3).Value = 0
            Check1(5).Value = 0
        Else
            txtWSProfMar.Locked = True
            txtWSProfMar.BackColor = txtHoldGrayBackColor.BackColor
        End If
        
    ElseIf Index = 5 Then ' Check box WS Fix Price
         
        If Check1(5).Value = 1 Then
            txtWSFixPrice.Locked = False
            txtWSFixPrice.BackColor = txtHoldWhiteBackColor.BackColor
            
            If txtWSFixPrice.Text <> "" Then
                txtWSFixPrice.SelStart = 0
                txtWSFixPrice.SelLength = Len(txtWSFixPrice.Text)
                txtWSFixPrice.SetFocus
            End If
            
            Check1(3).Value = 0
            Check1(4).Value = 0
        Else
            txtWSFixPrice.Locked = True
            txtWSFixPrice.BackColor = txtHoldGrayBackColor.BackColor
        End If
        
    End If

Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error ->Check1_Click")
End Sub

Private Sub chkAll_Click()
    
    If chkAll.Value = 1 Then
        PrevSup = cboSupplier.Text
        chkSup = True: cboSupplier.Text = ""
        SupCode = ""
        IsSelect = False
        
        Screen.MousePointer = vbHourglass
        tmrRefresh.Enabled = True
        DoEvents
        Call Load(frm_OnProcess2)
        frm_OnProcess2.Show vbModal
        
    Else
    
        chkSup = False
        IsSelect = True
        If PrevSup <> "" Then cboSupplier.Text = PrevSup
        
        pSupCode = SupCode
        
        Screen.MousePointer = vbHourglass
        tmrRefresh.Enabled = True
        DoEvents
        Call Load(frm_OnProcess2)
        frm_OnProcess2.Show vbModal
        
    End If
    
End Sub


Private Sub FindZeroSRP()
    Dim Qry As String
    
    If rsParts.State = adStateOpen Then rsParts.Close
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format([  COST],'Standard') AS [UNIT COST], Format([STOREPRICE],'Standard') AS [SRP], SPAREPARTS.SPAREPARTID, SPAREPARTS.PRICEHIST, SPAREPARTS.PRICEREF, SPAREPARTS.SPFID, Format([DLRPRICE],'Standard') AS [DP] " & _
          "From SPAREPARTS WHERE (((SPAREPARTS.STOREPRICE)=0)) ORDER BY SPAREPARTS.PARTNUMBER;"
    
    'mySQL
'        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION " & _
'          ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND',ifnull(sptr.SixMonSale,0) as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'          " From SPAREPARTS as sp LEFT JOIN (select sum(`out`) as SixMonSale,sparepartid from spitr where  transdate >= DATE_SUB(now(), INTERVAL 6 MONTH) and transdate <= Now() and transtype='OUT'" & _
'          " and remarks<>'TRANSFER' GROUP by sparepartid) as sptr ON sp.sparepartid = sptr.sparepartid WHERE sp.STOREPRICE=0 ORDER BY sp.PARTNUMBER;"
    
         Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION " & _
          ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN',sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
          " From SPAREPARTS as sp WHERE sp.STOREPRICE=0 ORDER BY sp.PARTNUMBER;"
          
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsParts.RecordCount = 0 Then
        flxFind.Clear
        flxFind.Rows = 2
        flxFind.FormatString = "|<PART NUMBER |<DESCRIPTION  |^NET OF VAT  |>SRP  |>WSP  |>QTY. |>SIX MSQ |>BACK ORDER |>LAST DATE IN |<APPLICABLE MODEL"
 
    Else
        Set flxFind.DataSource = rsParts
        flxFind.Refresh
        flxFind.ColAlignment(3) = flexAlignRightCenter
        flxFind.ColAlignment(4) = flexAlignRightCenter
        flxFind.ColAlignment(5) = flexAlignRightCenter
    End If
    'format cur
    Call FormatCur
    Call flxFind_Click

End Sub

Private Sub GetSMSQTYLDI(pSPid As Integer)
On Error GoTo er
    Dim tmpRs As New ADODB.Recordset
    
    Call modSRP.sp_get_SMSQTY_LDI(pSPid, tmpRs)
    If tmpRs.State = 0 Then Exit Sub
    
    If tmpRs.RecordCount <= 0 Then
        tmpRs.Close
        Exit Sub
    End If


    tmpRs.MoveFirst

        While Not tmpRs.EOF
                flxFind.TextMatrix(flxFind.row, 7) = tmpRs!SMSQTY
                flxFind.TextMatrix(flxFind.row, 9) = Format(tmpRs!TRANSDATE, "mm/dd/yyyy")
        tmpRs.MoveNext
        
    Wend
    
    If tmpRs.State <> 0 Then tmpRs.Close
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->GetSMSQTYLDI()")
End Sub
Private Sub cmd0SRP_Click()
    
    Screen.MousePointer = vbHourglass
    tmr0SRP.Enabled = True
    DoEvents
    Call Load(frm_OnProcess2)
    frm_OnProcess2.Show vbModal

End Sub

Private Sub cmdExit_Click()
    
    Ans = MsgBox("ARE YOU SURE YOU WANT TO EXIT?", vbQuestion + vbYesNo, Me.Caption)
    If Ans = vbNo Then Exit Sub
    
    Unload Me
End Sub

Private Sub cmdFind_Click()
On Error GoTo er
        
        
        cmdTrans(0).Enabled = False
        
        flxSRPPriceHis.Rows = 1
        flxSRPPriceHis.Rows = 2
        flxWSPriceHis.Rows = 1
        flxWSPriceHis.Rows = 2
        
        Call ClearEntryForm
        
        txtSRPProfRatio.Text = ""
        txtSRPProfMar.Text = ""
        txtSRPFixPrice.Text = ""
        txtSRPMaxDis.Text = ""
        
        txtWSProfRatio.Text = ""
        txtWSProfMar.Text = ""
        txtWSFixPrice.Text = ""
    
        Screen.MousePointer = vbHourglass
        tmrRefresh.Enabled = True
        DoEvents
        Call Load(frm_OnProcess2)
        frm_OnProcess2.Show vbModal
        
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error ->cmdFind_Click")
End Sub

Private Sub cmdRefresh_Click()

    If cboSupplier.Text = "" Then SupCode = 0
    
    Screen.MousePointer = vbHourglass
    tmrRefresh.Enabled = True
    DoEvents
    Call Load(frm_OnProcess2)
    frm_OnProcess2.Show vbModal
    
End Sub

Private Sub cmdTrans_Click(Index As Integer)
On Error GoTo er
'cmdTrans Legend
'0 ->EDIT
'1 ->SAVE
'2 ->CANCEL

    Dim Level As String
    Dim message As String
    
    If Index = 0 Then 'Edit
        Ans = MsgBox("Do you want to edit price?", vbYesNo + vbInformation, Me.Caption)
        If Ans = vbNo Then Exit Sub
        
        flxrow = flxFind.row
        
        cmdTrans(0).Visible = False
        cmdTrans(1).Visible = True
        cmdTrans(2).Visible = True
                
        txtRefNo.Locked = False
        txtRefNo.BackColor = txtHoldWhiteBackColor.BackColor
        txtRefNo.SetFocus

        txtNetofVat.Locked = False
        txtNetofVat.BackColor = txtHoldWhiteBackColor.BackColor
        
        txtRefDate.Locked = False
        txtRefDate.Text = Format(Now(), "mm / dd / yyyy")
        txtRefDate.Locked = True
        
        txtSRPMaxDis.Locked = False
        txtSRPMaxDis.BackColor = txtHoldWhiteBackColor.BackColor
        
        cmdExit.Enabled = False
        cmdRefresh.Enabled = False
        cmd0SRP.Enabled = False
        SPpricing.Enabled = False
        
        didRecordEntry = True
        
    ElseIf Index = 1 Then 'SAVE
        Dim maxdis As Integer
        Dim SRP As String
        Dim wsp As String
        
        If ValidateEntry() <> "" Then
            MsgBox ValidateEntry, vbExclamation, Me.Caption
            Exit Sub
        End If
        
        Ans = MsgBox("DO YOU WANT TO SAVE CHANGE?", vbQuestion + vbYesNo, Me.Caption)
        If Ans = vbNo Then Exit Sub
            
       Screen.MousePointer = vbHourglass
            
            If txtSRPMaxDis.Text <> "" Then
                maxdis = txtSRPMaxDis.Text
            End If
            
            If Val(Replace(txtSRPFixPrice.Text, ",", "")) < 1 Then ' if no srp
                SRP = Val(Replace(txtSRP.Text, ",", ""))
            Else
                SRP = Val(Replace(txtSRPFixPrice.Text, ",", ""))
            End If
            
            If Val(Replace(txtWSFixPrice.Text, ",", "")) < 1 Then 'If no wsp
                wsp = Val(Replace(txtWS.Text, ",", ""))
            Else
                wsp = Val(Replace(txtWSFixPrice.Text, ",", ""))
            End If
            
            Call modSRP.sp_priceupdate_save(Trim(SPid), Trim(txtRefNo.Text), Val(Replace(txtNetofVat.Text, ",", "")) _
                                            , SRP, Format(txtSRPProfMar.Text, "#,##0") _
                                            , txtSRPProfRatio.Text & "%", wsp _
                                            , Format(txtWSProfMar.Text, "#,##0"), txtWSProfRatio.Text & "%", maxdis, Level, message)
    
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

        cmdTrans(0).Visible = True
        cmdTrans(1).Visible = False
        cmdTrans(2).Visible = False
        
        cmdTrans(0).Enabled = False
        
        txtSRP.Text = Format(SRP, "#,##0.00")
        txtWS.Text = Format(wsp, "#,##0.00")
        
        flxFind.TextMatrix(flxrow, 4) = Format(SRP, "#,##0.00")
        flxFind.TextMatrix(flxrow, 5) = Format(wsp, "#,##0.00")
        
        Call LockedEntryObject(False)
        
        DoEvents
        Call LoadWSHist(SPid)
        DoEvents
        Call LoadSRPHist(SPid)
        
        cmdExit.Enabled = True
        cmdRefresh.Enabled = True
        cmd0SRP.Enabled = True
        SPpricing.Enabled = True
        
        didRecordEntry = False
    
    ElseIf Index = 2 Then 'CANCEL
        Ans = MsgBox("Cancel the currrent operation?", vbYesNo + vbQuestion, Me.Caption)
        If Ans = vbNo Then Exit Sub

        cmdTrans(0).Visible = True
        cmdTrans(1).Visible = False
        cmdTrans(2).Visible = False
        cmdTrans(0).Enabled = False
    
        Call ClearEntryForm
        Call LockedEntryObject(False)
        
        cmdExit.Enabled = True
        cmdRefresh.Enabled = True
        cmd0SRP.Enabled = True
        SPpricing.Enabled = True
        
        didRecordEntry = False
        flxrow = 0
        
    End If

Exit Sub
er:
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdTrans_Click()")
End Sub

Function ValidateEntry() As String
    Dim vMsg As String, Qry As String
    Dim i As Integer, Cntr As Integer
    
    vMsg = ""
    
    If txtRefNo.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field ref. no."
    If txtSRPFixPrice.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field SRP fix price."
    If txtWSFixPrice.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field WS fix price."
              
    ValidateEntry = vMsg
    
End Function

Private Sub flxFind_Click()

    If didRecordEntry = True Then Exit Sub
    
    If flxFind.TextMatrix(flxFind.row, 1) = "" Then Exit Sub

    'flxFind.HighLight = flexHighlightWithFocus
    
    Screen.MousePointer = vbHourglass
    
    txtRefNo.Text = ""
    
'0       1             2            3           4      5       6      7         8              9                10              11          12
'"|<PART NUMBER |<DESCRIPTION  |^NET OF VAT  |>SRP  |>WSP  |>QTY. |>SIX MSQ |>BACK ORDER |>LAST DATE IN |<APPLICABLE MODEL |SPAREPARTS |MAX DISCOUNT"
    
    txtRefDate.Text = Format(Now(), "mm/dd/YYYY")
    txtPartNo.Text = flxFind.TextMatrix(flxFind.row, 1)
    txtDesc.Text = flxFind.TextMatrix(flxFind.row, 2)
    txtNetofVat.Text = Format(flxFind.TextMatrix(flxFind.row, 3), "#,##0.00")
    txtSRP.Text = Format(flxFind.TextMatrix(flxFind.row, 4), "#,##0.00")
    txtWS.Text = Format(flxFind.TextMatrix(flxFind.row, 5), "#,##0.00")
    SPid = flxFind.TextMatrix(flxFind.row, 12)
    txtSRPMaxDis.Text = flxFind.TextMatrix(flxFind.row, 11)
    
    txtSRPProfRatio.Text = "0"
    txtSRPProfMar.Text = "0"
    txtSRPFixPrice.Text = "0"
'    txtSRPMaxDis.Text = "0"
    
    txtWSProfRatio.Text = "0"
    txtWSProfMar.Text = "0"
    txtWSFixPrice.Text = "0"
    
    fRow = 1
    fCol = 1
    
    DoEvents
    Call LoadSRPHist(SPid)
    DoEvents
    Call LoadWSHist(SPid)
    
    DoEvents
    Call getSRPProfMarginRatio
    Call getWSProfMarRatio
    
    Check1(0).Value = 0
    Check1(1).Value = 0
    Check1(2).Value = 0
    Check1(3).Value = 0
    Check1(4).Value = 0
    Check1(5).Value = 0

    fRow = i + 1
    fCol = 2
    
    cmdTrans(0).Enabled = True
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub LoadSRPHist(pSPid As Integer)
On Error GoTo er
    Dim tmpRs As New ADODB.Recordset
    Dim SPid As Integer
    
    
    Call modSRP.sp_srp_history(pSPid, tmpRs)
    If tmpRs.State = 0 Then Exit Sub
    
    If tmpRs.RecordCount <= 0 Then
        tmpRs.Close
        flxSRPPriceHis.Rows = 1
        flxSRPPriceHis.Rows = 2
        Exit Sub
    End If
    
'   0       1      2     3      4       5
'DATE, REFERENCE, COST, SRP, RATIO, MARGIN

    flxSRPPriceHis.Rows = 1
    flxSRPPriceHis.Rows = 2

    tmpRs.MoveFirst
           
        While Not tmpRs.EOF
            
            If flxSRPPriceHis.TextMatrix(1, 1) = "" Then
                flxSRPPriceHis.TextMatrix(1, 1) = tmpRs!SRPDATE
                flxSRPPriceHis.TextMatrix(1, 2) = tmpRs!REFERENCE
                flxSRPPriceHis.TextMatrix(1, 3) = Format(tmpRs!Cost, "#,##0.00")
                flxSRPPriceHis.TextMatrix(1, 4) = Format(tmpRs!SRP, "#,##0.00")
                flxSRPPriceHis.TextMatrix(1, 5) = tmpRs!PRORATIO
                flxSRPPriceHis.TextMatrix(1, 6) = Format(tmpRs!ProMar, "#,##0.00")
            
            Else
            
                flxSRPPriceHis.AddItem _
                vbTab & tmpRs!SRPDATE & _
                vbTab & tmpRs!REFERENCE & _
                vbTab & Format(tmpRs!Cost, "#,##0.00") & _
                vbTab & Format(tmpRs!SRP, "#,##0.00") & _
                vbTab & tmpRs!PRORATIO & _
                vbTab & Format(tmpRs!ProMar, "#,##0.00")
              
            End If
            
        tmpRs.MoveNext
        
    Wend
    
    If tmpRs.State <> 0 Then tmpRs.Close
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->LoadSRPHist()")
End Sub

Private Sub LoadWSHist(pSPid As Integer)
On Error GoTo er
    Dim tmpRs As New ADODB.Recordset
    
    Call modSRP.sp_ws_history(pSPid, tmpRs)
    If tmpRs.State = 0 Then Exit Sub
    
    If tmpRs.RecordCount <= 0 Then
        tmpRs.Close
        flxWSPriceHis.Rows = 1
        flxWSPriceHis.Rows = 2
        Exit Sub
    End If
    
'   0       1            2          3           4       5
'^DATE |^REF. NO. |^PREV. COST |^PREV. SRP |^RATIO |^MARGIN

    flxWSPriceHis.Rows = 1
    flxWSPriceHis.Rows = 2
      
    tmpRs.MoveFirst
           
        While Not tmpRs.EOF
            
            If flxWSPriceHis.TextMatrix(1, 1) = "" Then
                flxWSPriceHis.TextMatrix(1, 1) = tmpRs!WSDATE
                flxWSPriceHis.TextMatrix(1, 2) = tmpRs!RefNo
                flxWSPriceHis.TextMatrix(1, 3) = Format(tmpRs!Cost, "#,##0.00")
                flxWSPriceHis.TextMatrix(1, 4) = Format(tmpRs!DP, "#,##0.00")
                flxWSPriceHis.TextMatrix(1, 5) = tmpRs!PRORATIO
                flxWSPriceHis.TextMatrix(1, 6) = Format(tmpRs!ProMar, "#,##0.00")

            Else
            
                flxWSPriceHis.AddItem _
                vbTab & tmpRs!WSDATE & _
                vbTab & tmpRs!RefNo & _
                vbTab & Format(tmpRs!Cost, "#,##0.00") & _
                vbTab & Format(tmpRs!DP, "#,##0.00") & _
                vbTab & tmpRs!PRORATIO & _
                vbTab & Format(tmpRs!ProMar, "#,##0.00")
              
            End If
            
        tmpRs.MoveNext
        
    Wend
    
    If tmpRs.State <> 0 Then tmpRs.Close

                            
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->LoadWSHist()")
End Sub

Private Sub flxFind_DblClick()
    Screen.MousePointer = vbHourglass
    tmrSMSQTYLSI.Enabled = True
    DoEvents
    Call Load(frm_OnProcess2)
    frm_OnProcess2.Show vbModal
End Sub

Private Sub flxFind_EnterCell()
    Call flxFind_Click
End Sub

Private Sub Form_Load()
On Error GoTo er
Dim tmpRs As ADODB.Recordset

    DoEvents
    Call modSRP.sp_supplier_list(tmpRs)
    If tmpRs.State = 1 Then
        If tmpRs.RecordCount > 0 Then
            Call modObjTools.fillArray(AllArrSupplier, tmpRs)
        End If
    End If
    If tmpRs.State <> 0 Then tmpRs.Close
    
    DoEvents
    Call modObjTools.addToListObjectUsingArrayWithCols(cboSupplier, arrSupplier, AllArrSupplier, 1, 1)
    
    DoEvents
    Call modSRP.sp_class_list(tmpRs)
    If tmpRs.State = 1 Then
        If tmpRs.RecordCount > 0 Then
            Call modObjTools.fillArray(AllArrClass, tmpRs)
        End If
    End If
    If tmpRs.State <> 0 Then tmpRs.Close

    DoEvents
    Call modObjTools.addToListObjectUsingArrayWithCols(cboType, arrClass, AllArrClass, 0, 0)
    
    
    
    Warehouse.sBar.Panels(6).Text = "Price Adjustment"
    pSupCode = ""
    
    DoEvents
    Screen.MousePointer = vbHourglass
    tmrRefresh.Enabled = True
    DoEvents
    Call Load(frm_OnProcess2)
    frm_OnProcess2.Show vbModal

    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->Form_Load()")
End Sub



Private Sub RXResizer1_ResizeComplete()
On Error Resume Next
       
       'If flxFind.TextMatrix(1, 1) <> "" Then Exit Sub
       
        DoEvents
        With flxFind
            '.Clear
            '.Rows = 2
             TCol = 13
            .Cols = TCol
            .FixedCols = 1
            
            '               0       1             2            3           4      5       6               7         8              9                10
            .FormatString = "|<PART NUMBER |<DESCRIPTION  |^NET OF VAT  |>SRP  |>WSP  |>QTY ON HAND |>SIX MSQ |>BACK ORDER |>LAST DATE IN |<APPLICABLE MODEL"
            
            .ColWidth(0) = 200
            .ColWidth(1) = (.Width - .ColWidth(0)) * 0.1  'PART NUMBER
            .ColWidth(2) = (.Width - .ColWidth(0)) * 0.28  'DESCRIPTION
            
            .ColWidth(3) = (.Width - .ColWidth(0)) * 0.07  'NET OF VAT
            .ColWidth(4) = (.Width - .ColWidth(0)) * 0.07  'SRP
            .ColWidth(5) = (.Width - .ColWidth(0)) * 0.07  'WSP
            
            .ColWidth(6) = (.Width - .ColWidth(0)) * 0.07  'QTY ON HAND
            .ColWidth(7) = (.Width - .ColWidth(0)) * 0.07  'SIX MSQ
            .ColWidth(8) = (.Width - .ColWidth(0)) * 0.07  'BACKORDER
            .ColWidth(9) = (.Width - .ColWidth(0)) * 0.07  'LAST DATE IN
            .ColWidth(10) = (.Width - .ColWidth(0)) * 0.1 'APPLICABLE MODEL
            .ColWidth(11) = 0 'SPAREPARTID
            .ColWidth(12) = 0 'MAX DISCOUNT
        End With
        
   
        With flxSRPPriceHis
             .Cols = 7
             '               0   1       2         3       4        5         6
             .FormatString = "|^DATE |^REF. NO. |^COST |^PRICE |^P-RATIO |^P-MARGIN"
             
             .ColWidth(0) = 200
             .ColWidth(1) = (.Width - .ColWidth(0)) * 0.15 'DATE
             .ColWidth(2) = (.Width - .ColWidth(0)) * 0.16  'REF. NO.
             .ColWidth(3) = (.Width - .ColWidth(0)) * 0.16 'PREV. COST
             .ColWidth(4) = (.Width - .ColWidth(0)) * 0.16 'PREV. SRP
             .ColWidth(5) = (.Width - .ColWidth(0)) * 0.15 'RATIO
             .ColWidth(6) = (.Width - .ColWidth(0)) * 0.15 'MARGIN
        
        End With
        
        With flxWSPriceHis
             .Cols = 7
             '               0    1      2         3     4         5          6
             .FormatString = "|^DATE |^REF. NO. |^COST |^PRICE |^P-RATIO |^P-MARGIN"
             
             .ColWidth(0) = 200
             .ColWidth(1) = (.Width - .ColWidth(0)) * 0.15  'DATE
             .ColWidth(2) = (.Width - .ColWidth(0)) * 0.16  'REF. NO.
             .ColWidth(3) = (.Width - .ColWidth(0)) * 0.16 'PREV. COST
             .ColWidth(4) = (.Width - .ColWidth(0)) * 0.16 'PREV. SRP
             .ColWidth(5) = (.Width - .ColWidth(0)) * 0.15 'RATIO
             .ColWidth(6) = (.Width - .ColWidth(0)) * 0.15 'MARGIN
        
        End With
        
        'DoEvents
        'Call flxHistColsChanger
        
        'Call FindDesc
        
'        Screen.MousePointer = vbHourglass
'        tmrRefresh.Enabled = True
'        DoEvents
'        frm_OnProcess2.OwnerForm = "SRP Updater"
'        Call Load(frm_OnProcess2)
'        frm_OnProcess2.Show vbModal
    
    
End Sub


Private Sub tmr0SRP_Timer()
On Error GoTo er

    tmr0SRP.Enabled = False
    Screen.MousePointer = vbHourglass
    frmPriceFactor.Enabled = False
    
    DoEvents
    DoEvents
    Call FindZeroSRP
    DoEvents
    DoEvents
    
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    
    DoEvents
    frmPriceFactor.Enabled = True
    Screen.MousePointer = vbDefault
    
    'flxFind.HighLight = flexHighlightNever
    
    didRecordEntry = False
    
Exit Sub
er:
    frmPriceFactor.Enabled = True
    Screen.MousePointer = vbDefault
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    Call modObjTools.ErrorMessage(Err.Description, "Error->tmr0SRP_Timer()")
End Sub

Private Sub tmrSMSQTYLSI_Timer()
On Error GoTo er

    tmrSMSQTYLSI.Enabled = False
    Screen.MousePointer = vbHourglass
    frmPriceFactor.Enabled = False
    
    DoEvents
    Call GetSMSQTYLDI(SPid)
    
    DoEvents
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    
    DoEvents
    frmPriceFactor.Enabled = True
    Screen.MousePointer = vbDefault
    
    'flxFind.HighLight = flexHighlightNever
    
    didRecordEntry = False
    
Exit Sub
er:
    frmPriceFactor.Enabled = True
    Screen.MousePointer = vbDefault
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    Call modObjTools.ErrorMessage(Err.Description, "Error->tmrSMSQTYLSI_Timer()")
End Sub


Private Sub tmrRefresh_Timer()
On Error GoTo er

    tmrRefresh.Enabled = False
    Screen.MousePointer = vbHourglass
    frmPriceFactor.Enabled = False
    
    DoEvents
    DoEvents
    'Call LoadRecordBySup(SupCode)
    If lblFindParts.Caption = "FIND PART DESCRIPTION:" Then
        Call FindDesc
    Else
        Call FindPart
    End If
    DoEvents
    DoEvents
    
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    
    DoEvents
    frmPriceFactor.Enabled = True
    Screen.MousePointer = vbDefault
    
    flxFind.HighLight = flexHighlightAlways
    
    didRecordEntry = False
    
Exit Sub
er:
    frmPriceFactor.Enabled = True
    Screen.MousePointer = vbDefault
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    Call modObjTools.ErrorMessage(Err.Description, "Error->tmrRefresh_Timer()")
End Sub

Private Sub txtFindPart_Change()
On Error GoTo er
    Dim Qry As String
    
    If txtFindPart.Text = "" Then Exit Sub
    If lblFindParts.Caption = "FIND PART DESCRIPTION" Then Exit Sub
    
    If bNone Then
        bNone = False
        Exit Sub
    End If
    
    If rsParts.State = adStateOpen Then rsParts.Close
    
    Qry = "SELECT PARTNUMBER FROM mrcwarehouse.spareparts WHERE PARTNUMBER LIKE '" + txtFindPart.Text + "%';"
    'InputBox Qry, Qry, Qry
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic

    If rsParts.RecordCount > 0 Then
        If Not bNone Then
            txtTag.Text = Len(txtFindPart.Text)
            bNone = True
        End If
        
        txtFindPart.Text = rsParts!PARTNUMBER
        txtFindPart.SelStart = Val(txtTag.Text)
        txtFindPart.SelLength = Len(txtFindPart.Text)
    End If

Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtFindPart_Change()")
End Sub

Private Sub txtFindPart_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Or KeyCode = 46 Then
        bNone = True
    End If
End Sub


Private Sub txtFindPart_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub txtrefno_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Check1(0).SetFocus
End Sub


Private Sub txtSRPFixPrice_KeyPress(KeyAscii As Integer)
On Error GoTo er

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtSRPFixPrice, KeyAscii, NumbersOnly)
    
    If KeyAscii = 13 Then
            Dim ProMar As Double
            Dim CostVat As Double
            Dim NewSRP As Double
            Dim Ratio As Double
            
            NewSRP = Val(Replace(Trim(txtSRPFixPrice.Text), ",", ""))
            'CostVat = Format(txtNetofVat.Text, "#,#00")
            CostVat = Val(Replace(txtNetofVat, ",", ""))
            CostVat = CostVat * 1.12
            CostVat = Format(CostVat, "#,#00")
            
            If CostVat = 0 Then
                MsgBox "This part number has no cost!", vbExclamation, Me.Caption
                Exit Sub
            End If
            
'            If NewSRP < CostVat Then
'                MsgBox "New SRP is less than new cost!", vbExclamation, Me.Caption
'                Exit Sub
'            Else
                ProMar = NewSRP - CostVat
 '           End If
            
            'Ratio = ProMar / NewSRP
            Ratio = ProMar / CostVat
            Ratio = Ratio * 100
            
            txtSRPFixPrice.Text = Format(NewSRP, "#,#00")
            txtSRPProfMar.Text = Format(ProMar, "#,#00")
            
'            If Ratio > 100 Then
'                Ratio = 100
'            End If
            
            txtSRPProfRatio.Text = Format(Ratio, "#,#00.00")
    
    End If

Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtSRPFixPrice_KeyPress()")
End Sub




Private Sub txtSRPFixPrice_LostFocus()
    If txtSRPFixPrice.Text < 1 Or txtSRPFixPrice.Text = "" Then Exit Sub
    Call txtSRPFixPrice_KeyPress(13)
End Sub


Private Sub txtSRPMaxDis_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtSRPMaxDis, KeyAscii, NumbersOnly)
End Sub

Private Sub txtSRPProfMar_KeyPress(KeyAscii As Integer)
On Error GoTo er

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtSRPProfMar, KeyAscii, NumbersOnly)
    
    If KeyAscii = 13 Then
            Dim CostVat As Double
            Dim NewSRP As Double
            Dim Ratio As Double
            
            'CostVat = Format(txtNetofVat.Text, "#,#00")
            CostVat = Val(Replace(txtNetofVat.Text, ",", ""))
            CostVat = CostVat * 1.12
            CostVat = Format(CostVat, "#,#00")
            
            If CostVat = 0 Then
                MsgBox "This part number has no cost!", vbExclamation, Me.Caption
                Exit Sub
            End If
            
            NewSRP = CostVat + Val(Replace(txtSRPProfMar.Text, ",", ""))
            
            'Ratio = txtSRPProfMar.Text / NewSRP
            Ratio = txtSRPProfMar.Text / CostVat
            Ratio = Ratio * 100
            
            txtSRPFixPrice.Text = Format(NewSRP, "#,#00")
            txtSRPProfMar.Text = Format(txtSRPProfMar.Text, "#,#00")
            
'            If Ratio > 100 Then
'                Ratio = 100
'            End If
            
            txtSRPProfRatio.Text = Format(Ratio, "#,#00.00")
    End If
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtSRPProfMar_KeyPress()")
End Sub

Private Sub txtSRPProfMar_LostFocus()
    If txtSRPProfMar.Text < 1 Or txtSRPProfMar.Text = "" Then Exit Sub
    Call txtSRPProfMar_KeyPress(13)
End Sub


Private Sub txtSRPProfRatio_KeyPress(KeyAscii As Integer)
On Error GoTo er

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtSRPProfRatio, KeyAscii, NumbersWithDecimalPoint)
    
    If KeyAscii = 13 Then
            Dim CostVat As Double
            Dim NewSRP As Double
            Dim Ratio As Double
            Dim ProfMar As Double
            

            CostVat = Val(Replace(txtNetofVat, ",", ""))
            CostVat = CostVat * 1.12
            CostVat = Format(CostVat, "#,#00")

            If CostVat = 0 Then
                MsgBox "This part number has no cost!", vbExclamation, Me.Caption
                Exit Sub
            End If
            
            Ratio = txtSRPProfRatio.Text
            
  '          If Ratio > 100 Then
                NewSRP = CostVat + (CostVat * (Ratio / 100))
                ProfMar = NewSRP - CostVat
                txtSRPFixPrice.Text = Format(NewSRP, "#,#00")
                txtSRPProfMar.Text = Format(ProfMar, "#,#00")
                txtSRPProfRatio.Text = Format(txtSRPProfRatio.Text, "#,#00.00")
'                Exit Sub
'            End If

'            Ratio = (100 - Ratio) / 100
'
'            If Ratio <= 0 Then Ratio = 1
'
'            NewSRP = CostVat / Ratio
'            ProfMar = NewSRP - CostVat
'
'            txtSRPFixPrice.Text = Format(NewSRP, "#,#00")
'            txtSRPProfMar.Text = Format(ProfMar, "#,#00")
'            txtSRPProfRatio.Text = Format(txtSRPProfRatio.Text, "#,#00.00")

    End If
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtSRPProfRatio_KeyPress()")
End Sub


Private Sub txtSRPProfRatio_LostFocus()
    If txtSRPProfRatio.Text < 1 Or txtSRPProfRatio.Text = "" Then Exit Sub
    Call txtSRPProfRatio_KeyPress(13)
End Sub


Private Sub txtWSFixPrice_KeyPress(KeyAscii As Integer)
On Error GoTo er

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtWSFixPrice, KeyAscii, NumbersOnly)
    
    If KeyAscii = 13 Then
            Dim ProMar As Double
            Dim CostVat As Double
            Dim NewWSP As Double
            Dim Ratio As Double
            
            
            NewWSP = Val(Replace(Trim(txtWSFixPrice.Text), ",", ""))
            'CostVat = Format(txtNetofVat.Text, "#,#00")
            CostVat = Val(Replace(txtNetofVat, ",", ""))
            CostVat = CostVat * 1.12
            CostVat = Format(CostVat, "#,#00")
            
            If CostVat = 0 Then
                MsgBox "This part number has no cost!", vbExclamation, Me.Caption
                Exit Sub
            End If
            
'            If NewWSP < CostVat Then
'                MsgBox "New WSP is less than new cost!", vbExclamation, Me.Caption
'                Exit Sub
'            Else
                ProMar = NewWSP - CostVat
 '           End If
            
            'Ratio = ProMar / NewWSP
            Ratio = ProMar / CostVat
            Ratio = Ratio * 100
            
            txtWSFixPrice.Text = Format(NewWSP, "#,#00")
            txtWSProfMar.Text = Format(ProMar, "#,#00")
            
'            If Ratio > 100 Then
'                Ratio = 100
'            End If
            
            txtWSProfRatio.Text = Format(Ratio, "#,#00.00")
    End If
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtWSFixPrice_KeyPress()")
End Sub

Private Sub txtWSFixPrice_LostFocus()
    If txtWSFixPrice.Text < 1 Or txtWSFixPrice.Text = "" Then Exit Sub
    Call txtWSFixPrice_KeyPress(13)
End Sub


Private Sub txtWSProfMar_KeyPress(KeyAscii As Integer)
On Error GoTo er

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtWSProfMar, KeyAscii, NumbersOnly)
    
    If KeyAscii = 13 Then
            Dim CostVat As Double
            Dim NewWSP As Double
            Dim Ratio As Double
            
            'CostVat = Format(txtNetofVat.Text, "#,#00")
            CostVat = Val(Replace(txtNetofVat.Text, ",", ""))
            CostVat = CostVat * 1.12
            CostVat = Format(CostVat, "#,#00")
    
            If CostVat = 0 Then
                MsgBox "This part number has no cost!", vbExclamation, Me.Caption
                Exit Sub
            End If
            
            NewWSP = CostVat + Val(Replace(txtWSProfMar.Text, ",", ""))
            
            'Ratio = txtWSProfMar.Text / NewWSP
            Ratio = Val(Replace(txtWSProfMar.Text, ",", "")) / CostVat
            Ratio = Ratio * 100
            
            txtWSFixPrice.Text = Format(NewWSP, "#,#00")
            txtWSProfMar.Text = Format(txtWSProfMar, "#,#00")
            
'            If Ratio > 100 Then
'                Ratio = 100
'            End If
            
            txtWSProfRatio.Text = Format(Ratio, "#,#00.00")
    End If
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtWSProfMar_KeyPress()")
End Sub

Private Sub txtWSProfMar_LostFocus()
    If txtWSProfMar.Text < 1 Or txtWSProfMar.Text = "" Then Exit Sub
    Call txtWSProfMar_KeyPress(13)
End Sub

Private Sub txtWSProfRatio_KeyPress(KeyAscii As Integer)
On Error GoTo er

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtWSProfRatio, KeyAscii, NumbersWithDecimalPoint)
        
    If KeyAscii = 13 Then
        Dim CostVat As Double
        Dim NewWSP As Double
        Dim Ratio As Double
        Dim ProfMar As Double
        
        CostVat = Val(Replace(txtNetofVat, ",", ""))
        CostVat = CostVat * 1.12
        CostVat = Format(CostVat, "#,#00")

        If CostVat = 0 Then
            MsgBox "This part number has no cost!", vbExclamation, Me.Caption
            Exit Sub
        End If
        
        Ratio = txtWSProfRatio.Text
        
'        If Ratio > 100 Then
            NewWSP = CostVat + (CostVat * (Ratio / 100))
            ProfMar = NewWSP - CostVat
            txtWSFixPrice.Text = Format(NewWSP, "#,#00")
            txtWSProfMar.Text = Format(ProfMar, "#,#00")
            txtWSProfRatio.Text = Format(txtWSProfRatio.Text, "#,#00.00")
'            Exit Sub
'        End If
'
'        Ratio = (100 - Ratio) / 100
'
'        If Ratio <= 0 Then Ratio = 1
'
'        NewWSP = CostVat / Ratio
'        ProfMar = NewWSP - CostVat
'
'        txtWSFixPrice.Text = Format(NewWSP, "#,#00")
'        txtWSProfMar.Text = Format(ProfMar, "#,#00")
'        txtWSProfRatio.Text = Format(txtWSProfRatio.Text, "#,#00.00")
    End If
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtWSProfRatio_KeyPress()")
End Sub

Private Sub txtWSProfRatio_LostFocus()
    If txtWSProfRatio.Text < 1 Or txtWSProfRatio.Text = "" Then Exit Sub
    Call txtWSProfRatio_KeyPress(13)
End Sub


Private Sub VScroll1_Change()
    Dim tmpRs As ADODB.Recordset
    VScroll1.Max = 2
    
    If VScroll1.Value = 0 Then
        lblFindParts.Caption = "FIND PART DESCRIPTION:"
        
    ElseIf VScroll1.Value = 1 Then
        lblFindParts.Caption = "FIND PART NUMBER:"
        
    ElseIf VScroll1.Value = 2 Then
        VScroll1.Value = 0
    End If
    
End Sub


Sub flxFindColsChanger()
On Error Resume Next
    With flxFind
        .Clear
        .Rows = 2
         TCol = 12
        .Cols = TCol
        .FixedCols = 1
        
        '               0       1             2            3           4      5       6      7        8              9
        .FormatString = "|<PART NUMBER |<DESCRIPTION  |^NET OF VAT  |>SRP  |>WSP  |>QTY. |>SIX MSQ |>BACK ORDER |<APPLICABLE MODEL"
        
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.15  'PART NUMBER
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.27  'DESCRIPTION
        
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.08  'NET OF VAT
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.08  'SRP
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.08  'WSP
        
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.07  'QTY.
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.07  'SIX MSQ
        .ColWidth(8) = (.Width - .ColWidth(0)) * 0.07  'BACKORDER
        .ColWidth(9) = (.Width - .ColWidth(0)) * 0.1   'APPLICABLE MODEL
        .ColWidth(10) = 0 'SPAREPARTID
        .ColWidth(11) = 0 'MAX DISCOUNT
    End With
    
End Sub


Sub flxHistColsChanger()
On Error Resume Next
    With flxSRPPriceHis
         .Clear
         .Rows = 2
         .Cols = 7
         '               0   1       2         3       4        5         6
         .FormatString = "|^DATE |^REF. NO. |^COST |^PRICE |^P-RATIO |^P-MARGIN"
         
         .ColWidth(0) = 200
         .ColWidth(1) = (.Width - .ColWidth(0)) * 0.15 'DATE
         .ColWidth(2) = (.Width - .ColWidth(0)) * 0.16  'REF. NO.
         .ColWidth(3) = (.Width - .ColWidth(0)) * 0.16 'PREV. COST
         .ColWidth(4) = (.Width - .ColWidth(0)) * 0.16 'PREV. SRP
         .ColWidth(5) = (.Width - .ColWidth(0)) * 0.15 'RATIO
         .ColWidth(6) = (.Width - .ColWidth(0)) * 0.15 'MARGIN
    
    End With
    
    
    With flxWSPriceHis
         .Clear
         .Rows = 2
         .Cols = 7
         '               0    1      2         3     4         5          6
         .FormatString = "|^DATE |^REF. NO. |^COST |^PRICE |^P-RATIO |^P-MARGIN"
         
         .ColWidth(0) = 200
         .ColWidth(1) = (.Width - .ColWidth(0)) * 0.15  'DATE
         .ColWidth(2) = (.Width - .ColWidth(0)) * 0.16  'REF. NO.
         .ColWidth(3) = (.Width - .ColWidth(0)) * 0.16 'PREV. COST
         .ColWidth(4) = (.Width - .ColWidth(0)) * 0.16 'PREV. SRP
         .ColWidth(5) = (.Width - .ColWidth(0)) * 0.15 'RATIO
         .ColWidth(6) = (.Width - .ColWidth(0)) * 0.15 'MARGIN
    
    End With

End Sub

'Public Sub loadRecord()
'On Error GoTo er
'    Dim tmpRs As New ADODB.Recordset
'
'    Call modSRP.sp_load_record(tmpRs)
'    If tmpRs.State = 0 Then Exit Sub
'
'    If tmpRs.RecordCount <= 0 Then
'        tmpRs.Close
'        Exit Sub
'    End If
'
'    Call flxFindColsChanger
'
''       1             2            3          4      5       6      7        8              9
''|^PART NUMBER |^DESCRIPTION  |^NET OF VAT  |^SRP  |^WSP  |^QTY. |^SIX MSQ |^BACKORDER |^APPLICABLE MODEL "
''SPAREPARTID, PARTNUMBER, DESCRIPTION, COST, STOREPRICE, DLRPRICE, QTYONHAND, BACKORDER, APPLICMODEL
'    tmpRs.MoveFirst
'
'        While Not tmpRs.EOF
'
'            If flxFind.TextMatrix(1, 1) = "" Then
'                flxFind.TextMatrix(1, 1) = tmpRs!PartNumber
'                flxFind.TextMatrix(1, 2) = tmpRs!Description
'                flxFind.TextMatrix(1, 3) = Format(tmpRs!Cost, "#,##0.00")
'                flxFind.TextMatrix(1, 4) = Format(tmpRs!StorePrice, "#,##0")
'                flxFind.TextMatrix(1, 5) = Format(tmpRs!DlrPrice, "#,##0")
'                flxFind.TextMatrix(1, 6) = tmpRs!QtyOnHand
'                flxFind.TextMatrix(1, 7) = tmpRs!SixMSQ
'                flxFind.TextMatrix(1, 8) = tmpRs!BACKORDER
'                flxFind.TextMatrix(1, 9) = tmpRs!ApplicModel
'                flxFind.TextMatrix(1, 10) = tmpRs!SPAREPARTID
'                flxFind.TextMatrix(1, 11) = tmpRs!MaxDiscount
'
'            Else
'
'                flxFind.AddItem _
'                vbTab & tmpRs!PartNumber & _
'                vbTab & tmpRs!Description & _
'                vbTab & Format(tmpRs!Cost, "#,##0.00") & _
'                vbTab & Format(tmpRs!StorePrice, "#,##0") & _
'                vbTab & Format(tmpRs!DlrPrice, "#,##0") & _
'                vbTab & tmpRs!QtyOnHand & _
'                vbTab & tmpRs!SixMSQ & _
'                vbTab & tmpRs!BACKORDER & _
'                vbTab & tmpRs!ApplicModel & _
'                vbTab & tmpRs!SPAREPARTID & _
'                vbTab & tmpRs!MaxDiscount
'
'            End If
'
'        tmpRs.MoveNext
'
'    Wend
'
'    If tmpRs.State <> 0 Then tmpRs.Close
'
'    flxFind.HighLight = flexHighlightNever
'
'Exit Sub
'er:
'    DoEvents
'    Screen.MousePointer = vbDefault
'    Call modObjTools.ErrorMessage(Err.Description, "Error->LoadRecord()")
'End Sub

'Public Sub LoadRecordBySup(pSupCode As String)
'On Error GoTo er
'    Dim tmpRs As New ADODB.Recordset
'
'    Call modSRP.sp_load_record(pSupCode, tmpRs)
'    If tmpRs.State = 0 Then Exit Sub
'
'    If tmpRs.RecordCount <= 0 Then
'        tmpRs.Close
'        Exit Sub
'    End If
'
'    Call flxFindColsChanger
'
'    frm_OnProcess2.ProgLoading.Value = 0
'
''       1             2            3          4      5       6      7        8              9
''|^PART NUMBER |^DESCRIPTION  |^NET OF VAT  |^SRP  |^WSP  |^QTY. |^SIX MSQ |^BACKORDER |^APPLICABLE MODEL "
''SPAREPARTID, PARTNUMBER, DESCRIPTION, COST, STOREPRICE, DLRPRICE, QTYONHAND, BACKORDER, APPLICMODEL
'    tmpRs.MoveFirst
'
'        While Not tmpRs.EOF
'
'            frm_OnProcess2.ProgLoading.Max = tmpRs.RecordCount
'
'            If frm_OnProcess2.ProgLoading.Value < frm_OnProcess2.ProgLoading.Max Then
'                frm_OnProcess2.ProgLoading.Value = frm_OnProcess2.ProgLoading.Value + 1
'            End If
'
'            If flxFind.TextMatrix(1, 1) = "" Then
'                flxFind.TextMatrix(1, 1) = tmpRs!PartNumber
'                flxFind.TextMatrix(1, 2) = tmpRs!Description
'                flxFind.TextMatrix(1, 3) = Format(tmpRs!Cost, "#,##0.00")
'                flxFind.TextMatrix(1, 4) = Format(tmpRs!StorePrice, "#,##0")
'                flxFind.TextMatrix(1, 5) = Format(tmpRs!DlrPrice, "#,##0")
'                flxFind.TextMatrix(1, 6) = tmpRs!QtyOnHand
'                flxFind.TextMatrix(1, 7) = tmpRs!SixMSQ
'                flxFind.TextMatrix(1, 8) = tmpRs!BACKORDER
'                flxFind.TextMatrix(1, 9) = tmpRs!ApplicModel
'                flxFind.TextMatrix(1, 10) = tmpRs!SPAREPARTID
'                flxFind.TextMatrix(1, 11) = tmpRs!MAXDISCOUNT
'
'            Else
'
'                flxFind.AddItem _
'                vbTab & tmpRs!PartNumber & _
'                vbTab & tmpRs!Description & _
'                vbTab & Format(tmpRs!Cost, "#,##0.00") & _
'                vbTab & Format(tmpRs!StorePrice, "#,##0") & _
'                vbTab & Format(tmpRs!DlrPrice, "#,##0") & _
'                vbTab & tmpRs!QtyOnHand & _
'                vbTab & tmpRs!SixMSQ & _
'                vbTab & tmpRs!BACKORDER & _
'                vbTab & tmpRs!ApplicModel & _
'                vbTab & tmpRs!SPAREPARTID & _
'                vbTab & tmpRs!MAXDISCOUNT
'
'            End If
'
'        tmpRs.MoveNext
'
'    Wend
'
'    If tmpRs.State <> 0 Then tmpRs.Close
'
'
'Exit Sub
'er:
'    DoEvents
'    Screen.MousePointer = vbDefault
'    Call modObjTools.ErrorMessage(Err.Description, "Error->LoadRecordBySup()")
'End Sub

Public Sub FindDesc()
On Error GoTo er
    Dim MySP() As String
    Dim Qry As String
    Dim myW As String, MySel As String
    

    MySP = Split(txtFindPart.Text, " ")
    myW = ""
    For i = LBound(MySP) To UBound(MySP)
        If myW = "" Then
            myW = " sp.DESCRIPTION Like '%" & MySP(i) & "%'"
        Else
            myW = myW & " AND sp.DESCRIPTION Like '%" & MySP(i) & "%'"
        End If
    Next i
    
    If lblFindParts.Caption = "FIND PART NUMBER:" Then
        txtFindPart.Text = ""
        VScroll1.Value = 0
    End If
    
    If rsParts.State = adStateOpen Then rsParts.Close
    If txtFindPart.Text = "" And chkAll.Value = 0 And SupCode = "" And cboType.Text = "" Then Exit Sub
    
    flxFind.Clear
    
'    If chkSup = True Then
'        SupCode = pSupCode
'    End If
    If txtFindPart.Text <> "" And chkAll.Value = 0 And cboSupplier.Text <> "" And cboType.Text = "" Then

'        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION " & _
'          ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND',ifnull(sptr.SixMonSale,0) as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'          " From SPAREPARTS as sp LEFT JOIN (select sum(`out`) as SixMonSale,sparepartid from spitr where  transdate >= DATE_SUB(now(), INTERVAL 6 MONTH) and transdate <= Now() and transtype='OUT'" & _
'          " and remarks<>'TRANSFER' GROUP by sparepartid) as sptr ON sp.sparepartid = sptr.sparepartid WHERE sp.SUPPLIERCODE= '" & SupCode & "' AND " & myW & " ORDER BY sp.PARTNUMBER;"
'
'        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
'            ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',ifnull(ldi.TRANSDATE,'') as LASTDATEIN" & _
'            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'            " From SPAREPARTS as sp LEFT JOIN (Select MAX(TRANSDATE) as TRANSDATE,sparepartid From spitr Where TRANSTYPE = 'RECEIVE' Group By SPAREPARTID)as ldi On sp.sparepartid = ldi.sparepartid" & _
'            " WHERE sp.SUPPLIERCODE= '" & SupCode & "' AND " & myW & " ORDER BY sp.PARTNUMBER;"

        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
                ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID " & _
            "From SPAREPARTS as sp WHERE sp.SUPPLIERCODE= '" & SupCode & "' AND " & myW & " ORDER BY sp.PARTNUMBER;"
  
        
    ElseIf chkAll.Value = 0 And txtFindPart.Text = "" And cboSupplier.Text <> "" And cboType.Text = "" Then
           
'        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION " & _
'          ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND',ifnull(sptr.SixMonSale,0) as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'          " From SPAREPARTS as sp LEFT JOIN (select sum(`out`) as SixMonSale,sparepartid from spitr where  transdate >= DATE_SUB(now(), INTERVAL 6 MONTH) and transdate <= Now() and transtype='OUT'" & _
'          " and remarks<>'TRANSFER' GROUP by sparepartid) as sptr ON sp.sparepartid = sptr.sparepartid WHERE sp.SUPPLIERCODE= '" & SupCode & "' ORDER BY sp.PARTNUMBER;"
        
'        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
'            ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',ifnull(ldi.TRANSDATE,'') as LASTDATEIN" & _
'            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'            " From SPAREPARTS as sp LEFT JOIN (Select MAX(TRANSDATE) as TRANSDATE,sparepartid From spitr Where TRANSTYPE = 'RECEIVE' Group By SPAREPARTID)as ldi On sp.sparepartid = ldi.sparepartid" & _
'            " WHERE sp.SUPPLIERCODE= '" & SupCode & "' ORDER BY sp.PARTNUMBER;"
            
        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
                ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID " & _
            "From SPAREPARTS as sp WHERE sp.SUPPLIERCODE= '" & SupCode & "' ORDER BY sp.PARTNUMBER;"
            
    ElseIf chkAll.Value = 1 And txtFindPart.Text <> "" Then
                
'        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION " & _
'          ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND',ifnull(sptr.SixMonSale,0) as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'          " From SPAREPARTS as sp LEFT JOIN (select sum(`out`) as SixMonSale,sparepartid from spitr where  transdate >= DATE_SUB(now(), INTERVAL 6 MONTH) and transdate <= Now() and transtype='OUT'" & _
'          " and remarks<>'TRANSFER' GROUP by sparepartid) as sptr ON sp.sparepartid = sptr.sparepartid WHERE " & myW & " ORDER BY sp.PARTNUMBER;"

'        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
'            ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',ifnull(ldi.TRANSDATE,'') as LASTDATEIN" & _
'            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'            " From SPAREPARTS as sp LEFT JOIN (Select MAX(TRANSDATE) as TRANSDATE,sparepartid From spitr Where TRANSTYPE = 'RECEIVE' Group By SPAREPARTID)as ldi On sp.sparepartid = ldi.sparepartid" & _
'            " WHERE " & myW & " ORDER BY sp.PARTNUMBER;"
            
        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
                ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID " & _
            "From SPAREPARTS as sp WHERE " & myW & " ORDER BY sp.PARTNUMBER;"
     
    ElseIf chkAll.Value = 1 And txtFindPart.Text = "" And cboSupplier.Text = "" And cboType.Text = "" Then
               
'        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION " & _
'          ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND',ifnull(sptr.SixMonSale,0) as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'          " From SPAREPARTS as sp LEFT JOIN (select sum(`out`) as SixMonSale,sparepartid from spitr where  transdate >= DATE_SUB(now(), INTERVAL 6 MONTH) and transdate <= Now() and transtype='OUT'" & _
'          " and remarks<>'TRANSFER' GROUP by sparepartid) as sptr ON sp.sparepartid = sptr.sparepartid ORDER BY sp.PARTNUMBER limit 50;"
    
'        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
'            ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',ifnull(ldi.TRANSDATE,'') as LASTDATEIN" & _
'            " ,sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'            " From SPAREPARTS as sp LEFT JOIN (Select MAX(TRANSDATE) as TRANSDATE,sparepartid From spitr Where TRANSTYPE = 'RECEIVE' Group By SPAREPARTID)as ldi On sp.sparepartid = ldi.sparepartid" & _
'            " ORDER BY sp.PARTNUMBER limit 50;"
            
        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
                ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID " & _
            "From SPAREPARTS as sp ORDER BY sp.PARTNUMBER limit 50;"
            
    'InputBox Qry, Qry, Qry
    ElseIf txtFindPart.Text <> "" And chkAll.Value = 0 And SupCode = "" Then
            
'        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION " & _
'          ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND',ifnull(sptr.SixMonSale,0) as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'          " From SPAREPARTS as sp LEFT JOIN (select sum(`out`) as SixMonSale,sparepartid from spitr where  transdate >= DATE_SUB(now(), INTERVAL 6 MONTH) and transdate <= Now() and transtype='OUT'" & _
'          " and remarks<>'TRANSFER' GROUP by sparepartid) as sptr ON sp.sparepartid = sptr.sparepartid  WHERE " & myW & " ORDER BY sp.PARTNUMBER;"

'        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
'            ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',ifnull(ldi.TRANSDATE,'') as LASTDATEIN" & _
'            " ,sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'            " From SPAREPARTS as sp LEFT JOIN (Select MAX(TRANSDATE) as TRANSDATE,sparepartid From spitr Where TRANSTYPE = 'RECEIVE' Group By SPAREPARTID)as ldi On sp.sparepartid = ldi.sparepartid" & _
'            " WHERE " & myW & " ORDER BY sp.PARTNUMBER;"
            
        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
                ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID " & _
            "From SPAREPARTS as sp WHERE " & myW & " ORDER BY sp.PARTNUMBER;"
            
    ElseIf txtFindPart.Text <> "" And chkAll.Value = 0 And cboSupplier.Text = "" Then
            
'        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION " & _
'          ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND',ifnull(sptr.SixMonSale,0) as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'          " From SPAREPARTS as sp LEFT JOIN (select sum(`out`) as SixMonSale,sparepartid from spitr where  transdate >= DATE_SUB(now(), INTERVAL 6 MONTH) and transdate <= Now() and transtype='OUT'" & _
'          " and remarks<>'TRANSFER' GROUP by sparepartid) as sptr ON sp.sparepartid = sptr.sparepartid  WHERE " & myW & " ORDER BY sp.PARTNUMBER;"
          
'        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
'            ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',ifnull(ldi.TRANSDATE,'') as LASTDATEIN" & _
'            " ,sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'            " From SPAREPARTS as sp LEFT JOIN (Select MAX(TRANSDATE) as TRANSDATE,sparepartid From spitr Where TRANSTYPE = 'RECEIVE' Group By SPAREPARTID)as ldi On sp.sparepartid = ldi.sparepartid" & _
'            " WHERE " & myW & " ORDER BY sp.PARTNUMBER;"
            
        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
                ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID " & _
            "From SPAREPARTS as sp WHERE " & myW & " ORDER BY sp.PARTNUMBER;"
            
    ElseIf txtFindPart.Text = "" And chkAll.Value = 0 And cboSupplier.Text = "" And cboType.Text = "" Then
        
'        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION " & _
'          ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND',ifnull(sptr.SixMonSale,0) as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'          " From SPAREPARTS as sp LEFT JOIN (select sum(`out`) as SixMonSale,sparepartid from spitr where  transdate >= DATE_SUB(now(), INTERVAL 6 MONTH) and transdate <= Now() and transtype='OUT'" & _
'          " and remarks<>'TRANSFER' GROUP by sparepartid) as sptr ON sp.sparepartid = sptr.sparepartid ORDER BY sp.PARTNUMBER limit 30;"
        
'        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
'            ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',ifnull(ldi.TRANSDATE,'') as LASTDATEIN" & _
'            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'            " From SPAREPARTS as sp LEFT JOIN (Select MAX(TRANSDATE) as TRANSDATE,sparepartid From spitr Where TRANSTYPE = 'RECEIVE' Group By SPAREPARTID)as ldi On sp.sparepartid = ldi.sparepartid" & _
'            " ORDER BY sp.PARTNUMBER limit 30;"
            
        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
                ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID " & _
            "From SPAREPARTS as sp ORDER BY sp.PARTNUMBER limit 30;"
            
    ElseIf txtFindPart.Text = "" And chkAll.Value = 1 And cboSupplier.Text = "" And cboType.Text <> "" Then
       
'        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION " & _
'          ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND',ifnull(sptr.SixMonSale,0) as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'          " From SPAREPARTS as sp LEFT JOIN (select sum(`out`) as SixMonSale,sparepartid from spitr where  transdate >= DATE_SUB(now(), INTERVAL 6 MONTH) and transdate <= Now() and transtype='OUT'" & _
'          " and remarks<>'TRANSFER' GROUP by sparepartid) as sptr ON sp.sparepartid = sptr.sparepartid  WHERE remarks= '" & Trim(cboType.Text) & "' ORDER BY sp.PARTNUMBER;"

'        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
'            ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',ifnull(ldi.TRANSDATE,'') as LASTDATEIN" & _
'            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'            " From SPAREPARTS as sp LEFT JOIN (Select MAX(TRANSDATE) as TRANSDATE,sparepartid From spitr Where TRANSTYPE = 'RECEIVE' Group By SPAREPARTID)as ldi On sp.sparepartid = ldi.sparepartid" & _
'            " WHERE remarks= '" & Trim(cboType.Text) & "' ORDER BY sp.PARTNUMBER;"
            
        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
                ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID " & _
            "From SPAREPARTS as sp WHERE remarks= '" & Trim(cboType.Text) & "' ORDER BY sp.PARTNUMBER;"
        
    ElseIf txtFindPart.Text <> "" And chkAll.Value = 0 And cboSupplier.Text <> "" And cboType.Text <> "" Then
  
'        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION " & _
'          ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND',ifnull(sptr.SixMonSale,0) as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'          " From SPAREPARTS as sp LEFT JOIN (select sum(`out`) as SixMonSale,sparepartid from spitr where  transdate >= DATE_SUB(now(), INTERVAL 6 MONTH) and transdate <= Now() and transtype='OUT'" & _
'          " and remarks<>'TRANSFER' GROUP by sparepartid) as sptr ON sp.sparepartid = sptr.sparepartid  WHERE remarks= '" & Trim(cboType.Text) & "' ORDER BY sp.PARTNUMBER;"

'        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
'            ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',ifnull(ldi.TRANSDATE,'') as LASTDATEIN" & _
'            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'            " From SPAREPARTS as sp LEFT JOIN (Select MAX(TRANSDATE) as TRANSDATE,sparepartid From spitr Where TRANSTYPE = 'RECEIVE' Group By SPAREPARTID)as ldi On sp.sparepartid = ldi.sparepartid" & _
'            " WHERE remarks= '" & Trim(cboType.Text) & "' ORDER BY sp.PARTNUMBER;"
            
        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
                ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID " & _
            "From SPAREPARTS as sp WHERE remarks= '" & Trim(cboType.Text) & "' ORDER BY sp.PARTNUMBER;"

    ElseIf chkAll.Value = 0 And cboSupplier.Text <> "" And cboType.Text <> "" Then
        
'        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION " & _
'          ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND',ifnull(sptr.SixMonSale,0) as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'          " From SPAREPARTS as sp LEFT JOIN (select sum(`out`) as SixMonSale,sparepartid from spitr where  transdate >= DATE_SUB(now(), INTERVAL 6 MONTH) and transdate <= Now() and transtype='OUT'" & _
'          " and remarks<>'TRANSFER' GROUP by sparepartid) as sptr ON sp.sparepartid = sptr.sparepartid  WHERE sp.SUPPLIERCODE= '" & SupCode & "' AND remarks= '" & Trim(cboType.Text) & "' ORDER BY sp.PARTNUMBER;"
    
'        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
'            ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',ifnull(ldi.TRANSDATE,'') as LASTDATEIN" & _
'            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'            " From SPAREPARTS as sp LEFT JOIN (Select MAX(TRANSDATE) as TRANSDATE,sparepartid From spitr Where TRANSTYPE = 'RECEIVE' Group By SPAREPARTID)as ldi On sp.sparepartid = ldi.sparepartid" & _
'            " WHERE sp.SUPPLIERCODE= '" & SupCode & "' AND remarks= '" & Trim(cboType.Text) & "' ORDER BY sp.PARTNUMBER;"
            
        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
                ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID " & _
            "From SPAREPARTS as sp WHERE sp.SUPPLIERCODE= '" & SupCode & "' AND remarks= '" & Trim(cboType.Text) & "' ORDER BY sp.PARTNUMBER;"
            
    ElseIf txtFindPart.Text <> "" And chkAll.Value = 1 And cboSupplier.Text = "" And cboType.Text <> "" Then
        
'        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION " & _
'          ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND',ifnull(sptr.SixMonSale,0) as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'          " From SPAREPARTS as sp LEFT JOIN (select sum(`out`) as SixMonSale,sparepartid from spitr where  transdate >= DATE_SUB(now(), INTERVAL 6 MONTH) and transdate <= Now() and transtype='OUT'" & _
'          " and remarks<>'TRANSFER' GROUP by sparepartid) as sptr ON sp.sparepartid = sptr.sparepartid WHERE " & myW & " AND remarks= '" & Trim(cboType.Text) & "' ORDER BY sp.PARTNUMBER;"

'        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
'            ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',ifnull(ldi.TRANSDATE,'') as LASTDATEIN" & _
'            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'            " From SPAREPARTS as sp LEFT JOIN (Select MAX(TRANSDATE) as TRANSDATE,sparepartid From spitr Where TRANSTYPE = 'RECEIVE' Group By SPAREPARTID)as ldi On sp.sparepartid = ldi.sparepartid" & _
'            " WHERE " & myW & " AND remarks= '" & Trim(cboType.Text) & "' ORDER BY sp.PARTNUMBER;"
            
        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
                ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID " & _
            "From SPAREPARTS as sp WHERE " & myW & " AND remarks= '" & Trim(cboType.Text) & "' ORDER BY sp.PARTNUMBER;"
        
    ElseIf txtFindPart.Text = "" And chkAll.Value = 0 And cboSupplier.Text = "" And cboType.Text <> "" Then
        
'        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION " & _
'          ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND',ifnull(sptr.SixMonSale,0) as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'          " From SPAREPARTS as sp LEFT JOIN (select sum(`out`) as SixMonSale,sparepartid from spitr where  transdate >= DATE_SUB(now(), INTERVAL 6 MONTH) and transdate <= Now() and transtype='OUT'" & _
'          " and remarks<>'TRANSFER' GROUP by sparepartid) as sptr ON sp.sparepartid = sptr.sparepartid WHERE remarks= '" & Trim(cboType.Text) & "' ORDER BY sp.PARTNUMBER;"

'        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
'            ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',ifnull(ldi.TRANSDATE,'') as LASTDATEIN" & _
'            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'            " From SPAREPARTS as sp LEFT JOIN (Select MAX(TRANSDATE) as TRANSDATE,sparepartid From spitr Where TRANSTYPE = 'RECEIVE' Group By SPAREPARTID)as ldi On sp.sparepartid = ldi.sparepartid" & _
'            " WHERE remarks= '" & Trim(cboType.Text) & "' ORDER BY sp.PARTNUMBER;"
            
        Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
                ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
            ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID " & _
            "From SPAREPARTS as sp WHERE remarks= '" & Trim(cboType.Text) & "' ORDER BY sp.PARTNUMBER;"
            
    End If
   
    If Qry = "" Then Exit Sub
    
   'InputBox Qry, Qry, Qry
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsParts.RecordCount = 0 Then
        flxFind.Clear
        flxFind.Rows = 2
        Call RXResizer1_ResizeComplete
    Else
        Set flxFind.DataSource = rsParts
        flxFind.Refresh
        flxFind.ColAlignment(3) = flexAlignRightCenter
        flxFind.ColAlignment(4) = flexAlignRightCenter
        flxFind.ColAlignment(5) = flexAlignRightCenter
        Call RXResizer1_ResizeComplete
        Call flxFind_Click
    End If
    
    Call FormatCur
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->FindDesc()")
End Sub

Public Sub FindPart()
On Error GoTo er
    Dim Qry As String
    Dim myW As String, MySel As String
    
    
    If rsParts.State = adStateOpen Then rsParts.Close
    If txtFindPart.Text = "" Then Exit Sub
    
    flxFind.Clear
    
        
'    Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION " & _
'      ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND',ifnull(sptr.SixMonSale,0) as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID" & _
'      " From SPAREPARTS as sp LEFT JOIN (select sum(`out`) as SixMonSale,sparepartid from spitr where  transdate >= DATE_SUB(now(), INTERVAL 6 MONTH) and transdate <= Now() and transtype='OUT'" & _
'      " and remarks<>'TRANSFER' GROUP by sparepartid) as sptr ON sp.sparepartid = sptr.sparepartid WHERE sp.PARTNUMBER LIKE '%" & Trim(txtFindPart.Text) & "%' ORDER BY sp.PARTNUMBER;"

'    Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
'            ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER',ifnull(ldi.TRANSDATE,'') as LASTDATEIN" & _
'        ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID " & _
'        "From SPAREPARTS as sp LEFT JOIN (Select MAX(TRANSDATE) as TRANSDATE,sparepartid From spitr Where TRANSTYPE = 'RECEIVE' Group By SPAREPARTID)as ldi On sp.sparepartid = ldi.sparepartid " & _
'        "WHERE sp.PARTNUMBER LIKE '%" & Trim(txtFindPart.Text) & "%' ORDER BY sp.PARTNUMBER;"

    Qry = "SELECT sp.PARTNUMBER,CONCAT(Ifnull(TRIM(sp.DESCRIPTION), ''), ' ', Ifnull(TRIM(sp.DESCRIPTION2), '')) AS DESCRIPTION" & _
            ",sp.COST as 'NET OF VAT',sp.STOREPRICE as 'SRP',sp.DLRPRICE as 'WSP',sp.QTYONHAND as 'QTY ON HAND','0' as 'Six MSQ',sp.BACKORDER as 'BACK ORDER','' as 'LAST DATE IN'" & _
        ",sp.APPLICMODEL as 'APPLICACBLE MODEL',sp.MAXDISCOUNT,sp.SPAREPARTID " & _
        "From SPAREPARTS as sp WHERE sp.PARTNUMBER LIKE '%" & Trim(txtFindPart.Text) & "%' ORDER BY sp.PARTNUMBER;"
   
   
    If Qry = "" Then Exit Sub
    
   'InputBox Qry, Qry, Qry
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsParts.RecordCount = 0 Then
        flxFind.Clear
        flxFind.Rows = 2
        flxFind.FormatString = "|<PART NUMBER |<DESCRIPTION  |^NET OF VAT  |>SRP  |>WSP  |>QTY. |>SIX MSQ |>BACK ORDER |>LAST DATE IN |<APPLICABLE MODEL"
        Call flxFindColsChanger
    Else
        Set flxFind.DataSource = rsParts
        flxFind.Refresh
        flxFind.ColAlignment(3) = flexAlignRightCenter
        flxFind.ColAlignment(4) = flexAlignRightCenter
        flxFind.ColAlignment(5) = flexAlignRightCenter
        Call flxFind_Click
    End If
    
    Call FormatCur
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->FindPart()")
End Sub


Sub FormatCur()
    Dim rw As Integer
    For rw = 1 To flxFind.Rows - 1
        flxFind.TextMatrix(rw, 3) = Format(flxFind.TextMatrix(rw, 3), "##,###,###.#0")
        flxFind.TextMatrix(rw, 4) = Format(flxFind.TextMatrix(rw, 4), "##,###,###.#0")
        flxFind.TextMatrix(rw, 5) = Format(flxFind.TextMatrix(rw, 5), "##,###,###.#0")
    Next rw
End Sub


Sub ClearEntryForm()
   
    txtRefDate.Text = ""
    txtNetofVat.Text = ""
   
    txtPartNo.Text = ""
    txtDesc.Text = ""
    
    txtSRP.Text = ""
    txtSRPProfRatio.Text = ""
    txtSRPProfMar.Text = ""
    txtSRPFixPrice.Text = ""
    txtSRPMaxDis.Text = ""
    
    txtWSProfRatio.Text = ""
    txtWSProfMar.Text = ""
    txtWS.Text = ""
    txtWSFixPrice.Text = ""
    
    flxSRPPriceHis.Rows = 1
    flxSRPPriceHis.Rows = 2
    
    flxWSPriceHis.Rows = 1
    flxWSPriceHis.Rows = 2
    
    Check1(0).Value = 0
    Check1(1).Value = 0
    Check1(2).Value = 0
    Check1(3).Value = 0
    Check1(4).Value = 0
    Check1(5).Value = 0
                              
End Sub


Sub LockedEntryObject(DataEntry As Boolean)

    DataEntry = Not DataEntry
        
        txtRefNo.Locked = DataEntry
        txtSRPMaxDis.Locked = DataEntry
        txtRefDate.Locked = DataEntry
        txtNetofVat.Locked = DataEntry
        txtPartNo.Locked = DataEntry
        txtDesc.Locked = DataEntry
        txtSRP.Locked = DataEntry
        txtSRPProfMar.Locked = DataEntry
        txtSRPProfRatio.Locked = DataEntry
        txtSRPFixPrice.Locked = DataEntry
        txtSRPMaxDis.Locked = DataEntry
        txtWSProfMar.Locked = DataEntry
        txtWSProfRatio.Locked = DataEntry
        txtWSFixPrice.Locked = DataEntry
                                                                                                                                                                                                                                 
    If txtRefNo.Locked Then
        
        txtRefNo.BackColor = txtHoldGrayBackColor.BackColor
        txtSRPMaxDis.BackColor = txtHoldGrayBackColor.BackColor
        txtRefDate.BackColor = txtHoldGrayBackColor.BackColor
        txtNetofVat.BackColor = txtHoldGrayBackColor.BackColor
        txtPartNo.BackColor = txtHoldGrayBackColor.BackColor
        txtDesc.BackColor = txtHoldGrayBackColor.BackColor
        txtSRP.BackColor = txtHoldGrayBackColor.BackColor
        txtSRPProfMar.BackColor = txtHoldGrayBackColor.BackColor
        txtSRPProfRatio.BackColor = txtHoldGrayBackColor.BackColor
        txtSRPFixPrice.BackColor = txtHoldGrayBackColor.BackColor
        txtSRPMaxDis.BackColor = txtHoldGrayBackColor.BackColor
        txtWSProfMar.BackColor = txtHoldGrayBackColor.BackColor
        txtWSProfRatio.BackColor = txtHoldGrayBackColor.BackColor
        txtWSFixPrice.BackColor = txtHoldGrayBackColor.BackColor
        
    Else
        
        txtRefNo.BackColor = txtHoldWhiteBackColor.BackColor
        txtSRPMaxDis.BackColor = txtHoldWhiteBackColor.BackColor
        txtRefDate.BackColor = txtHoldWhiteBackColor.BackColor
        txtNetofVat.BackColor = txtHoldWhiteBackColor.BackColor
        txtPartNo.BackColor = txtHoldWhiteBackColor.BackColor
        txtDesc.BackColor = txtHoldWhiteBackColor.BackColor
        txtSRP.BackColor = txtHoldWhiteBackColor.BackColor
        txtSRPProfMar.BackColor = txtHoldWhiteBackColor.BackColor
        txtSRPProfRatio.BackColor = txtHoldWhiteBackColor.BackColor
        txtSRPFixPrice.BackColor = txtHoldWhiteBackColor.BackColor
        txtSRPMaxDis.BackColor = txtHoldWhiteBackColor.BackColor
        txtWSProfMar.BackColor = txtHoldWhiteBackColor.BackColor
        txtWSProfRatio.BackColor = txtHoldWhiteBackColor.BackColor
        txtWSFixPrice.BackColor = txtHoldWhiteBackColor.BackColor
        
    End If

End Sub
