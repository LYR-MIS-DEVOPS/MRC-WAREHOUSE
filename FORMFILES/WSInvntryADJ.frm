VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form WSInvntryADJ 
   Caption         =   "Wholesale Inventory Adjustments"
   ClientHeight    =   9285
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   10965
   Tag             =   "XYF"
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   120
      TabIndex        =   22
      Tag             =   "XYF"
      Top             =   7320
      Width           =   10695
      Begin VB.TextBox txtAprvBy 
         Height          =   285
         Left            =   7320
         TabIndex        =   36
         Tag             =   "XYF"
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox txtPrepBy 
         BackColor       =   &H80000003&
         Height          =   285
         Left            =   1560
         TabIndex        =   35
         Tag             =   "XYF"
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox txtRemarks 
         Height          =   285
         Left            =   240
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   1080
         Width           =   10335
      End
      Begin VB.TextBox txtReason 
         Height          =   285
         Left            =   240
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   480
         Width           =   10335
      End
      Begin VB.Label Label2 
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
         Left            =   240
         TabIndex        =   38
         Tag             =   "XYF"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
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
         Left            =   6000
         TabIndex        =   37
         Tag             =   "XYF"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lbl2 
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
         Left            =   240
         TabIndex        =   24
         Tag             =   "XYF"
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "REASON FOR ADJUSTMENT:"
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
         Left            =   240
         TabIndex        =   23
         Tag             =   "XYF"
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.TextBox txtChrgAmt 
      Height          =   285
      Left            =   7200
      TabIndex        =   33
      Top             =   10080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cboEmp 
      Height          =   315
      Left            =   5760
      TabIndex        =   32
      Top             =   10080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtRef 
      Height          =   285
      Left            =   7080
      MaxLength       =   20
      TabIndex        =   31
      Top             =   9360
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid flxPRN 
      Height          =   975
      Left            =   3120
      TabIndex        =   30
      Top             =   9360
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1720
      _Version        =   393216
      Cols            =   5
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "&A"
      Height          =   375
      Left            =   9240
      TabIndex        =   29
      Top             =   90000
      Width           =   135
   End
   Begin VB.TextBox txtPartNo 
      Height          =   285
      Left            =   3240
      TabIndex        =   28
      Top             =   10440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtQtyA 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4440
      TabIndex        =   27
      Top             =   10080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Left            =   3240
      TabIndex        =   26
      Top             =   10080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxParts 
      Height          =   1815
      Left            =   480
      TabIndex        =   25
      Top             =   9960
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3201
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtQtyTip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   360
      TabIndex        =   21
      Text            =   "Qty. Purchased: 0"
      Top             =   9480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtPrnCopy 
      Height          =   285
      Left            =   8040
      TabIndex        =   20
      Top             =   9360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   10080
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   10560
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Top             =   10080
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Top             =   10560
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Left            =   8760
      TabIndex        =   12
      Top             =   9360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1000
      Left            =   120
      TabIndex        =   4
      Tag             =   "XYF"
      Top             =   720
      Width           =   10695
      Begin VB.ComboBox cboADJSTATUS 
         Height          =   315
         Left            =   2040
         TabIndex        =   46
         Tag             =   "XYF"
         Top             =   538
         Width           =   1335
      End
      Begin VB.ComboBox txtAdjType 
         Height          =   315
         Left            =   2040
         TabIndex        =   34
         Tag             =   "XYF"
         Top             =   195
         Width           =   3015
      End
      Begin VB.TextBox txtAdjDate 
         Alignment       =   2  'Center
         BackColor       =   &H80000003&
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
         Left            =   8520
         TabIndex        =   3
         Tag             =   "XYF"
         Top             =   538
         Width           =   1815
      End
      Begin VB.TextBox txtAdjNo 
         Alignment       =   2  'Center
         BackColor       =   &H80000003&
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
         Left            =   8520
         TabIndex        =   0
         Tag             =   "XYF"
         Top             =   200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "ADJUSTMENT STATUS:"
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
         TabIndex        =   45
         Tag             =   "XYF"
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ADJUSTMENT DATE"
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
         Left            =   6720
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   570
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ADJUSTMENT NO."
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
         Left            =   6720
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   225
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "ADJUSTMENT TYPE:"
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
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   225
         Width           =   1815
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
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
      Image           =   "WSInvntryADJ.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   14
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "Su&mmary"
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
      Image           =   "WSInvntryADJ.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4620
      TabIndex        =   15
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Print"
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
      Image           =   "WSInvntryADJ.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6120
      TabIndex        =   16
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
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
      Image           =   "WSInvntryADJ.frx":10CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   7620
      TabIndex        =   17
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
      Image           =   "WSInvntryADJ.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   9120
      TabIndex        =   18
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
      Image           =   "WSInvntryADJ.frx":1C02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1620
      TabIndex        =   19
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
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
      Image           =   "WSInvntryADJ.frx":219C
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame3 
      Height          =   5655
      Left            =   120
      TabIndex        =   39
      Tag             =   "XYF"
      Top             =   1680
      Width           =   10695
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   8880
         TabIndex        =   48
         Tag             =   "XYF"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000003&
         Height          =   285
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   42
         Tag             =   "XYF"
         Top             =   5280
         Width           =   2655
      End
      Begin MSFlexGridLib.MSFlexGrid flxAdj 
         Height          =   4695
         Left            =   120
         TabIndex        =   40
         Tag             =   "XYF"
         Top             =   600
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   8281
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         RowHeightMin    =   315
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
         FormatString    =   "^|<PART NUMBER |<DESCRIPTION |^AVL QTY |^ QTY. ADJ. |>PRICE |>TOTAL"
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
      Begin lvButton.lvButtons_H btnFind 
         Height          =   315
         Left            =   7670
         TabIndex        =   49
         Tag             =   "XYF"
         Top             =   220
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "Find Next"
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
         Image           =   "WSInvntryADJ.frx":2736
         cBack           =   -2147483633
      End
      Begin VB.Label lblDatabase 
         BackStyle       =   0  'Transparent
         Caption         =   "STOCKS TO TRANSFER"
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
         TabIndex        =   43
         Tag             =   "TLF"
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label lblDatabase 
         BackStyle       =   0  'Transparent
         Caption         =   "STOCKS TO TRANSFER"
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
         TabIndex        =   44
         Tag             =   "TLF"
         Top             =   270
         Width           =   2895
      End
      Begin VB.Label lblTotal 
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
         Left            =   6840
         TabIndex        =   41
         Tag             =   "XYF"
         Top             =   5280
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   -240
      Top             =   1800
   End
   Begin VB.CommandButton cmdD 
      Caption         =   "&D"
      Height          =   375
      Left            =   9360
      TabIndex        =   47
      Top             =   90000
      Width           =   135
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   -240
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11205
      m_MDI_Height    =   9855
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ADJUSTMENTS"
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
      Left            =   9135
      TabIndex        =   10
      Tag             =   "TRF"
      Top             =   480
      Width           =   1440
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
      Top             =   450
      Width           =   2895
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
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ADJUSTMENTS"
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
      Left            =   9165
      TabIndex        =   11
      Tag             =   "TRF"
      Top             =   480
      Width           =   1440
   End
   Begin VB.Menu mnuADJ 
      Caption         =   "SpareParts"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Row         Alt+A"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "WSInvntryADJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsParts As New ADODB.Recordset
Dim rsMain As New ADODB.Recordset
Dim ActRow As Integer, ActCol As Integer, BtnIndx As Integer
Dim strPartial As String, strTotal As String, ans As String
Dim formAction As String
Dim StatSrch As Boolean
Dim AddNew As Boolean, LockFlx As Boolean, DispOnly As Boolean, PrtSrch As Boolean, PrtBln As Boolean
Dim CustSrch As Boolean
Dim fRow As Integer, fCol As Integer, TCol As Integer
Dim MyErr As String
Dim myMsg As String

Public Sub DisplayInvAdj(MyRef As String, MyAdjType As String)
    Dim rsAdj As New ADODB.Recordset
    Dim rsChrg As New ADODB.Recordset
    Dim rsDtls As New ADODB.Recordset
    Dim Qry As String, AdjType As String
    Dim AdjQ As Integer
    
    Call ClearEntries
    Qry = "call sp_mrcwarehouse.sp_stockadj_header('" & MyRef & "','" & MyAdjType & "')"
    
    With rsAdj
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        
        If .RecordCount > 0 Then
            .MoveFirst
            
            txtAdjType.Text = IIf(IsNull(!AdjType), "", !AdjType)
            txtADJNO.Text = IIf(IsNull(!ADJNO), "", !ADJNO)
            cboADJSTATUS.Text = IIf(IsNull(!ADJSTATUS), "", !ADJSTATUS)
            txtADJDATE.Text = IIf(IsNull(!ADJDATE), "", Format(!ADJDATE, "MM/DD/YYYY"))
            txtReason.Text = IIf(IsNull(!ADJREASON), "", !ADJREASON)
            txtRemarks.Text = IIf(IsNull(!ADJREMARKS), "", !ADJREMARKS)
            txtPrepBy.Text = IIf(IsNull(!PREPBY), "", !PREPBY)
            txtAprvBy.Text = IIf(IsNull(!ApprovedBy), "", !ApprovedBy)
        End If
        .Close
    End With
    
    Qry = "call sp_mrcwarehouse.sp_stockadj_details('" & MyRef & "')"
    With rsDtls
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        
        If .RecordCount > 0 Then
            .MoveFirst
            flxAdj.Rows = 1
            flxAdj.Rows = 2
            
            flxAdj.Rows = .RecordCount + 1
            
            For i = 1 To flxAdj.Rows - 1
                flxAdj.TextMatrix(i, 1) = IIf(IsNull(!PartNumber), "", !PartNumber)
                flxAdj.TextMatrix(i, 2) = IIf(IsNull(!Description), "", !Description)
                flxAdj.TextMatrix(i, 3) = IIf(IsNull(!AVLQTY), 0, !AVLQTY)
                flxAdj.TextMatrix(i, 4) = IIf(IsNull(!QTYADJ), 0, !QTYADJ)
                flxAdj.TextMatrix(i, 5) = IIf(IsNull(!Price), 0, Format(!Price, "#,###,###.00"))
                flxAdj.TextMatrix(i, 6) = IIf(IsNull(!Ttl), 0, Format(!Ttl, "#,###,###.00"))
                flxAdj.TextMatrix(i, 7) = IIf(IsNull(!SPAREPARTID), 0, !SPAREPARTID)
                .MoveNext
            Next i
        End If
        .Close
    End With
    Call ComputeTotal
End Sub

Sub CreateTempTable()
On Error GoTo er
       LYRcon.Execute "Drop Temporary Table IF  Exists sp_mrcwarehouse.TMP_STOCKADJ_DTLS;"

       DoEvents
       LYRcon.Execute " Create Temporary Table IF Not Exists sp_mrcwarehouse.TMP_STOCKADJ_DTLS( " & _
        "    SPAREPARTID    Int(11)             " & _
        "   ,AVLQTY         Int(11)             " & _
        "   ,QTYADJ         Int(11)             " & _
        "   ,PRICE          Decimal(10,2)       " & _
        "   ,SEQNO          Int(11)             " & _
        ");"
        
        Qry = "Insert into sp_mrcwarehouse.TMP_STOCKADJ_DTLS(SPAREPARTID,AVLQTY,QTYADJ,PRICE,SEQNO) VALUES "
        For i = 1 To flxAdj.Rows - 1
            If i > 1 Then Qry = Qry & ","
            Qry = Qry & "(" & Val(flxAdj.TextMatrix(i, 7)) & "," & Val(flxAdj.TextMatrix(i, 3)) & "," & Val(flxAdj.TextMatrix(i, 4)) & ", '" & Replace(flxAdj.TextMatrix(i, 5), ",", "") & "','" & i & "')"
        Next i
       
       Call LYRcon.Execute(Qry)
       
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CreateTempTable()")
End Sub

Function ValidationSave() As String
    MyErr = ""
    If flxAdj.TextMatrix(flxAdj.Rows - 1, 1) = "" Then
        If flxAdj.Rows = 2 Then
            flxAdj.Rows = 1
            flxAdj.Rows = 2
            MyErr = MyErr & vbCrLf & "No Record Found!"
        Else
            flxAdj.RemoveItem flxAdj.Rows - 1
        End If
    End If
    
    'If txtAdjNo.Text = "" Then MyErr = MyErr & vbCrLf & "No Adjustment Number Found!"
    If txtAprvBy.Text = "" Then MyErr = MyErr & vbCrLf & "No Approved By Found!"
    
    If flxAdj.Rows > 2 Then
        If flxAdj.TextMatrix(flxAdj.Rows - 1, 1) = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on Last Row"
            
        For i = 1 To flxAdj.Rows - 1
            If flxAdj.TextMatrix(i, 4) <= 0 Then
                MyErr = MyErr & vbCrLf & "0 Adjustment Quanity @ row" & i
            End If
        Next i
    End If
    ValidationSave = MyErr
End Function


Private Sub SetButtonVisibility(DataEntry As Boolean)
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->VIEW SUMM
'3 ->EXIT
'4 ->SAVE
'5 ->CANCEL


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


Private Sub ClearEntries()
    txtAdjType.Text = ""
    txtADJNO.Text = "<AUTO>"
    txtADJDATE.Text = Format$(ServerTime, "MM/DD/YYYY")
    txtReason.Text = ""
    txtTotal.Text = ""
    txtRemarks.Text = ""
    txtPrepBy.Text = ""
    txtAprvBy.Text = ""
    flxAdj.Rows = 1
    flxAdj.Rows = 2
    cboADJSTATUS.Text = ""
End Sub

Private Sub LockData(Bln As Boolean)
    Frame1.Enabled = Not Bln
    Frame2.Enabled = Not Bln
    txtPrepBy.Locked = Bln
    txtAprvBy.Locked = Bln
    txtADJDATE.Locked = Not Bln
    txtADJNO.Locked = Not Bln
    LockFlx = Bln
    'chkCharges.Enabled = Not Bln
End Sub

Public Sub ComputeTotal()
    Dim GrdTtl As Single, i As Integer
    Dim TtlPrice As Currency
    
    GrdTtl = 0
    For i = 1 To flxAdj.Rows - 1
        TtlPrice = Val(flxAdj.TextMatrix(i, 4)) * PureNum(flxAdj.TextMatrix(i, 5))
        GrdTtl = GrdTtl + TtlPrice
        flxAdj.TextMatrix(i, 6) = Format(TtlPrice, "#,##0.00")
    Next i
    txtTotal.Text = Format$(GrdTtl, "#,##0.00")
    
End Sub


Private Sub btnFind_Click()
    If txtFind.Text = "" Then Exit Sub
    With flxAdj
        For i = fRow To .Rows - 1
            For j = fCol To TCol - 1
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 0: .row = i
                    .ColSel = 7:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = TCol - 1, i + 1, i)
                    fCol = IIf(j = TCol - 1, 1, j + 1)
                    ActRow = 1
                    Exit Sub
                End If
            Next j
            fCol = 1
        Next i
    End With
    MsgBox "No match found!", vbInformation + vbOKOnly, "Find"
    fRow = 1
    fCol = 1
End Sub

Private Sub cmdA_Click()
    Call mnuAdd_Click
End Sub

Private Sub LoadSpareParts()
    Dim Qry As String
        
    If txtAdjType.Text = "WHOLESALE TO RETAIL" Then
        Qry = "SELECT  SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, IFNULL(SPAREPARTS.WSQTYONHAND,0) as QTYONHAND " & _
              "From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
    ElseIf txtAdjType.Text = "RETAIL TO WHOLESALE" Then
        Qry = "SELECT  SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE,  CASE WHEN (SPAREPARTS.QTYONHAND - IFNULL(SPAREPARTS.WSQTYONHAND,0)) > 0 THEN (SPAREPARTS.QTYONHAND - IFNULL(SPAREPARTS.WSQTYONHAND,0)) ELSE 0 END as QTYONHAND " & _
              "From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
    End If
    If rsParts.State = adStateOpen Then rsParts.Close
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxParts.DataSource = rsParts
    With flxParts
        .Width = flxAdj.ColWidth(1)
        .ColWidth(0) = 0
        .ColWidth(1) = 0    'SPid
        .ColWidth(2) = flxAdj.ColWidth(1)   'PartNo
        .ColWidth(3) = 0    'Description
        .ColWidth(4) = 0    'Store Price
        .ColWidth(5) = 0    'Qty On Hand
    End With
End Sub


Private Sub cmdD_Click()
Call mnuDelete_Click
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Select Case Index
        Case 0  'Add new
            formAction = "add"
            Call ClearEntries
            Call LockData(False)
            Call SetButtonVisibility(True)
            txtPrepBy.Text = EmployeesName
            cboADJSTATUS.Text = "PENDING"
            txtAdjType.SetFocus
        Case 1  'Edit
            If cboADJSTATUS.Text <> "FINAL" Then
                formAction = "edit"
                Call LockData(False)
                Call SetButtonVisibility(True)
                txtAdjType.SetFocus
                If txtAdjType.Text <> "" Then
                    Call LoadSpareParts
                End If
            Else
                MsgBox "The Adjustment was already Final!", vbInformation + vbOKOnly, "Validation"
                Exit Sub
            End If
        Case 2  'Summary
            FrmSummary.mySrc = "WS Inventory Adjustment"
            FrmSummary.Show
            Unload Me
            Exit Sub
        Case 3  'Print
            ans = MsgBox("Do you want to Print Inventory Adjustment?", vbQuestion + vbYesNo, "Print")
            If ans = vbNo Then Exit Sub
            Qry = "call sp_mrcwarehouse.sp_stockadj_details('" & txtADJNO.Text & "')"
            
            PRNmc.QryPrn = Qry
            PRNmc.ParamValue1 = SysBranch
            PRNmc.ParamValue2 = SysAddress  'cboType.Text
            PRNmc.ParamValue3 = txtAdjType.Text
            PRNmc.ParamValue4 = txtADJNO.Text
            PRNmc.ParamValue5 = cboADJSTATUS.Text
            PRNmc.ParamValue6 = txtADJDATE.Text
            PRNmc.ParamValue7 = txtReason.Text
            PRNmc.ParamValue8 = txtRemarks.Text
            PRNmc.ParamValue9 = txtPrepBy.Text
            PRNmc.ParamValue10 = txtAprvBy.Text
            PRNmc.MySource = "WS Inventory Adjustment"
            PRNmc.Show
        Case 4  'Exit
            Unload Me
        Case 5  'Save
            myMsg = ValidationSave
            
            If myMsg <> "" Then
                MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
        
            ans = MsgBox("Do you want to save?", vbQuestion + vbYesNo, "Save")
            If ans = vbNo Then Exit Sub
            
            Call CreateTempTable
            If formAction = "add" Then
                Qry = "call sp_mrcwarehouse.sp_stockadj_create('" & txtAdjType.Text & "','" & cboADJSTATUS.Text & "','" & txtReason.Text & "','" & txtRemarks.Text & "','" & UserID & "','" & txtAprvBy.Text & "')"
                
                With rsMain
                    If .State = adStateOpen Then .Close
                    .Open Qry, LYRcon, adOpenStatic, adLockOptimistic
                    
                    If .RecordCount > 0 Then
                        .MoveFirst
                        txtADJNO.Text = !ADJNO
                        If !level = "Granted" Then
                            MsgBox "" & !message & "", vbInformation + vbOKOnly, "SAVE"
                        Else
                            MsgBox "" & !message & "", vbInformation + vbOKOnly, "SAVE"
                            .Close
                            Exit Sub
                        End If
                    End If
                    .Close
                End With
                
            ElseIf formAction = "edit" Then
                Qry = "call sp_mrcwarehouse.sp_stockadj_update('" & txtAdjType.Text & "','" & txtADJNO.Text & "','" & cboADJSTATUS.Text & "','" & txtReason.Text & "','" & txtRemarks.Text & "','" & UserID & "','" & txtAprvBy.Text & "')"
                
                With rsMain
                    If .State = adStateOpen Then .Close
                    .Open Qry, LYRcon, adOpenStatic, adLockOptimistic
                    
                    If .RecordCount > 0 Then
                        .MoveFirst
                        If !level = "Granted" Then
                            MsgBox "" & !message & "", vbInformation + vbOKOnly, "SAVE"
                        Else
                            MsgBox "" & !message & "", vbInformation + vbOKOnly, "SAVE"
                            .Close
                            Exit Sub
                        End If
                    End If
                    .Close
                End With
            End If
            
            Call SetButtonVisibility(False)
            Call LockData(True)
            formAction = ""
        Case 6  'Cancel
            ans = MsgBox("Do you want to cancel?", vbQuestion + vbYesNo)
            
            If ans = vbNo Then Exit Sub
            Call SetButtonVisibility(False)
            Call LockData(True)
            formAction = ""
    End Select
End Sub

Private Sub flxAdj_Click()
    ActRow = flxAdj.row
    ActCol = flxAdj.Col
    
    If ActCol = 0 Or ActRow = 0 Or LockFlx Then Exit Sub
    If txtAdjType.Text <> "" Then
        Select Case ActCol
            Case 1 'PartNumber
                With txtPartNo
                    If flxAdj.TextMatrix(ActRow, 1) <> "" Then flxAdj.TextMatrix(ActRow, 8) = flxAdj.TextMatrix(ActRow, 1) & "|" & flxAdj.TextMatrix(ActRow, 2) & "|" & flxAdj.TextMatrix(ActRow, 3) & "|" & flxAdj.TextMatrix(ActRow, 4) & "|" & flxAdj.TextMatrix(ActRow, 5) & "|" & flxAdj.TextMatrix(ActRow, 6) & "|" & flxAdj.TextMatrix(ActRow, 7)
                    .Top = Frame3.Top + flxAdj.CellTop + flxAdj.Top
                    .Left = Frame3.Left + flxAdj.CellLeft + flxAdj.Left
                    .Width = flxAdj.CellWidth
                    .Text = flxAdj.Text
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .ZOrder
                    .SetFocus
                End With
            Case 4    'qty
                With txtQtyA
                    If flxAdj.TextMatrix(ActRow, 4) <> "" Then .Tag = flxAdj.TextMatrix(ActRow, 4)
                    .Top = Frame3.Top + flxAdj.CellTop + flxAdj.Top
                    .Left = Frame3.Left + flxAdj.CellLeft + flxAdj.Left
                    .Width = flxAdj.CellWidth
                    .Text = flxAdj.Text
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .ZOrder
                    .SetFocus
                End With
        End Select
    End If
End Sub

Private Sub flxAdj_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxAdj_Click
End Sub

Private Sub flxAdj_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If LockFlx Then Exit Sub
    If Button = vbRightButton And ActCol = 0 Then
'        If txtAdjType.Text <> "BACK ORDER" Then
'            If flxAdj.TextMatrix(ActRow, 7) = "" Then
'                mnuCardex.Enabled = False
'                mnuCardexRed.Enabled = False
'            Else
'                mnuCardex.Enabled = True
'                mnuCardexRed.Enabled = True
'            End If
            Call PopupMenu(mnuADJ)
'        Else
'            If flxAdj.TextMatrix(ActRow, 8) = "" Then
'                mnuBOCardex.Enabled = False
'                mnuBOCardexRed.Enabled = False
'            Else
'                mnuBOCardex.Enabled = True
'                mnuBOCardexRed.Enabled = True
'            End If
'            Call PopupMenu(mnuSPBOrder)
'        End If
    End If
End Sub

Private Sub flxAdj_Scroll()
    txtPartNo.Visible = False
    txtQtyA.Visible = False
End Sub

Private Sub flxParts_Click()
    Dim TtlPrice As Currency
    
    With flxParts
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.row, 2)
        'Insert Additional retrieval here
        flxAdj.TextMatrix(ActRow, 2) = .TextMatrix(.row, 3) 'Description
        flxAdj.TextMatrix(ActRow, 3) = .TextMatrix(.row, 5) 'Qtyonhand
        flxAdj.TextMatrix(ActRow, 5) = Format$(Val(.TextMatrix(.row, 4)), "#,##0.00") 'Price
        flxAdj.TextMatrix(ActRow, 7) = .TextMatrix(.row, 1) 'SPid
        'Qty Adj Available
        If Val(flxAdj.TextMatrix(ActRow, 4)) > 0 Then
            TtlPrice = Val(flxAdj.TextMatrix(ActRow, 4)) * Val(.TextMatrix(.row, 4))
            flxAdj.TextMatrix(ActRow, 6) = Format(TtlPrice, "#,###,##0.00")
        Else
            flxAdj.TextMatrix(ActRow, 6) = "0.00"
        End If
        Call ComputeTotal
    End With
End Sub

Private Sub flxParts_EnterCell()
    Dim TtlPrice As Currency
    
    With flxParts
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.row, 2)
        'Insert Additional retrieval here
        flxAdj.TextMatrix(ActRow, 2) = .TextMatrix(.row, 3) 'Description
        flxAdj.TextMatrix(ActRow, 3) = .TextMatrix(.row, 5) 'Qtyonhand
        flxAdj.TextMatrix(ActRow, 5) = Format$(Val(.TextMatrix(.row, 4)), "#,##0.00") 'Price
        flxAdj.TextMatrix(ActRow, 7) = .TextMatrix(.row, 1) 'SPid
        'Qty and SRP >0
        If Val(flxAdj.TextMatrix(ActRow, 4)) > 0 Then
            TtlPrice = Val(flxAdj.TextMatrix(ActRow, 4)) * Val(.TextMatrix(.row, 4))
            flxAdj.TextMatrix(ActRow, 6) = Format(TtlPrice, "#,###,##0.00")
        Else
            flxAdj.TextMatrix(ActRow, 6) = "0.00"
        End If
        Call ComputeTotal
    End With
End Sub

Private Sub flxParts_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Tmp() As String
    Dim i As Integer
    
    If KeyCode = vbKeyEscape Then
        Tmp = Split(flxParts.Tag, "|")
        For i = LBound(Tmp) To UBound(Tmp)
            flxAdj.TextMatrix(ActRow, i + 1) = Tmp(i)
        Next i
        txtPartNo.Visible = False
        flxParts.Visible = False
        flxAdj.SetFocus
        flxAdj.Col = 1
    End If

End Sub

Private Sub flxParts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtPartNo_LostFocus
        If flxAdj.TextMatrix(ActRow, 7) = "" Then   'No SPID
            flxAdj.SetFocus
            flxAdj.row = ActRow
            flxAdj.Col = 1
        Else
            flxAdj.SetFocus
            flxAdj.row = ActRow
            flxAdj.Col = 4
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    Call LockData(True)
    
    TCol = 7
    txtADJDATE.Text = Format$(ServerTime, "MM/DD/YYYY")
    txtADJNO.Text = "<AUTO>"
    
    txtAdjType.AddItem ""
    txtAdjType.AddItem "WHOLESALE TO RETAIL"
    txtAdjType.AddItem "RETAIL TO WHOLESALE"
    
    cboADJSTATUS.AddItem "PENDING"
    cboADJSTATUS.AddItem "FINAL"
    '===================BUTTON POSITION===================
    
    If cmdTrans(0).Visible = False Then
            'Align Save Button
            cmdTrans(5).Left = cmdTrans(0).Left
            cmdTrans(5).Top = cmdTrans(0).Top
            'Align Cancel Button
            cmdTrans(6).Left = cmdTrans(1).Left
            cmdTrans(6).Top = cmdTrans(1).Top
     End If
    
    cmdA.Top = 15000

    Warehouse.sBar.Panels(6).Text = "Spare Parts Inventory Adjustment Form"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Warehouse.sBar.Panels(6).Text = ""
    If rsParts.State = adStateOpen Then rsParts.Close
End Sub



Private Sub mnuAdd_Click()
    If flxAdj.TextMatrix(ActRow, 7) = "" Or flxAdj.TextMatrix(ActRow, 6) = "" Or Val(flxAdj.TextMatrix(ActRow, 4)) = 0 Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        flxAdj.SetFocus
        flxAdj.Col = ActCol
        Exit Sub
    Else
        If flxAdj.Rows = 6 Then
            MsgBox "You can only add upto 5 Part Numbers per Adjustment!", vbExclamation + vbOKOnly, "Limit Exceed"
            Exit Sub
        End If
        If ActRow = flxAdj.Rows - 1 Then flxAdj.AddItem ""
        flxAdj.SetFocus
        flxAdj.Col = 1
        flxAdj.row = flxAdj.Rows - 1
    End If
End Sub



Private Sub mnuDelete_Click()
    Dim ans As Integer, i As Integer
    
    ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If ans = vbYes Then
        If ActRow = 1 And flxAdj.Rows = 2 Then
            flxAdj.Rows = 1
            flxAdj.Rows = 2
        Else
            flxAdj.RemoveItem ActRow
        End If
        Call ComputeTotal
    End If
End Sub

Private Sub RXResizer1_ResizeComplete()
    With flxAdj
        .ColWidth(0) = 350
        .ColWidth(1) = (.Width - 380) * 0.2
        .ColWidth(2) = (.Width - 380) * 0.34
        .ColWidth(3) = (.Width - 380) * 0.12
        .ColWidth(4) = (.Width - 380) * 0.12
        .ColWidth(5) = (.Width - 380) * 0.1
        .ColWidth(6) = (.Width - 380) * 0.1
        .ColWidth(7) = 0 '(.Width - 380) * 0.1
        .ColWidth(8) = 0 '(.Width - 380) * 0.1
        'txtTotal.Left = .Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + .ColWidth(5) + .ColWidth(6)
    End With
    txtTotal.Top = flxAdj.Height + flxAdj.Top
    txtTotal.Left = flxAdj.Left + (flxAdj.Width - flxAdj.ColWidth(6))
    txtTotal.Width = flxAdj.ColWidth(6)
    lblTotal.Top = txtTotal.Top
    lblTotal.Left = txtTotal.Left - lblTotal.Width 'flxadj.Left + (flxadj.Width - (flxadj.ColWidth(4) + flxadj.ColWidth(5)))

End Sub


Function GetAdjustmentNo() As String
    Dim rsAdjCntr As New ADODB.Recordset
    Dim Qry As String
    Dim MyNo As Long
    
    'Access
    'Qry = "SELECT Max(ADJCNTR.ADJNO) AS MaxOfADJNO FROM ADJCNTR;"
    'mySQL
    If txtAdjType.Text <> "BACK ORDER" Then
        Qry = "SELECT Max(round(ADJCNTR.ADJNO)) AS MaxOfADJNO FROM ADJCNTR WHERE ADJTYPE = 'INVENTORY';"
    Else
        'Qry = "SELECT Max(ADJCNTR.ADJNO) AS MaxOfADJNO FROM ADJCNTR WHERE ADJTYPE = 'BACK ORDER';"
        Qry = "SELECT COUNT(ADJNO) as ADJNO From invntryadj Where MONTH(ADJDATE) = MONTH(NOW()) AND ADJTYPE = 'BACK ORDER'"
    End If
    With rsAdjCntr
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MyNo = IIf(IsNull(.Fields(0)), 1, .Fields(0) + 1)
        Else
            MyNo = 1
        End If
        .Close
    End With
    If txtAdjType.Text <> "BACK ORDER" Then
        GetAdjustmentNo = Format(MyNo, "0000")
    Else
        GetAdjustmentNo = Format(MyNo, "0000") & "-BO"
    End If
End Function



Private Sub tmrDrop_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 7 Then
            tmrDrop.Enabled = False
        End If
    End If
End Sub

Private Sub tmrDropRev_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 5 Then
            tmrDropRev.Enabled = False
            'cmdTrans(0).SetFocus
        End If
    End If
End Sub

Private Sub tmrFloat_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top - 50
    If cmdTrans(BtnIndx).Top <= -375 Then
        cmdTrans(BtnIndx).Visible = False
        BtnIndx = BtnIndx - 1
        If BtnIndx = -1 Then
            tmrFloat.Enabled = False
            BtnIndx = 5 'Index of Button Save
            cmdTrans(5).Top = -375
            cmdTrans(5).Left = 0
            cmdTrans(5).Visible = True
            cmdTrans(6).Top = -375
            cmdTrans(6).Left = cmdTrans(4).Width
            cmdTrans(6).Visible = True
            tmrDrop.Enabled = True
        End If
    End If
End Sub

Private Sub tmrFloatRev_Timer()
    Dim i As Integer
    
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top - 50
    If cmdTrans(BtnIndx).Top <= -375 Then
        cmdTrans(BtnIndx).Visible = False
        BtnIndx = BtnIndx - 1
        If BtnIndx = 4 Then 'Index of Button Exit
            tmrFloatRev.Enabled = False
            BtnIndx = 0
            For i = 0 To 4
                cmdTrans(i).Visible = True
            Next i
            tmrDropRev.Enabled = True
        End If
    End If
End Sub


Private Sub txtAdjDate_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890/-"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        flxAdj.SetFocus
        flxAdj.row = 1
        flxAdj.Col = 2
    End If
End Sub

Private Sub txtAdjNo_GotFocus()
    If Val(txtADJNO.Text) > 0 Then
        txtADJNO.SelStart = 0
        txtADJNO.SelLength = Len(txtADJNO.Text)
    End If
End Sub

Private Sub txtAdjNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtADJDATE.SetFocus
End Sub

Private Sub txtAdjNo_LostFocus()
    Dim Qry As String
    
    'Verify AdjNo if existing
    If Val(txtADJNO.Text) > 0 And AddNew Then
        '***Access
        'Qry = "SELECT ADJCNTR.ADJNO From ADJCNTR WHERE (((ADJCNTR.ADJNO)=" & Val(txtAdjNo.Text) & "));"
        'mySQL
        Qry = "SELECT STOCKSADJUSTMENT.ADJNO From STOCKSADJUSTMENT WHERE STOCKSADJUSTMENT.ADJNO='" & txtADJNO.Text & "';"
        
        If CtrlNumExist(Qry) Then
            MsgBox "Adjustment No. " & txtADJNO.Text & " already exist!", vbExclamation + vbOKOnly, "Validation Error"
        End If
    End If
    
End Sub

Private Sub txtAdjType_Change()
On Error GoTo er
    If StatSrch Then
            StatSrch = False
            Exit Sub
    End If
    'catID = 0
    With txtAdjType
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtAdjType-Change()")
End Sub

Private Sub txtAdjType_Click()
If txtAdjType.Text <> "" Then
    Call LoadSpareParts
    txtADJNO.SetFocus
    If txtAdjType.Text = "WHOLESALE TO RETAIL" Then
        lblDatabase(0).Caption = "WHOLESALE STOCKS"
        lblDatabase(1).Caption = "WHOLESALE STOCKS"
        flxAdj.Rows = 1
        flxAdj.Rows = 2
    ElseIf txtAdjType.Text = "RETAIL TO WHOLESALE" Then
        lblDatabase(0).Caption = "WAREHOUSE STOCKS"
        lblDatabase(1).Caption = "WAREHOUSE STOCKS"
        flxAdj.Rows = 1
        flxAdj.Rows = 2
    End If
Else
    If txtAdjType.Text = "" Then
        lblDatabase(0).Caption = "STOCKS TO TRANSFER"
        lblDatabase(1).Caption = "STOCKS TO TRANSFER"
    End If
End If
End Sub

Private Sub txtAdjType_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        Call txtAdjType_Click
    End If
End Sub

Private Sub txtaprvby_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If LockFlx Then
            cmdTrans(0).SetFocus
        Else
            cmdTrans(5).SetFocus
        End If
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(txtFind, KeyAscii, ValidChars)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = 13 Then Call btnFind_Click
End Sub

Private Sub txtPartNo_Change()
    Dim TtlPrice As Currency, DiscPrice As Currency
    Dim PctDisc As Single
    
    With flxParts
        If PrtSrch Then
            PrtSrch = False
            Exit Sub
        End If
        .Top = txtPartNo.Height + txtPartNo.Top
        .Left = txtPartNo.Left
        .Width = txtPartNo.Width
        .Visible = True
        .ZOrder
        If txtPartNo.Text = "" Then .Visible = False: Exit Sub
        rsParts.Find "[PARTNUMBER] LIKE '" + txtPartNo.Text + "%'", , adSearchForward, 1
        If Not rsParts.EOF Then
            If Not PrtBln Then
                txtPartNo.Tag = Len(txtPartNo.Text)
                PrtBln = True
            End If
            .TopRow = rsParts.AbsolutePosition - 1
            .row = rsParts.AbsolutePosition - 1
            .Col = 0: .ColSel = 5
            txtPartNo.Text = rsParts!PartNumber
            'Insert Additional retrieval here
            flxAdj.TextMatrix(ActRow, 7) = flxParts.TextMatrix(flxParts.row, 1) 'SP id
            flxAdj.TextMatrix(ActRow, 2) = flxParts.TextMatrix(flxParts.row, 3) 'Description
            flxAdj.TextMatrix(ActRow, 3) = flxParts.TextMatrix(flxParts.row, 5) 'Qtyonhand
            flxAdj.TextMatrix(ActRow, 5) = Format$(Val(flxParts.TextMatrix(flxParts.row, 4)), "###,##0.00") 'Price
            If Val(flxAdj.TextMatrix(ActRow, 4)) > 0 Then   'Qty Adj.
                TtlPrice = Val(flxAdj.TextMatrix(ActRow, 4)) * Val(flxParts.TextMatrix(flxParts.row, 4))
                flxAdj.TextMatrix(ActRow, 6) = Format(TtlPrice, "#,###,##0.00")
                Call ComputeTotal
            Else
                flxAdj.TextMatrix(ActRow, 6) = "0.00"
            End If
            txtPartNo.SelStart = Val(txtPartNo.Tag)
            txtPartNo.SelLength = Len(txtPartNo.Text)
            PrtBln = False
        Else
'            rsParts.Requery
            flxAdj.TextMatrix(ActRow, 2) = ""
            flxAdj.TextMatrix(ActRow, 3) = ""
            flxAdj.TextMatrix(ActRow, 4) = ""
            flxAdj.TextMatrix(ActRow, 5) = ""
            flxAdj.TextMatrix(ActRow, 6) = ""
            flxAdj.TextMatrix(ActRow, 7) = ""
            .Visible = False
        End If

    End With
End Sub

Private Sub txtPartNo_GotFocus()
    With flxParts
        .Top = txtPartNo.Height + txtPartNo.Top
        .Left = txtPartNo.Left
        .Width = txtPartNo.Width
    End With
End Sub

Private Sub txtPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim Tmp() As String
    
    Select Case KeyCode
        Case vbKeyDown
            flxParts.Tag = "skiplostfocus"
            flxParts.Visible = True
            flxParts.SetFocus
        Case vbKeyEscape
            txtPartNo.Text = Empty
            Tmp = Split(flxAdj.TextMatrix(ActRow, 7), "|")
            For i = LBound(Tmp) To UBound(Tmp)
                flxAdj.TextMatrix(ActRow, i + 1) = Tmp(i)
            Next i
            flxParts.Tag = "skiplostfocus"
            flxAdj.SetFocus
            flxAdj.Col = 1
            txtPartNo.Visible = False
            flxParts.Visible = False
        Case vbKeyDelete
            PrtSrch = True
        Case vbKeyBack
            PrtSrch = True
    End Select
End Sub

Private Sub txtPartNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If flxAdj.TextMatrix(ActRow, 7) = "" Then   'No SPID
            flxAdj.SetFocus
            flxAdj.row = ActRow
            flxAdj.Col = 1
        Else
            flxAdj.SetFocus
            flxAdj.row = ActRow
            flxAdj.Col = 4
        End If
        flxParts.Visible = False
    End If
End Sub

Private Sub txtPartNo_LostFocus()
    If flxParts.Tag = "skiplostfocus" Then
        flxParts.Tag = ""
        Exit Sub
    End If
    If flxAdj.TextMatrix(ActRow, 7) = "" Then
        MsgBox "Unrecognized Part Number!" & vbCrLf & "Please select a Part Number from list.", vbExclamation + vbOKOnly, "Warning"
        txtPartNo.Visible = False
        flxParts.Visible = False
        Exit Sub
    End If
    flxAdj.TextMatrix(ActRow, 1) = txtPartNo.Text
    txtPartNo.Visible = False
    flxParts.Visible = False
End Sub

Private Sub txtPrepBy_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtAprvBy.SetFocus
End Sub

Private Sub txtQtyA_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtQtyA.Text = Empty
        flxAdj.TextMatrix(ActRow, 4) = txtQtyA.Tag
        txtQtyA.Tag = "skiplostfocus"
        flxAdj.SetFocus
        flxAdj.Col = 4
        txtQtyA.Visible = False
    End If
End Sub

Private Sub txtQtyA_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        With flxAdj
            .SetFocus
            If .TextMatrix(ActRow, 7) <> "" And Val(txtQtyA.Text) > 0 And Val(txtQtyA.Text) <= Val(.TextMatrix(ActRow, 3)) Then
                If ActRow = .Rows - 1 Then .AddItem ""
                .row = .Rows - 1
                .Col = 1
            ElseIf .TextMatrix(ActRow, 7) = "" Then
                .row = ActRow
                .Col = 1
            ElseIf .TextMatrix(ActRow, 4) = "" Then
                .row = ActRow
                .Col = 4
            End If
        End With
    End If
End Sub

Private Sub txtQtyA_LostFocus()
    If txtQtyA.Tag = "skiplostfocus" Then
        txtQtyA.Tag = ""
        Exit Sub
    End If
    
    If flxAdj.TextMatrix(ActRow, 7) <> "" And Val(flxAdj.TextMatrix(ActRow, 4)) >= Val(flxAdj.TextMatrix(ActRow, 3)) Then
        MsgBox "QUANTITY ADJUSTMENT IS GREATER THAN THE AVAILABLE QUANTITY! PLEASE ENTER AN ALLOWED QUANTITY", vbInformation + vbOKOnly, "VALIDATION"
        flxAdj.row = ActRow
        flxAdj.Col = 4
        flxAdj.TextMatrix(ActRow, 4) = 0
        txtQtyA.Visible = False
        Exit Sub
    End If
    flxAdj.TextMatrix(ActRow, 4) = txtQtyA.Text
    Call ComputeTotal
    txtQtyA.Visible = False
End Sub

Private Sub txtReason_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtRemarks.SetFocus
End Sub
Function GetPartsMan(MyID As Integer) As String
    Dim rsUser As New ADODB.Recordset
    Dim Qry As String
    '***Access
    'Qry = "SELECT PCUSER.USERNAME, PCUSER.EMPID From PCUSER WHERE (((PCUSER.EMPID)=" & MyId & "));"
    'mySQL
    Qry = "SELECT PCUSER.USERNAME, PCUSER.EMPID From PCUSER WHERE PCUSER.EMPID=" & MyID & ";"
    
    With rsUser
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetPartsMan = IIf(IsNull(!UserName), "", UCase$(!UserName))
        Else
            GetPartsMan = ""
        End If
        .Close
    End With
End Function

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAprvBy.SetFocus
End Sub
