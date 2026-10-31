VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form BranchPO 
   Caption         =   "BRANCH P.O ORDER TAKING"
   ClientHeight    =   9015
   ClientLeft      =   2595
   ClientTop       =   -18210
   ClientWidth     =   13875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   13875
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboAprBy 
      Height          =   315
      Left            =   9720
      TabIndex        =   56
      Tag             =   "XYF"
      Top             =   7560
      Width           =   3975
   End
   Begin VB.TextBox txtOrderType 
      Height          =   375
      Left            =   13920
      TabIndex        =   47
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtDateCon 
      Height          =   285
      Left            =   14400
      TabIndex        =   46
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtOrder 
      Height          =   240
      Left            =   1560
      TabIndex        =   23
      Tag             =   "XYF"
      Top             =   9720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTotalAlloc 
      Height          =   405
      Left            =   4080
      TabIndex        =   44
      Tag             =   "XYF"
      Top             =   9600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAllocQty 
      Height          =   375
      Left            =   11160
      TabIndex        =   43
      Top             =   9720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtPrnCopy 
      Height          =   375
      Left            =   11160
      TabIndex        =   42
      Top             =   9240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2055
      Left            =   8160
      TabIndex        =   36
      Top             =   9480
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3625
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtBranchSPOID"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtBranchID"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtSPOID"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.TextBox txtSPOID 
         Height          =   405
         Left            =   240
         TabIndex        =   39
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtBranchID 
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtBranchSPOID 
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.TextBox txtRemarks 
      Height          =   315
      Left            =   1830
      TabIndex        =   32
      Tag             =   "XYF"
      Top             =   7260
      Width           =   7365
   End
   Begin VB.TextBox txtPreparedby 
      Height          =   315
      Left            =   1830
      TabIndex        =   30
      Tag             =   "XYF"
      Top             =   7650
      Width           =   7365
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   135
      Top             =   9390
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   135
      Top             =   9870
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   735
      Top             =   9390
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   735
      Top             =   9870
   End
   Begin VB.TextBox txtTotalAmount 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   12150
      TabIndex        =   10
      Tag             =   "XYF"
      Text            =   "0.00"
      Top             =   6840
      Visible         =   0   'False
      Width           =   1470
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   300
      TabIndex        =   4
      Tag             =   "XYF"
      Top             =   6900
      Visible         =   0   'False
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame2 
      Caption         =   "Annotation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   945
      Left            =   300
      TabIndex        =   9
      Tag             =   "XYF"
      Top             =   8040
      Width           =   13410
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "WHOLESALE MAXIMUM ORDER QUANTITY"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1335
         TabIndex        =   49
         Tag             =   "XYF"
         Top             =   600
         Width           =   3330
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WSmQ ="
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   525
         TabIndex        =   48
         Tag             =   "XYF"
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO = "
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   11130
         TabIndex        =   20
         Tag             =   "XYF"
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AVQ = "
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   6435
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RO = "
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   3765
         TabIndex        =   18
         Tag             =   "XYF"
         Top             =   240
         Width           =   420
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BO ="
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   9075
         TabIndex        =   17
         Tag             =   "XYF"
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RmQ ="
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   525
         TabIndex        =   16
         Tag             =   "XYF"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Purchase Order"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   11625
         TabIndex        =   15
         Tag             =   "XYF"
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "AVAILABLE QUANTITY"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6930
         TabIndex        =   14
         Tag             =   "XYF"
         Top             =   240
         Width           =   1725
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "RESERVED ORDER"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4260
         TabIndex        =   13
         Tag             =   "XYF"
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label6 
         Caption         =   "BACK-ORDER"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9570
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "RETAIL MAXIMUM ORDER QTY"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1095
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   240
         Width           =   2400
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msflxHead 
      Height          =   300
      Left            =   240
      TabIndex        =   21
      Tag             =   "XYF"
      Top             =   9240
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   529
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   16777215
      FocusRect       =   2
      ScrollBars      =   0
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
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
   Begin VB.Frame fmHeader 
      Height          =   1245
      Left            =   240
      TabIndex        =   1
      Tag             =   "XYF"
      Top             =   480
      Width           =   13455
      Begin VB.TextBox txtOrderDate 
         Height          =   315
         Left            =   7800
         TabIndex        =   45
         Tag             =   "XYF"
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtReference 
         Height          =   315
         Left            =   7800
         TabIndex        =   41
         Tag             =   "XYF"
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtStatus 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   11040
         Locked          =   -1  'True
         TabIndex        =   40
         Tag             =   "XYF"
         Top             =   360
         Width           =   1875
      End
      Begin VB.TextBox txtbranch 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   1695
         Locked          =   -1  'True
         TabIndex        =   25
         Tag             =   "XYF"
         Top             =   375
         Width           =   3900
      End
      Begin VB.ComboBox cboMos 
         Height          =   315
         Left            =   4080
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   750
         Width           =   1500
      End
      Begin VB.Label lblDateConf 
         Alignment       =   1  'Right Justify
         Caption         =   "P.O. STATUS"
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
         Left            =   9660
         TabIndex        =   33
         Tag             =   "XYF"
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label lblDateConf 
         Alignment       =   1  'Right Justify
         Caption         =   "REFERENCE #"
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
         Left            =   6240
         TabIndex        =   28
         Tag             =   "XYF"
         Top             =   450
         Width           =   1455
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1680
         TabIndex        =   22
         Tag             =   "XYF"
         Top             =   720
         Width           =   2355
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "MONTH/S (SQ)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   210
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   780
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ORDER DATE"
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
         Left            =   5880
         TabIndex        =   3
         Tag             =   "XY"
         Top             =   810
         Width           =   1830
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "BRANCH NAME"
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
         Left            =   210
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   435
         Width           =   1335
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxAllPOM 
      Height          =   1875
      Left            =   5220
      TabIndex        =   35
      Top             =   9480
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   3307
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin MSFlexGridLib.MSFlexGrid flxSP 
      Height          =   4845
      Left            =   180
      TabIndex        =   0
      Tag             =   "XYF"
      Top             =   1800
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   8546
      _Version        =   393216
      Rows            =   3
      Cols            =   16
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   16777215
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
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
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   3675
      TabIndex        =   50
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
      Image           =   "BranchPO.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   105
      TabIndex        =   51
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "P&rocess P.O."
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
      Image           =   "BranchPO.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   1890
      TabIndex        =   52
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "View Su&mmary"
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
      Image           =   "BranchPO.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   5160
      TabIndex        =   53
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
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
      Image           =   "BranchPO.frx":10CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   6960
      TabIndex        =   54
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "Save"
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
      Image           =   "BranchPO.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   8760
      TabIndex        =   55
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "Cancel"
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
      Image           =   "BranchPO.frx":1C02
      cBack           =   -2147483633
   End
   Begin VB.Label Label19 
      Caption         =   "APPROVED BY :"
      Height          =   255
      Left            =   9720
      TabIndex        =   58
      Tag             =   "XYF"
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Label Label18 
      Caption         =   "REMARKS:"
      Height          =   255
      Left            =   0
      TabIndex        =   57
      Tag             =   "XYF"
      Top             =   0
      Width           =   1515
   End
   Begin VB.Label Label3 
      Caption         =   "TOTAL "
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
      TabIndex        =   34
      Tag             =   "XYF"
      Top             =   6900
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label20 
      Caption         =   "REMARKS:"
      Height          =   255
      Left            =   300
      TabIndex        =   31
      Tag             =   "XYF"
      Top             =   7290
      Width           =   1515
   End
   Begin VB.Label Label12 
      Caption         =   "PREPARED BY:"
      Height          =   255
      Left            =   300
      TabIndex        =   29
      Tag             =   "XYF"
      Top             =   7680
      Width           =   1515
   End
   Begin VB.Label lblSQSERVE 
      Height          =   330
      Left            =   3300
      TabIndex        =   27
      Top             =   9360
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblBID 
      Height          =   330
      Left            =   2685
      TabIndex        =   26
      Top             =   9360
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblChk 
      Caption         =   "ü"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2250
      TabIndex        =   24
      Top             =   9360
      Visible         =   0   'False
      Width           =   255
   End
   Begin ResizeXtra.RXResizer RXResizePos 
      Left            =   12060
      Top             =   660
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14115
      m_MDI_Height    =   9585
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
      Left            =   480
      TabIndex        =   6
      Tag             =   "XYF"
      Top             =   120
      Width           =   2610
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      Height          =   255
      Index           =   0
      Left            =   8700
      TabIndex        =   5
      Tag             =   "XYF"
      Top             =   6900
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu mnuPrint 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuPickBPO 
         Caption         =   "PICK BRANCH P.O."
      End
      Begin VB.Menu mnuFinalBPO 
         Caption         =   "FINAL BRANCH P.O."
      End
   End
   Begin VB.Menu mnuOrder 
      Caption         =   "Order Taking"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete Row"
      End
   End
   Begin VB.Menu mnuconfirm 
      Caption         =   ""
      Visible         =   0   'False
   End
End
Attribute VB_Name = "BranchPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AddNew As Boolean                  'Boolean expression for Adding new records
Dim BtnIndx As Integer                 'Index for top Lacation of Button
Dim rsBrPO As New ADODB.Recordset
Dim pmos, myLastUpdt, Qry As String
Public mtrBName As String
Public isEdit As Boolean
Dim ActRow As Integer, ActCol As Integer, myOrderIDSQ As Long
Dim myOrderIDCof As Long
Dim refnumber As String
Dim bORDERID As Integer
Dim LstUpdt As String, myBranchCategory As String


Sub ClearEntries()

    flxSP.Rows = 1
    flxSP.Rows = 2
    txtbranch.Text = ""
    txtReference.Text = ""
    txtStatus.Text = ""
    txtBranchSPOID.Text = ""
    txtBranchID.Text = ""
    txtSPOID.Text = ""
    txtPrnCopy.Text = ""
    txtAllocQty.Text = ""
    cboMos.Text = ""
    txtTotalAmount.Text = ""
    txtRemarks.Text = ""
    txtPreparedby.Text = ""
    txtOrderDate.Text = ""
    txtOrderType.Text = ""
End Sub

Sub LockFields()
    
    txtbranch.Locked = True
    txtReference.Locked = True
    txtStatus.Locked = True
    cboMos.Locked = True
    txtPreparedby.Locked = True
    txtOrderDate.Locked = True
    
End Sub

Sub SaveBranchOrder()
    Dim pRW As Long, TotOrder As Long, TotBackOrder As Long, NoServe As Long, POTranstype As String
    Dim pos As Integer, TOTALQtyTK As Long, TotAmount As Currency, LastUpdtTk As String
    
        If flxSP.TextMatrix(1, 1) = "" Then
            MsgBox ("No Records were going to Update"), vbExclamation, "Empty"
            Exit Sub
        End If
             
             If AddNew Then
                'Qry = "SELECT MAX(spordertaking.SQSERVE) As SQLASTORDER From spordertaking Where spordertaking.BRANCHID= " & Val(lblBID.Caption) & " AND spordertaking.CONFIRMNO <> 'CANCELLED';"
                Qry = "SELECT MAX(spordertaking.SQSERVE) As SQLASTORDER From spordertaking;"
               
                NoServe = Val(GetFieldValue(Qry)) + 1
           
                LastUpdtTk = Format(Now, "yyyy-mm-dd hh:mm:ss")
                LYRcon.Execute "INSERT INTO spordertaking(BRANCHID, BRANCHNAME, MOS, DATESERVE, ORDERTYPE, QTYTAKING, BRANCHBO, TOTALAMOUNT, SQSERVE, REMARKS, PREPAREDBY, APPROVEDBY, LASTUPDATE, MYPC)VALUES(" & Val(lblBID.Caption) & ", " & _
                    "'" & txtbranch.Text & "', " & Val(cboMos) & ", '" & Format(txtDateServe.Text, "yyyy-mm-dd") & "', '" & cbodrtype.Text & "', " & TOTALQtyTK & ", " & TotBackOrder & ", " & TotAmount & ", " & NoServe & ", '" & txtRemarks.Text & "', " & _
                    "'" & txtPreparedby.Text & "', '" & txtApprovedBy.Text & "', '" & LastUpdtTk & "', '" & ComputerName & "');"
                
                Qry = "SELECT spordertaking.ORDERIDSQ From spordertaking Where spordertaking.SQSERVE = " & NoServe & ";"
                ' AND spordertaking.MYPC = '" & ComputerName & "';"
                myOrderIDSQ = GetFieldValue(Qry)
                
                Qry = "SELECT MAX(spordrconfirm.ORDERID) AS LASTORDERID From spordrconfirm;"
                myOrderIDCof = Val(GetFieldValue(Qry)) + 1
           
                lblSQSERVE.Caption = NoServe
                POTranstype = "ORDER TEMP"
                
                For pRW = 1 To flxSP.Rows - 1
                    LYRcon.Execute "INSERT INTO spbrchconfirmdtl(ORDERID, SPAREPARTID, BRANCHID, " & _
                        "TRANSTYPE, MQTYONHAND, MBACKORDER, MQTYSALES, BQTYONHAND, BBACKORDER, BPORDER, ORDERTAKING, PRICE, REFERENCE, " & _
                        "UPDATEDBY)VALUES(" & myOrderIDCof & ", " & flxSP.TextMatrix(pRW, 14) & ", " & Val(lblBID.Caption) & ", '" & POTranstype & "', " & _
                        "" & Val(flxSP.TextMatrix(pRW, 3)) & ", " & Val(flxSP.TextMatrix(pRW, 5)) & ", " & Val(flxSP.TextMatrix(pRW, 6)) & ", " & _
                        "" & Val(flxSP.TextMatrix(pRW, 7)) & ", " & Val(flxSP.TextMatrix(pRW, 8)) & ", " & Val(flxSP.TextMatrix(pRW, 9)) & ", " & _
                        "" & Val(flxSP.TextMatrix(pRW, 10)) & ", " & Val(Replace(flxSP.TextMatrix(pRW, 11), ",", "")) & ", " & _
                        "'" & flxSP.TextMatrix(pRW, 13) & "', " & UserID & ");"
                Next pRW
                
                For pos = 1 To flxAllPOM.Rows - 1
                      
                    Qry = "SELECT SUM(spbrchconfirmdtl.BPORDER) As TOTALORDER From spbrchconfirmdtl " & _
                        " where spbrchconfirmdtl.REFERENCE = '" & flxAllPOM.TextMatrix(pos, 0) & "' AND  spbrchconfirmdtl.BRANCHID= " & Val(lblBID.Caption) & " AND spbrchconfirmdtl.ORDERID='" & myOrderIDCof & "';"
                    TotOrder = Val(GetFieldValue(Qry))
    
                    Qry = "SELECT SUM(spbrchconfirmdtl.BBACKORDER) As TOTALBACKORDER From spbrchconfirmdtl " & _
                        " where spbrchconfirmdtl.REFERENCE = '" & flxAllPOM.TextMatrix(pos, 0) & "' AND  spbrchconfirmdtl.BRANCHID= " & Val(lblBID.Caption) & " AND spbrchconfirmdtl.ORDERID='" & myOrderIDCof & "';"
                    TotBackOrder = Val(GetFieldValue(Qry))
                    
                    Qry = "SELECT SUM(spbrchconfirmdtl.ORDERTAKING) As TOTALORDER From spbrchconfirmdtl " & _
                        " where spbrchconfirmdtl.REFERENCE = '" & flxAllPOM.TextMatrix(pos, 0) & "' AND  spbrchconfirmdtl.BRANCHID= " & Val(lblBID.Caption) & " AND spbrchconfirmdtl.ORDERID='" & myOrderIDCof & "';"
                    TOTALQtyTK = Val(GetFieldValue(Qry))
                    
                    Qry = "SELECT SUM(spbrchconfirmdtl.PRICE) As TOTALORDER From spbrchconfirmdtl " & _
                        " where spbrchconfirmdtl.REFERENCE = '" & flxAllPOM.TextMatrix(pos, 0) & "' AND  spbrchconfirmdtl.BRANCHID= " & Val(lblBID.Caption) & " AND spbrchconfirmdtl.ORDERID='" & myOrderIDCof & "';"
                    TotAmount = Val(GetFieldValue(Qry))
                    
                    myLastUpdt = Format(Now, "yyyy-mm-dd hh:mm:ss")
                    LYRcon.Execute "INSERT INTO spordrconfirm(ORDERID, BRANCHID, BRANCHNAME, REFERENCE, ORDERTYPE, QTYORDER, BACKORDER, TOTALAMOUNTPO, SQSERVE, LASTUPDATE, MYPC)VALUES(" & myOrderIDCof & ", " & Val(lblBID.Caption) & ", " & _
                        "'" & txtbranch.Text & "', '" & flxAllPOM.TextMatrix(pos, 0) & "', '" & cbodrtype.Text & "', " & TotOrder & ", " & TotBackOrder & ", " & Replace(txtTotalAmount.Text, ",", "") & ", " & _
                        "" & NoServe & ", '" & myLastUpdt & "', '" & ComputerName & "');"
                
                Next pos
                
                MsgBox ("Purchase Order Successfully Save"), vbInformation
            Else
                    
                    LastUpdtTk = Format(Now, "yyyy-mm-dd hh:mm:ss")
                   
                    LYRcon.Execute "UPDATE spordertaking SET BRANCHID = " & Val(lblBID.Caption) & ", BRANCHNAME = '" & txtbranch.Text & "', MOS = " & Val(cboMos) & ", " & _
                        "DATESERVE = '" & Format(txtDateServe.Text, "yyyy-mm-dd") & "', ORDERTYPE = '" & cbodrtype.Text & "', QTYTAKING = " & TOTALQtyTK & ", " & _
                        "BRANCHBO = " & TotBackOrder & ", TOTALAMOUNT = " & TotAmount & ", SQSERVE = " & Val(lblSQSERVE.Caption) & ", PREPAREDBY = '" & txtPreparedby.Text & "', " & _
                        "REMARKS = '" & txtRemarks.Text & "', APPROVEDBY = '" & txtApprovedBy.Text & "', LASTUPDATE = '" & LastUpdtTk & "' Where spordertaking.ORDERIDSQ = " & myOrderIDSQ & ";"
                        'lblSQSERVE.Caption = NoServe
                        POTranstype = "ORDER TEMP"
                    
                    
                    LYRcon.Execute "DELETE from spbrchconfirmdtl " & _
                        "Where spbrchconfirmdtl.ORDERID = " & myOrderIDCof & " AND spbrchconfirmdtl.BRANCHID = " & Val(lblBID.Caption) & ";"
                    For pRW = 1 To flxSP.Rows - 1
                     
                     LYRcon.Execute "INSERT INTO spbrchconfirmdtl(ORDERID, SPAREPARTID, BRANCHID, " & _
                        "TRANSTYPE, MQTYONHAND, MBACKORDER, MQTYSALES, BQTYONHAND, BBACKORDER, BPORDER, ORDERTAKING, PRICE, REFERENCE, " & _
                        "UPDATEDBY)VALUES(" & myOrderIDCof & ", " & flxSP.TextMatrix(pRW, 14) & ", " & Val(lblBID.Caption) & ", '" & POTranstype & "', " & _
                        "" & Val(flxSP.TextMatrix(pRW, 3)) & ", " & Val(flxSP.TextMatrix(pRW, 5)) & ", " & Val(flxSP.TextMatrix(pRW, 6)) & ", " & _
                        "" & Val(flxSP.TextMatrix(pRW, 7)) & ", " & Val(flxSP.TextMatrix(pRW, 8)) & ", " & Val(flxSP.TextMatrix(pRW, 9)) & ", " & _
                        "" & Val(flxSP.TextMatrix(pRW, 10)) & ", " & Val(Replace(flxSP.TextMatrix(pRW, 11), ",", "")) & ", " & _
                        "'" & flxSP.TextMatrix(pRW, 13) & "', " & UserID & ");"
                     Next pRW
                        
                     For pos = 1 To flxAllPOM.Rows - 1
                        
                        Qry = "SELECT SUM(spbrchconfirmdtl.BPORDER) As TOTALORDER From spbrchconfirmdtl " & _
                        " where spbrchconfirmdtl.REFERENCE = '" & flxAllPOM.TextMatrix(pos, 0) & "' AND  spbrchconfirmdtl.BRANCHID= " & Val(lblBID.Caption) & " AND spbrchconfirmdtl.ORDERID=" & myOrderIDCof & ";"
                        TotOrder = Val(GetFieldValue(Qry))
    
                        Qry = "SELECT SUM(spbrchconfirmdtl.BBACKORDER) As TOTALBACKORDER From spbrchconfirmdtl " & _
                            " where spbrchconfirmdtl.REFERENCE = '" & flxAllPOM.TextMatrix(pos, 0) & "' AND  spbrchconfirmdtl.BRANCHID= " & Val(lblBID.Caption) & " AND spbrchconfirmdtl.ORDERID=" & myOrderIDCof & ";"
                        TotBackOrder = Val(GetFieldValue(Qry))
                        
                        Qry = "SELECT SUM(spbrchconfirmdtl.ORDERTAKING) As TOTALORDER From spbrchconfirmdtl " & _
                            " where spbrchconfirmdtl.REFERENCE = '" & flxAllPOM.TextMatrix(pos, 0) & "' AND  spbrchconfirmdtl.BRANCHID= " & Val(lblBID.Caption) & " AND spbrchconfirmdtl.ORDERID=" & myOrderIDCof & ";"
                        TOTALQtyTK = Val(GetFieldValue(Qry))
                        
                        Qry = "SELECT SUM(spbrchconfirmdtl.PRICE) As TOTALORDER From spbrchconfirmdtl " & _
                            " where spbrchconfirmdtl.REFERENCE = '" & flxAllPOM.TextMatrix(pos, 0) & "' AND  spbrchconfirmdtl.BRANCHID= " & Val(lblBID.Caption) & " AND spbrchconfirmdtl.ORDERID=" & myOrderIDCof & ";"
                        TotAmount = Val(GetFieldValue(Qry))
     
                        myLastUpdt = Format(Now, "yyyy-mm-dd hh:mm:ss")
                        LYRcon.Execute "Update spordrconfirm SET BRANCHID = " & Val(lblBID.Caption) & ", BRANCHNAME = '" & txtbranch.Text & "', " & _
                            "ORDERTYPE = '" & cbodrtype.Text & "', QTYORDER = " & TotOrder & ", BACKORDER = " & TotBackOrder & ", TOTALAMOUNTPO = " & TotAmount & ", " & _
                            "SQSERVE = " & Val(lblSQSERVE.Caption) & ", LASTUPDATE = '" & myLastUpdt & "', MYPC = '" & ComputerName & "' " & _
                            "Where spordrconfirm.SQSERVE = " & myOrderIDSQ & " AND spordrconfirm.REFERENCE = '" & flxAllPOM.TextMatrix(pos, 0) & "' AND spordrconfirm.BRANCHID = " & Val(lblBID.Caption) & ";"
                    
                    Next pos
                MsgBox ("Purchase Order Successfully Updated"), vbInformation
             End If
End Sub

Private Sub BtnSaveCancel(Bln As Boolean)
    Dim i As Integer
    
    If Bln Then
        For i = 1 To 4
            cmdTrans(i).Visible = False
        Next i
        cmdTrans(5).Left = 0
        cmdTrans(6).Left = cmdTrans(5).Width
        cmdTrans(5).Visible = True
        cmdTrans(6).Visible = True
        
    Else
        For i = 1 To 4
            cmdTrans(i).Visible = True
        Next i
        cmdTrans(5).Left = 0
        cmdTrans(6).Left = cmdTrans(5).Width
        cmdTrans(5).Visible = False
        cmdTrans(6).Visible = False
        
    End If
End Sub


Private Sub cmdTrans_Click(Index As Integer)
  Dim rsTrans As New ADODB.Recordset
  Dim allocNum As Integer, printCount As Integer
    
    Select Case Index
    
        Case 0 'PROCESS PO
            If Val(txtAllocQty.Text) > 0 Then
                MsgBox "P.O. Already Process!", vbExclamation + vbOKOnly, "Edit"
                Exit Sub
            End If
            If Val(txtSPOID.Text) = 0 Then
                 MsgBox "No Record to be process!", vbExclamation + vbOKOnly, "Edit"
                Exit Sub
            End If
            flxSP.Enabled = True
            fmHeader.Enabled = True
            Call BtnSaveCancel(True)
            txtDateServe.Text = Date
         
         Case 1 'PROCESS PO or EDIT
            Call LockFields
            If Val(txtPrnCopy.Text) > 0 And (txtStatus.Text = "PICKING" Or txtStatus.Text = "INTRANSIT" Or txtStatus.Text = "PACKING") Then
                MsgBox "This is not allowed because the Branch P.O. is already Final!", vbExclamation + vbOKOnly, "Edit"
                Exit Sub
            End If
            If Val(txtSPOID.Text) = 0 Then
                 MsgBox "No Record Found!", vbExclamation + vbOKOnly, "Edit"
                Exit Sub
            End If
            
            ans = MsgBox("Do you want to proceed this action?", vbQuestion + vbYesNo, "Edit")
            If ans = vbNo Then Exit Sub
            Qry = "SELECT concat(employees.FIRSTNAME,' ',employees.LASTNAME) as PrepBy From employees Where employees.EmpId = " & UserID & ";"
            txtPreparedby.Text = GetFieldValue(Qry)
            fmHeader.Enabled = True
            flxSP.Enabled = True
            AddNew = True
            Call BtnSaveCancel(True)
            Call ComputeTotalAllocation
        
        
        Case 2 'SUMMARY
            FrmSummary.mySrc = "Branch Purchase Order"
            FrmSummary.Show
            Unload Me
        
        Case 3 'PRINT
            If txtStatus.Text = "PICKING" Or txtStatus.Text = "PACKING" Or txtStatus.Text = "INTRANSIT" Then
                mnuPickBPO.Enabled = False
                mnuFinalBPO.Enabled = True 'CHANGE TO FALSE
            ElseIf Val(txtPrnCopy.Text) = 0 Then
                mnuFinalBPO.Enabled = False
            End If
            Call PopupMenu(mnuPrint)
            
            
        Case 4 'EXIT
            Unload Me
            
        Case 5 'SAVE
            Call ComputeTotalAllocation
            ans = MsgBox("Do you want to Save Branch P.O.", vbQuestion + vbYesNo, "Save")
            If ans = vbNo Then Exit Sub
            LstUpdt = Format(Now, "yyyy-mm-dd hh:mm:ss")
            
            If AddNew Then
                LYRcon.Execute "Update sporderbranch SET PICKEDBY = '" & txtPreparedby.Text & "', REMARKS = '" & txtRemarks.Text & "', UPDATEDBY = " & UserID & ", ALLOCATED = " & txtTotalAlloc.Text & ", MOS = " & cboMos.Text & ", TOTALAMOUNTPO = '" & Format(txtTotalAmount.Text, "#,###,###.00") & "', LASTUPDATE = '" & LstUpdt & "' " & _
                                "Where sporderbranch.REFERENCE = '" & txtReference.Text & "' AND sporderbranch.BRANCHSPOID = " & txtBranchSPOID.Text & ";"
            End If
            With flxSP
                    For i = 2 To .Rows - 1
                        '.TextMatrix(i, 9) = IIf(.TextMatrix(i, 9) = "", 0, .TextMatrix(i, 9))
                        LYRcon.Execute "UPDATE spbranchdtl SET ALLOCATEDQTY = " & Val(.TextMatrix(i, 10)) & ", RESERVEDORDER = " & Val(.TextMatrix(i, 5)) & ", BACKORDERQTY = " & Val(.TextMatrix(i, 8)) & " Where spbranchdtl.REFERENCE = '" & .TextMatrix(i, 15) & "' and spbranchdtl.SPAREPARTID = " & .TextMatrix(i, 14) & ";"
                    Next i
            End With
            
            Call BtnSaveCancel(False)
            
        Case 6 'CANCEL
            ans = MsgBox("Do you want to Cancel the Process?", vbQuestion + vbYesNo, "Cancel")
            If ans = vbNo Then Exit Sub
            Call BtnSaveCancel(False)
            Call ClearEntries
        
            
            
    
    End Select
End Sub



Private Sub mnuPickBPO_click()
    'MsgBox "picking"
    Dim strLen As Integer
    Dim Qry As String
    Dim newString As String
    Qry = "SELECT BRANCHCODE FROM branch WHERE BRANCHNAME = '" & txtbranch.Text & "';"
    strLen = Len(GetFieldValue(Qry))
    newString = Left(txtReference.Text, Len(txtReference.Text) - strLen)
    
            If Val(txtSPOID.Text) = 0 Then
                MsgBox "No records found", vbExclamation + vbOKOnly, "Print"
                Exit Sub
            End If
            
            ans = MsgBox("Do you want to Print Branch PO?", vbQuestion + vbYesNo, "Print")
            If ans = vbNo Then Exit Sub
            
            frmPrintPO.MySource = "PICKBPO"
            frmPrintPO.QryPrn = "call sp_mrcwarehouse.sp_printPickPO ('" & txtbranch.Text & "'," & Val(txtBranchSPOID.Text) & ",'" & txtReference.Text & "')"
            frmPrintPO.ParamValue1 = SysBranch
            frmPrintPO.ParamValue2 = SysAddress
            frmPrintPO.ParamValue3 = SysTel
            frmPrintPO.ParamValue4 = "" 'txtbranch.Text
            frmPrintPO.ParamValue5 = newString
            frmPrintPO.ParamValue6 = txtOrderDate.Text
            frmPrintPO.ParamValue7 = txtPreparedby.Text
            frmPrintPO.ParamValue8 = cboAprBy.Text
            frmPrintPO.Show
            
End Sub

Private Sub mnuFinalBPO_click()
    'MsgBox "final"
            If Val(txtSPOID.Text) = 0 Then
                MsgBox "No records found", vbExclamation + vbOKOnly, "FINAL BRANCH P.O."
                Exit Sub
            End If
            
            ans = MsgBox("Do you want to Print Branch PO?", vbQuestion + vbYesNo, "Print")
            If ans = vbNo Then Exit Sub
            
            With flxSP
                    For i = 2 To .Rows - 1
                        If .TextMatrix(i, 8) > 0 Then 'if naay bo
                            LYRcon.Execute "INSERT INTO backorder(BRANCHSPOID, REFERENCE, TRANSDATE, SPAREPARTID, BACKORDERQTY, BOBALANCE, BRANCHNAME, BRANCHID) " & _
                                            "VALUES(" & Val(txtBranchSPOID.Text) & ", '" & .TextMatrix(i, 15) & "', '" & Format(txtOrderDate.Text, "yyyy-mm-dd") & "', " & Val(.TextMatrix(i, 14)) & ", " & Val(.TextMatrix(i, 8)) & ", " & Val(.TextMatrix(i, 8)) & ", '" & txtbranch.Text & "', '" & txtBranchID.Text & "');"
                        
                        End If
                    Next i
                
            End With
            
            frmPrintPO.MySource = "Branch Purchase Order"
            frmPrintPO.QryPrn = "call sp_mrcwarehouse.sp_printFinalPO('" & txtbranch.Text & "'," & Val(txtBranchSPOID.Text) & ",'" & txtReference.Text & "')"
            frmPrintPO.ParamValue1 = SysBranch
            frmPrintPO.ParamValue2 = SysAddress
            frmPrintPO.ParamValue3 = SysTel
            frmPrintPO.ParamValue4 = "" 'txtbranch.Text
            frmPrintPO.ParamValue5 = "" 'txtReference.Text
            frmPrintPO.ParamValue6 = txtOrderDate.Text
            frmPrintPO.ParamValue7 = txtPreparedby.Text
            frmPrintPO.Show
            
End Sub

Private Sub flxSP_Click()
    ActRow = flxSP.row
    ActCol = flxSP.col
    
    If ActCol <> 10 Then Exit Sub
    'If txtPrnCopy.Text > 0 Then Exit Sub
    
    With txtOrder
        .Top = flxSP.CellTop + flxSP.Top
        .Left = flxSP.CellLeft + flxSP.Left
        .Width = flxSP.CellWidth
        .Text = Trim(flxSP.Text)
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .ZOrder
        .SetFocus
    End With
    
End Sub

Private Sub flxSP_KeyPress(KeyAscii As Integer)
    Call flxSP_Click
End Sub


Private Sub flxSP_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuOrder
    End If
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxSP_Click
    fmHeader.Enabled = False
End Sub
Private Sub Form_Load()
    fmHeader.Enabled = False
    flxSP.Enabled = True
    Call ComputeTotalAllocation
    Call loadApprovedBy
    cboAprBy.ListIndex = 0
End Sub

Private Sub loadApprovedBy()
Dim rsApproval As New ADODB.Recordset
Dim Qry As String
Qry = "SELECT * FROM employees s WHERE s.`POSITION` = 'admin. officer' AND s.DEPARTMENT = 'admin' and username <> 'inactive' and inactive = 0;"
With rsApproval
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboAprBy.Clear
            Do While Not .EOF
                cboAprBy.AddItem !FIRSTNAME & " " & !MIDDLENAME & " " & !LASTNAME
                .MoveNext
            Loop
        End If
    End With
End Sub

Private Sub mnuconfirm_Click()
    Dim Qry As String
    Dim ans As Integer, pRW As Long, ref As Long
        
        If lblSTATUS.Caption = "" Then
            MsgBox ("No Records to Print"), vbExclamation, "No Record"
            Exit Sub
        End If
        Qry = "SELECT spordertaking.PRINTPICK from spordertaking Where spordertaking.BRANCHNAME = '" & txtbranch & "' AND spordertaking.SQSERVE = " & Val(lblSQSERVE.Caption) & ";"
        If GetFieldValue(Qry) = 0 Then
            ans = MsgBox("Please prepare Print-out for picking", vbQuestion, "Order Picking")
            Exit Sub
        End If
        Qry = "SELECT spordertaking.DATECONFIRM from spordertaking Where spordertaking.BRANCHNAME = '" & txtbranch & "' AND spordertaking.SQSERVE = " & Val(lblSQSERVE.Caption) & ";"
        If GetFieldValue(Qry) = Null Or GetFieldValue(Qry) = "" Then
            ans = MsgBox("Print Final Confirmation", vbQuestion + vbYesNo, "Order Picking")
            If ans = vbNo Then Exit Sub
        End If
           LYRcon.Execute "DELETE from spbrchconfirmdtl " & _
                            "Where spbrchconfirmdtl.ORDERID = " & myOrderIDCof & " AND spbrchconfirmdtl.BRANCHID = " & Val(lblBID.Caption) & " AND spbrchconfirmdtl.ORDERTAKING=0;"
            For ref = 1 To flxAllPOM.Rows - 1
                    Qry = "SELECT SUM(spbrchconfirmdtl.BBACKORDER) As TOTALBACKORDER From spbrchconfirmdtl " & _
                        " where spbrchconfirmdtl.REFERENCE = '" & flxAllPOM.TextMatrix(ref, 0) & "' AND  spbrchconfirmdtl.BRANCHID= " & Val(lblBID.Caption) & " AND spbrchconfirmdtl.ORDERID='" & myOrderIDCof & "';"
                    If Val(GetFieldValue(Qry)) = 0 Then
                        LYRcon.Execute "Update sporderbranch SET ALLOCATED = 1 Where sporderbranch.REFERENCE =  '" & flxAllPOM.TextMatrix(ref, 0) & "' AND sporderbranch.BRANCHID = " & Val(lblBID.Caption) & ";"
                    End If
            Next ref
        
        DoEvents
        
        PRNmc.MySource = "SPBPOACONFIRM"
        PRNmc.QryPrn = "Select DISTINCT spbrchconfirmdtl.ORDERTAKING, spareparts.PARTNUMBER, spareparts.DESCRIPTION, spbrchconfirmdtl.REFERENCE, spareparts.COST, spbrchconfirmdtl.ORDERTAKING * spareparts.COST AS TOTALAMOUNT " & _
            "From spbrchconfirmdtl INNER JOIN spareparts " & _
                "ON (spbrchconfirmdtl.SPAREPARTID = spareparts.SPAREPARTID) INNER JOIN spordrconfirm " & _
                "ON (spbrchconfirmdtl.ORDERID = spordrconfirm.ORDERID) " & _
            "Where spbrchconfirmdtl.BRANCHID =  " & Val(lblBID.Caption) & " AND spordrconfirm.SQSERVE = " & Val(lblSQSERVE.Caption) & " ORDER BY spareparts.PARTNUMBER;"
        
        PRNmc.ParamValue1 = SysBranch
        PRNmc.ParamValue2 = SysAddress
        PRNmc.ParamValue3 = SysTel
        PRNmc.ParamValue4 = txtbranch.Text
        Qry = "SELECT branch.ADDRESS From branch Where branch.BRANCHID = " & Val(lblBID) & ";"
        PRNmc.ParamValue5 = GetFieldValue(Qry)
        PRNmc.ParamValue6 = txtDateServe.Text
        PRNmc.ParamValue7 = lblSQSERVE.Caption
        PRNmc.ParamValue8 = txtPreparedby.Text
        PRNmc.ParamValue9 = CStr(txtDateCon.Text)
        PRNmc.ParamValue10 = "JAMIE CARLO N. ANG"
        
        PRNmc.Show
 End Sub

Private Sub mnuDelete_Click()
Dim ans As Integer
    ans = MsgBox("Delete Current Row record?", vbYesNo + vbQuestion, "Delete Row")
    If ans = vbYes Then
        If flxSP.Rows = 2 Then
            flxSP.Rows = 1
            flxSP.Rows = 2
        Else
            flxSP.RemoveItem (flxSP.row)
        End If
    End If
End Sub
Private Sub mnupick_Click()
    Dim Qry As String
    Dim ans As Integer
        
        If lblSTATUS.Caption = "" Then
            MsgBox ("View Records to Print"), vbExclamation, "No Record"
            Exit Sub
        End If
        'Qry = "SELECT spordrconfirm.PRINTPICK from spordrconfirm Where spordrconfirm.BRANCHNAME = '" & txtbranch & "' AND  spordrconfirm.REFERENCE = '" & txtPONumber.Text & "' AND  spordrconfirm.SQSERVE = " & Val(lblSQSERVE.Caption) & ";"
        Qry = "SELECT spordertaking.PRINTPICK from spordertaking Where spordertaking.BRANCHNAME = '" & txtbranch & "' AND  spordertaking.SQSERVE = " & Val(lblSQSERVE.Caption) & ";"
         
        If GetFieldValue(Qry) = 0 Then
            ans = MsgBox("Do you want to Print P.O for Order Picking", vbQuestion + vbYesNo, "Order Picking")
            If ans = vbNo Then Exit Sub
            mnuconfirm.Enabled = True
        End If
        'One Reference Id
        
        PRNmc.MySource = "SPBPO"
        If flxAllPOM.Rows = 2 Then
            PRNmc.QryPrn = "SELECT DISTINCT spareparts.PARTNUMBER, spareparts.DESCRIPTION, spareparts.QTYONHAND, spareparts.LOCATION, spbrchconfirmdtl.REFERENCE, spbrchconfirmdtl.BBACKORDER " & _
                " From spbrchconfirmdtl " & _
                " INNER JOIN spareparts ON (spbrchconfirmdtl.SPAREPARTID = spareparts.SPAREPARTID) " & _
                " INNER JOIN spordertaking ON (spbrchconfirmdtl.BRANCHID = spordertaking.BRANCHID) " & _
                " AND (spbrchconfirmdtl.ORDERID = spordertaking.SQSERVE) " & _
                " Where spbrchconfirmdtl.BRANCHID = " & Val(lblBID.Caption) & "  AND " & _
                " spordertaking.ORDERIDSQ = " & myOrderIDSQ & " AND spbrchconfirmdtl.BBACKORDER > 0 Order By spareparts.PARTNUMBER;"
        Else
            PRNmc.QryPrn = "SELECT DISTINCT spareparts.PARTNUMBER, spareparts.DESCRIPTION, spareparts.QTYONHAND, spareparts.LOCATION, " & _
                " spbrchconfirmdtl.REFERENCE, spbrchconfirmdtl.BBACKORDER " & _
                "From spbrchconfirmdtl " & _
                "INNER JOIN spareparts ON (spbrchconfirmdtl.SPAREPARTID = spareparts.SPAREPARTID) " & _
                "INNER JOIN spordertaking ON (spbrchconfirmdtl.BRANCHID = spordertaking.BRANCHID) " & _
                "AND (spbrchconfirmdtl.ORDERID = spordertaking.SQSERVE) " & _
                "Where spbrchconfirmdtl.BRANCHID = " & Val(lblBID.Caption) & " AND spbrchconfirmdtl.BBACKORDER > 0 Order By SPAREPARTS.PartNumber;"
        
        End If
        'PRNmc.MySource = "SPBPO"
        'PRNmc.QryPrn = "SELECT DISTINCT spareparts.PARTNUMBER, spareparts.DESCRIPTION, spareparts.QTYONHAND, spareparts.LOCATION, spbrchconfirmdtl.REFERENCE, spbrchconfirmdtl.BBACKORDER " & _
        '    "From spbrchconfirmdtl " & _
        '    "INNER JOIN spareparts ON (spbrchconfirmdtl.SPAREPARTID = spareparts.SPAREPARTID) " & _
        '    "INNER JOIN spordertaking ON (spbrchconfirmdtl.BRANCHID = spordertaking.BRANCHID) " & _
        '    "Where spbrchconfirmdtl.BRANCHID = " & Val(lblBID.Caption) & " AND spordertaking.ORDERIDSQ = " & myOrderIDSQ & " AND spbrchconfirmdtl.BBACKORDER>0 ORDER BY spareparts.PARTNUMBER;"
        
        'PRNmc.QryPrn = "SELECT DISTINCT spareparts.PARTNUMBER, spareparts.DESCRIPTION, spareparts.QTYONHAND, spareparts.LOCATION, spbrchconfirmdtl.REFERENCE, " & _
            "spbrchconfirmdtl.BBACKORDER From spbrchconfirmdtl " & _
            "INNER JOIN spareparts ON (spbrchconfirmdtl.SPAREPARTID = spareparts.SPAREPARTID) " & _
            "INNER JOIN spordertaking ON (spbrchconfirmdtl.BRANCHID = spordertaking.BRANCHID) " & _
            "AND (spbrchconfirmdtl.ORDERID = spordertaking.ORDERIDSQ) " & _
            "Where spbrchconfirmdtl.BRANCHID = " & Val(lblBID.Caption) & " AND spordertaking.ORDERIDSQ = " & myOrderIDSQ & ";"
       
        PRNmc.ParamValue1 = SysBranch
        PRNmc.ParamValue2 = SysAddress
        PRNmc.ParamValue3 = SysTel
        PRNmc.ParamValue4 = txtbranch.Text
        Qry = "SELECT branch.ADDRESS From branch Where branch.BRANCHID = '" & Val(lblBID) & "';"
        PRNmc.ParamValue5 = GetFieldValue(Qry)
        PRNmc.ParamValue6 = txtDateServe.Text
        PRNmc.ParamValue7 = lblSQSERVE.Caption
        PRNmc.ParamValue8 = txtPreparedby.Text
        PRNmc.Show
End Sub




Private Sub RXResizePos_ResizeComplete()
    Call ResizeCompleteHead
End Sub



Private Sub tmrDrop_Timer() 'Timer for drop position of buton
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 8 Then
            tmrDrop.Enabled = False
        End If
    End If
End Sub
Private Sub tmrDropRev_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 6 Then
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
            BtnIndx = 6 'Index of Button Save
            cmdTrans(6).Top = -375
            cmdTrans(6).Left = 0
            cmdTrans(6).Visible = True
            cmdTrans(7).Top = -375
            cmdTrans(7).Left = cmdTrans(4).Width
            cmdTrans(7).Visible = True
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
            For i = 0 To 5
                cmdTrans(i).Visible = True
            Next i
            tmrDropRev.Enabled = True
        End If
    End If
End Sub
Private Sub ResizeCompleteHead()
Dim AdCol, aCol As Integer, aCel As Integer, reAln As Integer, BrWidth As Long
    
    With flxSP
        For aCol = 0 To 14
           .ColAlignment(aCol) = 4
        Next aCol
        'msflxHead.TextMatrix(0, 0) = "S P A R E P A R T S"
       ' msflxHead.TextMatrix(0, 1) = "WAREHOUSE"
        'msflxHead.TextMatrix(0, 2) = "BRANCH"
        .TextMatrix(0, 0) = "S P A R E P A R T S"
        .TextMatrix(0, 1) = "S P A R E P A R T S"
        .TextMatrix(0, 2) = "S P A R E P A R T S"
        .TextMatrix(0, 3) = "W A R E H O U S E"
        .TextMatrix(0, 4) = "W A R E H O U S E"
        .TextMatrix(0, 5) = "W A R E H O U S E"
        .TextMatrix(0, 6) = "W A R E H O U S E"
        .TextMatrix(0, 7) = txtbranch.Text
        .TextMatrix(0, 8) = txtbranch.Text
        .TextMatrix(0, 9) = txtbranch.Text
        .TextMatrix(0, 10) = txtbranch.Text
        .TextMatrix(0, 11) = txtbranch.Text
        .TextMatrix(0, 12) = txtbranch.Text
        .TextMatrix(0, 13) = txtbranch.Text
        
        .TextMatrix(1, 1) = "PARTNUMBER"
        .TextMatrix(1, 2) = "DESCRIPTION"
        .TextMatrix(1, 3) = "RmQ"
        .TextMatrix(1, 4) = "WSmQ"
        .TextMatrix(1, 5) = "w/STF"
        .TextMatrix(1, 6) = "AVQ"
        .TextMatrix(1, 7) = "QH"
        .TextMatrix(1, 8) = "BO"
        .TextMatrix(1, 9) = "OQ"
        .TextMatrix(1, 10) = "ALQ"
        .TextMatrix(1, 11) = "PRICE"
        .TextMatrix(1, 12) = "TOTAL"
        .TextMatrix(1, 13) = "LOCATION"
                
        
        .ColWidth(0) = 400
        .ColWidth(1) = (.Width - 400) * 0.13
        .ColWidth(2) = (.Width - 400) * 0.29
        .ColWidth(3) = (.Width - 400) * 0.05
        .ColWidth(4) = (.Width - 400) * 0.05
        .ColWidth(5) = (.Width - 400) * 0.05
        .ColWidth(6) = (.Width - 400) * 0.05
        .ColWidth(7) = (.Width - 400) * 0.05
        .ColWidth(8) = (.Width - 400) * 0.05
        .ColWidth(9) = (.Width - 400) * 0.05
        .ColWidth(10) = (.Width - 400) * 0.05
        .ColWidth(11) = (.Width - 400) * 0#
        .ColWidth(12) = (.Width - 400) * 0#
        .ColWidth(13) = (.Width - 400) * 0.11
        .ColWidth(14) = 0
        .ColWidth(15) = 0
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        msflxHead.ColWidth(0) = .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + 10
        msflxHead.ColWidth(1) = .ColWidth(3) * 3
       
        For AdCol = 6 To 11
            BrWidth = BrWidth + .ColWidth(AdCol)
        Next AdCol
        msflxHead.ColWidth(2) = BrWidth + .ColWidth(13)
        fmHeader.Width = flxSP.Width
        txtTotalAmount.Width = flxSP.ColWidth(12)
        txtTotalAmount.Left = flxSP.Width - flxSP.ColWidth(12) - flxSP.ColWidth(13)
        Label3.Left = txtTotalAmount.Left - txtTotalAmount.Width
        
        
    End With
End Sub
Sub ShowRecord(BId As Integer)  'P.O from Branch
    Dim rsPdn As New ADODB.Recordset
    Dim Qry As String, TOtPPO As Long, BOrderBal As Long, nR As Long
    Dim nRow As Long, fl As Integer, myW As String, myW1 As String, ASQSERVE As Long
    Dim r As Integer, myREFNO As String
    
    If BranchPOSum.chkPO.Value = 1 Then
        
        flxAllPOM.Clear
        flxAllPOM.Rows = 1
        BranchPOSum.flxAllPO.FixedRows = 1
        
        For r = 1 To BranchPOSum.flxAllPO.Rows - 1
             
            Qry = "SELECT * From spordertaking INNER JOIN spordrconfirm ON (spordertaking.SQSERVE = spordrconfirm.ORDERID) " & _
                  " Where spordrconfirm.REFERENCE = '" & BranchPOSum.flxAllPO.TextMatrix(r, 0) & "' AND spordertaking.FINALQT = 0;"
            If rsPdn.State = adStateOpen Then rsPdn.Close
            rsPdn.Open Qry, LYRcon, adOpenKeyset, adLockPessimistic
            If rsPdn.RecordCount > 0 Then
                PdnPO = PdnPO & BranchPOSum.flxAllPO.TextMatrix(r, 0) & " "
            End If
            Qry = "SELECT sporderbranch.ALLOCATED From sporderbranch Where sporderbranch.REFERENCE = '" & BranchPOSum.flxAllPO.TextMatrix(r, 0) & "' AND sporderbranch.BRANCHID = " & BId & ";"
            If Val(GetFieldValue(Qry)) = 0 Then
                flxAllPOM.AddItem BranchPOSum.flxAllPO.TextMatrix(r, 0)
            End If
        Next r
        
        If PdnPO <> "" Then
            MsgBox ("Reference No: " & PdnPO & " are Already Added for Order Taking..Please Print for Final Confirmation"), vbExclamation, "PENDING"
            flxAllPOM.Clear
            flxAllPOM.Rows = 2
            flxAllPOM.FixedRows = 1
            Exit Sub
        End If
        flxAllPOM.FixedRows = 1
    Else
         myREFNO = BranchPOSum.cboRefNo.Text
         flxAllPOM.Clear
         flxAllPOM.Rows = 1
         
         flxAllPOM.AddItem myREFNO
         flxAllPOM.FixedRows = 1
    End If
    cmdLoadPO.Enabled = False
   
    myW1 = ""
    If BranchPOSum.chkPO.Value = 0 Then
        myW1 = "AND sporderbranch.REFERENCE = '" & myREFNO & "'"
    Else
        For i = 1 To flxAllPOM.Rows - 1
            If i = 1 Then
                myW1 = myW1 & " AND sporderbranch.REFERENCE = '" & flxAllPOM.TextMatrix(i, 0) & "'"
            Else
                myW1 = myW1 & " OR sporderbranch.REFERENCE = '" & flxAllPOM.TextMatrix(i, 0) & "'"
            End If
        Next i
    End If
    myW = myW1
    
    Qry = "SELECT sporderbranch.REFERENCE, sporderbranch.BRANCHNAME, sporderbranch.BRANCHID, sporderbranch.ORDERTYPE, sporderbranch.ORDERDATE " & _
          "From sporderbranch Where sporderbranch.BRANCHID = " & BId & " " & myW & " ORDER BY sporderbranch.REFERENCE;"
    
    With rsBrPO
         If .State = adStateOpen Then .Close
         .Open Qry, LYRcon, adOpenKeyset, adLockPessimistic
             Call ClearEntries
             If .RecordCount = 0 Then Exit Sub: .Close
                 
                 lblBID.Caption = rsBrPO!branchID
                 txtbranch.Text = rsBrPO!branchName
                 'txtPONumber.Text = rsBrPO!REFERENCE
                 'txtOdrtype.Text = rsBrPO!ORDERTYPE
                 'txtPODate.Text = rsBrPO!ORDERDATE
                 lblSTATUS.Caption = "PENDING"
                 .Close
                myW = ""
                myW1 = ""
                If BranchPOSum.chkPO.Value = 0 Then
                        myW1 = "AND spordrconfirm.REFERENCE = '" & myREFNO & "'"
                Else
                    For i = 0 To flxAllPOM.Rows - 1
                        If i = 1 Then
                            myW1 = myW1 & " AND spordrconfirm.REFERENCE = '" & flxAllPOM.TextMatrix(i, 0) & "'"
                        Else
                            myW1 = myW1 & " OR spordrconfirm.REFERENCE = '" & flxAllPOM.TextMatrix(i, 0) & "'"
                        End If
                    Next i
                End If
                myW = myW1
                 
            'Qry = "SELECT SPORDERTAKING.SQSERVE  From SPORDERTAKING Where SPORDERTAKING.BRANCHID = " & Val(lblBID.Caption) & " AND SPORDERTAKING.FINALQT = 0;"
            'ASQSERVE = Val(GetFieldValue(Qry))
            
            'Qry = "SELECT spordrconfirm.ORDERID From spordrconfirm Where spordrconfirm.SQSERVE = " & ASQSERVE & ";"
           
            'If .State = adStateOpen Then .Close
            '.Open Qry, LYRcon, adOpenKeyset, adLockPessimistic
            'If .RecordCount > 0 Then
            '        MsgBox ("Please Print Previous Order Taking"), vbExclamation, "   Not Confirm P.O"
            '        .Close
            '         Call ClearEntries
            '        Exit Sub
            'End If
                With BranchPOSum.flxSP
                    fl = 0
                    For nRow = 1 To .Rows - 1
                        If Val(.TextMatrix(nRow, 6)) > 0 Then
                                fl = fl + 1
                                flxSP.Rows = fl + 1
                                flxSP.TextMatrix(fl, 9) = .TextMatrix(nRow, 4)
                                flxSP.TextMatrix(fl, 10) = ""
                                flxSP.TextMatrix(fl, 14) = .TextMatrix(nRow, 9)
                                flxSP.TextMatrix(fl, 8) = .TextMatrix(nRow, 6)
                                flxSP.TextMatrix(fl, 13) = .TextMatrix(nRow, 3)
                                
                        End If
                    Next nRow
                   
                End With
                
                If flxSP.Rows = 2 And flxSP.TextMatrix(1, 13) = "" Then flxAllPOM.Clear: flxAllPOM.Rows = 2: flxAllPOM.FixedRows = 1: Exit Sub
                If .State = adStateOpen Then .Close
                For nRow = 1 To flxSP.Rows - 1
                        Qry = "SELECT spbranchdtl.SPAREPARTID, spareparts.PARTNUMBER, spareparts.DESCRIPTION, spareparts.QTYONHAND, spareparts.BACKORDER, spbranchdtl.BQTYONHAND, " & _
                              "spareparts.cost, spareparts.WSQTYONHAND From spbranchdtl INNER JOIN spareparts ON (spbranchdtl.SPAREPARTID = spareparts.SPAREPARTID) " & _
                              "Where spareparts.SPAREPARTID =  " & Val(flxSP.TextMatrix(nRow, 13)) & " AND spbranchdtl.REFERENCE ='" & flxSP.TextMatrix(nRow, 13) & "' AND spbranchdtl.BRANCHNAME = '" & mtrBName & "';"
                        If .State = adStateOpen Then .Close
                        .Open Qry, LYRcon, adOpenKeyset, adLockPessimistic
                        If .RecordCount > 0 Then
                            flxSP.TextMatrix(nRow, 1) = !PartNumber
                            flxSP.TextMatrix(nRow, 2) = !Description
                            flxSP.TextMatrix(nRow, 3) = !QtyOnHand
                            flxSP.TextMatrix(nRow, 4) = !WSQTYONHAND
                            flxSP.TextMatrix(nRow, 5) = !BACKORDER
                            flxSP.TextMatrix(nRow, 6) = Val(GetSalesQtyMain(flxSP.TextMatrix(nRow, 13)))
                            flxSP.TextMatrix(nRow, 7) = !BQTYONHAND
                            flxSP.TextMatrix(nRow, 11) = Format(!Cost, "#,###,###.#0")
                            flxSP.TextMatrix(nRow, 12) = "0.00"
                        End If
                 Next nRow
                 cmdTrans(1).Enabled = True
                 cmdTrans(3).Enabled = True
                 
                 .Close
    End With
End Sub
Sub ShowRequest(refnumber, bORDERID)
Dim myWHQty As Long
Dim rsPOdetails As New ADODB.Recordset
Dim rsSPdetails As New ADODB.Recordset
    
   myWHQty = 0
   myBranchCategory = ""
   
   Qry = "SELECT sporderbranch.BRANCHNAME, sporderbranch.SPORDERID, sporderbranch.ORDERTYPE, sporderbranch.MOS, sporderbranch.ORDERDATE, sporderbranch.REFERENCE, " & _
            "sporderbranch.BRANCHSPOID, sporderbranch.BRANCHID, sporderbranch.`STATUS`, sporderbranch.ALLOCATED, sporderbranch.PRINTPICK, sporderbranch.PICKEDBY " & _
            "From sporderbranch Where sporderbranch.REFERENCE = '" & refnumber & "' AND sporderbranch.BRANCHSPOID = " & bORDERID & ";"
    
    With rsPOdetails
        txtStatus.Enabled = False
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            txtbranch.Text = !branchName
            txtReference.Text = !REFERENCE
            txtStatus.Text = IIf(IsNull(!Status), "", !Status)
            txtBranchSPOID.Text = !BRANCHSPOID
            txtBranchID.Text = !branchID
            txtSPOID.Text = !SPORDERID
            txtPrnCopy.Text = !PRINTPICK
            txtAllocQty.Text = !ALLOCATED
            txtOrderDate.Text = !ORDERDATE
            txtPreparedby.Text = IIf(IsNull(!PICKEDBY), "", !PICKEDBY)
            cboMos.Text = IIf(IsNull(!mos), "6", !mos)
            txtOrderType.Text = IIf(IsNull(!ORDERTYPE), "", !ORDERTYPE)
        End If
            If Val(txtAllocQty.Text) > 0 Then
                cmdTrans(1).Caption = "Edit P.O."
            End If
    End With
        
    myBranchCategory = GetFieldValue("Select BRANCHCATEGORY From BRANCH WHERE BRANCHID = " & Val(txtBranchID.Text) & "")
    
    Qry = "SELECT spbranchdtl.SPAREPARTID, spareparts.PARTNUMBER, spareparts.DESCRIPTION, spbranchdtl.BQTYONHAND, spbranchdtl.`ORDER`, spbranchdtl.ORDERBAL, spbranchdtl.BACKORDERQTY, " & _
            "spareparts.COST, spbranchdtl.ALLOCATEDQTY, spareparts.QTYONHAND, spareparts.BACKORDER, spareparts.LOCATION, spbranchdtl.REFERENCE, spbranchdtl.BRANCHSPOID,spareparts.WSQTYONHAND " & _
            "From spbranchdtl INNER JOIN spareparts ON (spbranchdtl.SPAREPARTID = spareparts.SPAREPARTID) " & _
            "Where spbranchdtl.REFERENCE = '" & refnumber & "' AND spbranchdtl.BRANCHSPOID = " & bORDERID & " Order By spareparts.LOCATION;"
            
        With rsSPdetails
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                        flxSP.Rows = 2
                        .MoveFirst
                            Do While Not .EOF
                                Dim reserveQty As Integer
                                reserveQty = GETSUMofALQ(!SPAREPARTID)
                                myWHQty = !QtyOnHand - IIf(IsNull(!WSQTYONHAND), 0, !WSQTYONHAND)
                                If myWHQty <= 0 Then
                                    myWHQty = 0
                                End If
                                'flxSP.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & !QtyOnHand & vbTab & Val(GETSUMofALQ(!SPAREPARTID)) & vbTab & IIf(Format(!QtyOnHand >= GETSUMofALQ(!SPAREPARTID)), Val(!QtyOnHand - GETSUMofALQ(!SPAREPARTID)), !QtyOnHand) & vbTab & !BQTYONHAND & vbTab & !BACKORDERQTY & vbTab & !Order & vbTab & !ALLOCATEDQTY & vbTab & Format(!Cost, "###0.00") & vbTab & Format(!ALLOCATEDQTY * !Cost, "###0.00") & vbTab & IIf(!Location = "", "NO LOCATION", !Location) & vbTab & !SPAREPARTID & vbTab & !REFERENCE
                                flxSP.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & myWHQty & vbTab & IIf(IsNull(!WSQTYONHAND), 0, !WSQTYONHAND) & vbTab & Val(reserveQty) & vbTab & IIf(Format((myWHQty + IIf(IsNull(!WSQTYONHAND), 0, !WSQTYONHAND)) >= reserveQty), Val((myWHQty + IIf(IsNull(!WSQTYONHAND), 0, !WSQTYONHAND)) - reserveQty), (myWHQty + IIf(IsNull(!WSQTYONHAND), 0, !WSQTYONHAND))) & vbTab & !BQTYONHAND & vbTab & !BACKORDERQTY & vbTab & !Order & vbTab & !ALLOCATEDQTY & vbTab & Format(!Cost, "###0.00") & vbTab & Format(!ALLOCATEDQTY * !Cost, "###0.00") & vbTab & IIf(!Location = "", "NO LOCATION", !Location) & vbTab & !SPAREPARTID & vbTab & !REFERENCE
                                .MoveNext
                                DoEvents
                            Loop
                End If
        End With
        
        With flxSP
            For i = .Rows - 1 To 1 Step -1
                If Val(.TextMatrix(i, 8)) = 0 Then
                    .TextMatrix(i, 8) = Val(.TextMatrix(i, 9)) - Val(.TextMatrix(i, 10))
                End If
            Next i
        End With
    
    Call ComputeTotalQty
    Call ComputeTotalAllocation
    Call ResizeCompleteHead
End Sub

Private Sub ComputeTotalQty()
    Dim i As Integer, Ttl As Long
    
    Ttl = 0
    With flxSP
        For i = 1 To .Rows - 1
            Ttl = Ttl + Val(.TextMatrix(i, 12))
        Next i
    End With
    txtTotalAmount.Text = Format(Ttl, "#,###,##0.00")
End Sub


Sub ShowRecordOT(SVRid As Long) 'ORDERIDSQ

Dim Qry As String, nRow As Long, ref As Integer
Dim rsPO As New ADODB.Recordset, rsRef As New ADODB.Recordset
    
    flxSP.Rows = 1: flxSP.Rows = 2
    fmHeader.Enabled = False: flxSP.Enabled = False
    myOrderIDSQ = SVRid
    With rsPO
        
        Qry = "SELECT spordertaking.ORDERIDSQ, spordertaking.BRANCHID, spordertaking.BRANCHNAME, spordertaking.MOS, spordertaking.DATESERVE, spordertaking.SQSERVE , spordertaking.REMARKS, spordertaking.PREPAREDBY, " & _
                "spordertaking.APPROVEDBY, spordertaking.PRINTPICK, spordertaking.DATECONFIRM, spordertaking.FINALQT " & _
                "From spordertaking WHERE spordertaking.SQSERVE = " & myOrderIDSQ & ";"
        
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockPessimistic
        If .RecordCount > 0 Then
            .MoveFirst
            lblBID.Caption = !branchID
            txtbranch.Text = !branchName
            cboMos.Text = !mos
            txtDateServe.Text = !DATESERVE
            txtDateCon.Text = IIf(IsNull(!DATECONFIRM), "", !DATECONFIRM)
            lblSQSERVE.Caption = !SQSERVE
            myOrderIDSQ = !ORDERIDSQ
            cbodrtype.Text = "(04) - STOCK"
            txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
            txtPreparedby.Text = IIf(IsNull(!PreparedBy), "", !PreparedBy)
            txtApprovedBy.Text = IIf(IsNull(!ApprovedBy), "", !ApprovedBy)
            If !PRINTPICK >= 1 And Not IsNull(!DATECONFIRM) Then
                cmdTrans(1).Enabled = False
                mnupick.Enabled = False
                mnuconfirm.Enabled = True
            ElseIf !PRINTPICK >= 1 And IsNull(!DATECONFIRM) Then
                cmdTrans(1).Enabled = True
                mnupick.Enabled = True
                mnuconfirm.Enabled = True
            ElseIf !PRINTPICK = 0 Then
                cmdTrans(1).Enabled = True
                mnupick.Enabled = True
                mnuconfirm.Enabled = False
            End If
            If .Fields("FINALQT") = 0 Then
                lblSTATUS.Caption = "PENDING"
            ElseIf .Fields("FINALQT") = 1 Then
                lblSTATUS.Caption = "FINAL"
            ElseIf .Fields("FINALQT") = 2 Then
                lblSTATUS.Caption = "CANCELLED"
            End If
            .Close
            
        End If
        
        Qry = "SELECT spordrconfirm.ORDERID FROM spordrconfirm WHERE spordrconfirm.SQSERVE = " & Val(lblSQSERVE.Caption) & " AND  spordrconfirm.BRANCHID = " & Val(lblBID.Caption) & ";"
        myOrderIDCof = Val(GetFieldValue(Qry))
        
        Qry = "SELECT spbrchconfirmdtl.REFERENCE, spbrchconfirmdtl.SPAREPARTID, spbrchconfirmdtl.BPORDER, spbrchconfirmdtl.BBACKORDER, spbrchconfirmdtl.ORDERTAKING, spbrchconfirmdtl.PRICE, spbrchconfirmdtl.REMARKS " & _
               "From spbrchconfirmdtl Where spbrchconfirmdtl.ORDERID =" & myOrderIDCof & ";"
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockPessimistic
             If .RecordCount = 0 Then Exit Sub: .Close
                 flxSP.Rows = .RecordCount + 1
                  .MoveFirst
                    For nRow = 1 To .RecordCount
                          flxSP.TextMatrix(nRow, 9) = rsPO(2).Value
                          flxSP.TextMatrix(nRow, 10) = IIf(IsNull(!ORDERTAKING) Or !ORDERTAKING = 0, "", !ORDERTAKING)
                          flxSP.TextMatrix(nRow, 8) = IIf(IsNull(!BBACKORDER), 0, !BBACKORDER)
                          flxSP.TextMatrix(nRow, 13) = IIf(IsNull(!REFERENCE), "", !REFERENCE)
                          flxSP.TextMatrix(nRow, 14) = !SPAREPARTID
                         .MoveNext
                     Next nRow
                    .Close
                    Qry = "SELECT spordrconfirm.REFERENCE FROM spordrconfirm WHERE spordrconfirm.SQSERVE = " & Val(lblSQSERVE.Caption) & " AND spordrconfirm.BRANCHID = " & Val(lblBID.Caption) & ";"
                    If rsRef.State = adStateOpen Then rsRef.Close
                    rsRef.Open Qry, LYRcon, adOpenKeyset, adLockPessimistic
                    rsRef.MoveFirst
                     flxAllPOM.Clear
                    Do While Not rsRef.EOF
                        flxAllPOM.AddItem rsRef!REFERENCE
                        rsRef.MoveNext
                    Loop
                        flxAllPOM.Rows = 2
                        flxAllPOM.FixedRows = 1
                    
                     For nRow = 1 To flxSP.Rows - 1
                        Qry = "SELECT spbrchconfirmdtl.SPAREPARTID, spareparts.PARTNUMBER, spareparts.DESCRIPTION, spareparts.QTYONHAND, spareparts.BACKORDER, spbrchconfirmdtl.BQTYONHAND, " & _
                              "spareparts.COST From spbrchconfirmdtl INNER JOIN spareparts ON (spbrchconfirmdtl.SPAREPARTID = spareparts.SPAREPARTID) " & _
                              "Where spareparts.SPAREPARTID =  " & flxSP.TextMatrix(nRow, 13) & " AND spbrchconfirmdtl.ORDERID=" & myOrderIDCof & " ORDER BY spareparts.PARTNUMBER;"
                        If .State = adStateOpen Then .Close
                        .Open Qry, LYRcon, adOpenKeyset, adLockPessimistic
                        flxSP.TextMatrix(nRow, 0) = nRow
                        flxSP.TextMatrix(nRow, 1) = !PartNumber
                        flxSP.TextMatrix(nRow, 2) = !Description
                        flxSP.TextMatrix(nRow, 3) = !QtyOnHand
                        flxSP.TextMatrix(nRow, 5) = !BACKORDER
                        flxSP.TextMatrix(nRow, 6) = Val(GetSalesQtyMain(flxSP.TextMatrix(nRow, 14)))
                        flxSP.TextMatrix(nRow, 7) = !BQTYONHAND
                        flxSP.TextMatrix(nRow, 11) = Format(!Cost, "#,###,###.#0")
                    Next nRow
                    Call ComputeTotalView
        .Close
    End With
End Sub
Private Sub cbomos_Change()
    cbomos_Click
End Sub
Private Sub cbomos_Click()
Dim mos As Integer
    mos = Val(cboMos.Text)
    pmos = "-" & mos
    lblDate.Caption = DateAdd("m", pmos, Date) & "  TO  " & Date
    If cboMos.Text = "All" Then
          lblDate.Caption = "12/19/2007" & "  TO  " & Date
    End If
End Sub
Function GetSalesQtyMain(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", pmos, "12/19/2007"), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", pmos, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValue(Qry))
    GetSalesQtyMain = MySales - MyRet
End Function

Function GETSUMofALQ(mySPID As Long) As Long
    Dim rsALQ As New ADODB.Recordset
    Dim sumALQ As Long
    
    Qry = "SELECT SUM(spbranchdtl.ALLOCATEDQTY) As SUMofALQ From spbranchdtl " & _
            "Where spbranchdtl.TRANSTYPE = 'ORDER TEMP' AND spbranchdtl.REMARKS = 'PACKING'  AND spbranchdtl.SPAREPARTID = " & mySPID & ";"
    With rsALQ
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            sumALQ = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            sumALQ = 0
        End If
        .Close
    End With
    GETSUMofALQ = sumALQ
    
End Function

Function GETAVQ(myIDSP As Long) As Long
    Dim rsAVQ As New ADODB.Recordset
    Dim AVQ As Long
    
    Qry = "SELECT Spareparts.QTYONHAND From spbranchdtl INNER JOIN spareparts ON (spbranchdtl.SPAREPARTID = spareparts.SPAREPARTID) " & _
            "Where spbranchdtl.SPAREPARTID = " & myIDSP & " AND spbranchdtl.REFERENCE = '" & refnumber & "' AND spbranchdtl.BRANCHSPOID = " & bORDERID & ";"
            
    With rsAVQ
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            AVQ = IIf(IsNull(!QtyOnHand), 0, !QtyOnHand)
        Else
            AVQ = 0
        End If
        .Close
    End With
    GETAVQ = AVQ
    
End Function






Private Sub txtOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyEscape Then
        If flxSP.col = 10 Then
            If ValidateEntry = True Then
                txtOrder.Text = ""
                txtOrder.Visible = False
                flxSP.SetFocus
                flxSP.row = ActRow
                flxSP.col = 10
                Call ComputeTotalAllocation
                txtOrder.ZOrder
                Exit Sub
            End If
                
                txtOrder.Text = flxSP.TextMatrix(ActRow, 10)
                flxSP.SetFocus
                flxSP.row = ActRow
                flxSP.col = 10
                Call ComputeTotalAllocation
                txtOrder.Visible = False
        End If
        
    End If
End Sub
Private Function ValidateEntry() As Boolean
    With txtOrder
        If txtOrder.Text = "" Then Exit Function
        'If flxsp.Col = 9 Then
           ' If Val(.Text) > flxSP.TextMatrix(ActRow, 7) Or Val(.Text) > flxSP.TextMatrix(ActRow, 8) Then
            '    MsgBox ("Order Taking not Greater Than the Backorder, Ordered parts"), vbExclamation, "Order"
             '   .SelStart = 0
             '   .SelLength = Len(txtOrder.Text)
             '   .ZOrder
             '   .SetFocus
             '   ValidateEntry = True
'            If myBranchCategory <> "WHOLESALE" Then
                If Val(.Text) > flxSP.TextMatrix(ActRow, 9) Then
                    MsgBox ("Allocation of Order Quantity is greater than the Ordered parts"), vbExclamation, "Order Allocation"
                    .SelStart = 0
                    .SelLength = Len(txtOrder.Text)
                    .ZOrder
                    .SetFocus
                    ValidateEntry = True
                ElseIf Val(.Text) > flxSP.TextMatrix(ActRow, 6) Then
                    MsgBox ("Allocation of Order Quantity is greater than the Available quantity"), vbExclamation, "Order"
                    .SelStart = 0
                    .SelLength = Len(txtOrder.Text)
                    .ZOrder
                    .SetFocus
                    ValidateEntry = True
                Else
                    ValidateEntry = False
                End If
'            Else
'                If Val(.Text) > flxSP.TextMatrix(ActRow, 4) Then
'                    MsgBox ("Allocation of Order Quantity is greater than the Wholesale Maximum Quantity"), vbExclamation, "Order"
'                    .SelStart = 0
'                    .SelLength = Len(txtOrder.Text)
'                    .ZOrder
'                    .SetFocus
'                    ValidateEntry = True
'                Else
'                    ValidateEntry = False
'                End If
'            End If
            'ElseIf flxsp.Col = 8 Then
            '    If Val(.Text) > Val(flxsp.TextMatrix(ActRow, 3)) Then
            '        MsgBox ("Quantity Order not Greater Than Available stock"), vbExclamation, "Order"
            '        .SelStart = 0
            '        .SelLength = Len(txtOrder.Text)
            '        .ZOrder
            '        .SetFocus
            '        ValidateEntry = True
            '    End If
            'End If
     End With
End Function
Private Sub txtOrder_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
            flxSP.SetFocus
            flxSP.row = ActRow
            
            flxSP.col = 10
    End If
    Call ComputeTotalAllocation
End Sub
Private Sub txtOrder_LostFocus()

    With flxSP
        If ValidateEntry = True Then Exit Sub
        
        If .col = 10 Then
            .TextMatrix(ActRow, 10) = txtOrder.Text
             txtOrder.Visible = False
            Call ComputeTotal
            Call ComputeTotalAllocation
            If Val(.TextMatrix(ActRow, 10)) >= 0 Then
               If Val(.TextMatrix(ActRow, 9)) > Val(.TextMatrix(ActRow, 10)) Then   'OQ>AQ THEN BO=ALQ-AVLQ
                   .TextMatrix(ActRow, 8) = Val(.TextMatrix(ActRow, 9)) - Val(.TextMatrix(ActRow, 10))
               ElseIf Val(.TextMatrix(ActRow, 9)) = Val(.TextMatrix(ActRow, 10)) Then   'OQ=AQ THEN BO=0
                   .TextMatrix(ActRow, 8) = 0
               ElseIf Val(.TextMatrix(ActRow, 10)) = 0 Then
                   .TextMatrix(ActRow, 8) = Val(.TextMatrix(ActRow, 9))
               End If
                .row = ActRow
                .col = 10
                '.CellFontName = "WINGDINGS"
                '.CellFontSize = 12
                '.Text = lblChk.Caption
                '.CellAlignment = 4
                .col = 10
            Else
                .TextMatrix(ActRow, 10) = ""
                .TextMatrix(ActRow, 0) = ""
                .col = 10
            End If
        ElseIf .col = 9 Then
            .TextMatrix(ActRow, 9) = txtOrder.Text
             txtOrder.Visible = False
             Call ComputeTotal
             Call ComputeTotalAllocation
             If Val(.TextMatrix(ActRow, 9)) > 0 Then
                 .row = ActRow
                 .col = 0
                 '.CellFontName = "WINGDINGS"
                 '.CellFontSize = 12
                 '.Text = lblChk.Caption
                 '.CellAlignment = 4
                 .col = 9
             Else
                 .TextMatrix(ActRow, 9) = ""
                 .TextMatrix(ActRow, 0) = ""
                 .col = 9
             End If
        
        End If
        'txtOrder.Text = ""
        'txtOrder.Visible = False
     End With
     
End Sub
Sub ComputeTotal()
Dim nrw As Integer, GrnTl As Double, SPTotal As Double
    With flxSP
        SPTotal = Val(flxSP.TextMatrix(ActRow, 10)) * Val(Replace(flxSP.TextMatrix(ActRow, 11), ",", ""))
        flxSP.TextMatrix(ActRow, 12) = Format(SPTotal, "###,###,##0.#0")
        For nrw = 1 To .Rows - 1
            GrnTl = GrnTl + Val(Replace(flxSP.TextMatrix(nrw, 12), ",", ""))
        Next nrw
    End With
    txtTotalAmount.Text = Format(GrnTl, "###,###,##0.#0")
End Sub
Sub ComputeTotalView()
Dim nrw As Integer, GrnTl As Double, SPTotal As Double
    
    With flxSP
        For nrw = 1 To .Rows - 1
            SPTotal = Val(flxSP.TextMatrix(nrw, 10)) * Val(Replace(flxSP.TextMatrix(nrw, 11), ",", ""))
            flxSP.TextMatrix(nrw, 12) = Format(SPTotal, "###,###,##0.#0")
            GrnTl = GrnTl + Val(Replace(flxSP.TextMatrix(nrw, 12), ",", ""))
        Next nrw
    End With
    txtTotalAmount.Text = Format(GrnTl, "###,###,##0.#0")
End Sub

Public Sub ComputeTotalAllocation()
    Dim Ttl As Single, i As Integer
        
    Ttl = 0
    With flxSP
        For i = 1 To .Rows - 1
               Ttl = Ttl + Val(Replace(.TextMatrix(i, 10), ",", ""))
        Next i
    End With
    txtTotalAlloc.Text = Format(Ttl, "0")
    
End Sub




