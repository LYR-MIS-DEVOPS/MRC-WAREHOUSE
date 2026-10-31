VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSummary 
   Caption         =   "Form Summary"
   ClientHeight    =   8595
   ClientLeft      =   225
   ClientTop       =   -5655
   ClientWidth     =   13845
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   13845
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab ssTabTRec 
      Height          =   4335
      Left            =   3600
      TabIndex        =   60
      Top             =   1920
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "TREASURY PORTION"
      TabPicture(0)   =   "FrmSummary.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label8"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label10"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label11"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtTDateDeposit"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtTChckDate"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtTRecDate"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "btnTCancel"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "btnTOk"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtTRecBy"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtTBankAcct"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtTChckNo"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtTCVNo"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtTRFPNO"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      Begin VB.TextBox txtTRFPNO 
         Height          =   495
         Left            =   5400
         TabIndex        =   71
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtTCVNo 
         Height          =   375
         Left            =   1515
         TabIndex        =   70
         Top             =   2850
         Width           =   3375
      End
      Begin VB.TextBox txtTChckNo 
         Height          =   375
         Left            =   1515
         TabIndex        =   67
         Top             =   1990
         Width           =   3375
      End
      Begin VB.TextBox txtTBankAcct 
         Height          =   375
         Left            =   1515
         TabIndex        =   65
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox txtTRecBy 
         Height          =   375
         Left            =   1515
         TabIndex        =   62
         Top             =   520
         Width           =   3375
      End
      Begin lvButton.lvButtons_H btnTOk 
         Height          =   375
         Left            =   720
         TabIndex        =   72
         Tag             =   "XYF"
         Top             =   3840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "&CONFIRM"
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
         Image           =   "FrmSummary.frx":001C
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H btnTCancel 
         Height          =   375
         Left            =   3000
         TabIndex        =   73
         Tag             =   "XYF"
         Top             =   3840
         Width           =   1215
         _ExtentX        =   2143
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
         Image           =   "FrmSummary.frx":05B6
         cBack           =   -2147483633
      End
      Begin MSMask.MaskEdBox txtTRecDate 
         Height          =   375
         Left            =   1515
         TabIndex        =   74
         Tag             =   "XYF"
         Top             =   960
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTChckDate 
         Height          =   375
         Left            =   1515
         TabIndex        =   75
         Tag             =   "XYF"
         Top             =   2400
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTDateDeposit 
         Height          =   375
         Left            =   1515
         TabIndex        =   77
         Tag             =   "XYF"
         Top             =   3270
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label11 
         Caption         =   "DATE DEPOSIT:"
         Height          =   255
         Left            =   240
         TabIndex        =   76
         Top             =   3330
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "CV NO.              :"
         Height          =   255
         Left            =   240
         TabIndex        =   69
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "CHECK DATE    :"
         Height          =   255
         Left            =   240
         TabIndex        =   68
         Top             =   2475
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "CHECK NO.       :"
         Height          =   255
         Left            =   240
         TabIndex        =   66
         Top             =   2070
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "BANK ACCT.      :"
         Height          =   255
         Left            =   240
         TabIndex        =   64
         Top             =   1650
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "RECEIVED ON  :"
         Height          =   255
         Left            =   240
         TabIndex        =   63
         Top             =   1035
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "RECEIVED BY   :"
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   600
         Width           =   1455
      End
   End
   Begin TabDlg.SSTab SSConfirmTab 
      Height          =   1575
      Left            =   14040
      TabIndex        =   55
      Tag             =   "XYF"
      Top             =   2640
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   2778
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Confirmation"
      TabPicture(0)   =   "FrmSummary.frx":0B50
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCancel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdOK"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtConfirmDate"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin MSMask.MaskEdBox txtConfirmDate 
         Height          =   375
         Left            =   2120
         TabIndex        =   57
         Tag             =   "XYF"
         Top             =   420
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin lvButton.lvButtons_H cmdOK 
         Height          =   375
         Left            =   840
         TabIndex        =   58
         Tag             =   "XYF"
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "&CONFIRM"
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
         Image           =   "FrmSummary.frx":0B6C
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   375
         Left            =   2040
         TabIndex        =   59
         Tag             =   "XYF"
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
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
         Image           =   "FrmSummary.frx":1106
         cBack           =   -2147483633
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Change Confirmation Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Tag             =   "XYF"
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Timer tmrAutoLoad 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   11040
      Top             =   0
   End
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10560
      Top             =   0
   End
   Begin TabDlg.SSTab tabLoadPO 
      Height          =   7455
      Left            =   120
      TabIndex        =   46
      Top             =   9000
      Visible         =   0   'False
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   13150
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "FrmSummary.frx":16A0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtBID"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtSPOID"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "flxOrder"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "pBarOrder"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtbName"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "flxBranch"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "flxPOdetails"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtOption"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.TextBox txtOption 
         Height          =   855
         Left            =   9480
         TabIndex        =   54
         Top             =   840
         Width           =   3255
      End
      Begin MSFlexGridLib.MSFlexGrid flxPOdetails 
         Height          =   5445
         Left            =   3960
         TabIndex        =   53
         Top             =   3120
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   9604
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBranch 
         Height          =   2535
         Left            =   3000
         TabIndex        =   52
         Top             =   120
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   4471
         _Version        =   393216
         Cols            =   4
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin VB.TextBox txtbName 
         Height          =   375
         Left            =   5160
         TabIndex        =   51
         Top             =   6240
         Width           =   2535
      End
      Begin MSComctlLib.ProgressBar pBarOrder 
         Height          =   255
         Left            =   600
         TabIndex        =   50
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxOrder 
         Height          =   2175
         Left            =   3840
         TabIndex        =   49
         Top             =   3000
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   3836
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
      End
      Begin VB.TextBox txtSPOID 
         Height          =   405
         Left            =   5160
         TabIndex        =   48
         Top             =   6720
         Width           =   2535
      End
      Begin VB.TextBox txtBID 
         Height          =   375
         Left            =   5160
         TabIndex        =   47
         Top             =   5760
         Width           =   2535
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxMain 
      Height          =   1095
      Left            =   600
      TabIndex        =   45
      Top             =   10000
      Visible         =   0   'False
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   1931
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
   End
   Begin VB.TextBox txtBranch1 
      Height          =   285
      Left            =   3240
      TabIndex        =   44
      Top             =   9600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxAdd 
      Height          =   1095
      Left            =   8400
      TabIndex        =   43
      Top             =   9960
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1931
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.TextBox txtCutOff 
      Height          =   285
      Left            =   14280
      TabIndex        =   34
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtYr 
      Height          =   285
      Left            =   14280
      TabIndex        =   33
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtMo 
      Height          =   285
      Left            =   15120
      TabIndex        =   32
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtCancelRsn 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1440
      TabIndex        =   28
      Top             =   9000
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.TextBox txtCriteria 
      Height          =   285
      Left            =   600
      TabIndex        =   18
      Top             =   9480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   120
      TabIndex        =   17
      Tag             =   "XYF"
      Top             =   8400
      Visible         =   0   'False
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox txtFSCamt 
      Height          =   285
      Left            =   3000
      TabIndex        =   16
      Top             =   8760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSum 
      Height          =   6975
      Left            =   120
      TabIndex        =   11
      Tag             =   "XYF"
      Top             =   1440
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   12303
      _Version        =   393216
      RowHeightMin    =   285
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Tag             =   "XY"
      Top             =   720
      Width           =   13695
      Begin MSComDlg.CommonDialog dlgDEMO 
         Left            =   7440
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin lvButton.lvButtons_H cmdUpdate 
         Height          =   375
         Left            =   9000
         TabIndex        =   42
         Tag             =   "XY"
         Top             =   150
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Update P.O."
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
         cBack           =   -2147483633
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   11640
         TabIndex        =   3
         Tag             =   "XY"
         Top             =   195
         Width           =   1935
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   10320
         TabIndex        =   4
         Tag             =   "XY"
         Top             =   150
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
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
         Image           =   "FrmSummary.frx":16BC
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dFrom 
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Top             =   180
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   150339585
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         ToolTipText     =   "Show Date End"
         Top             =   150
         Width           =   375
         _ExtentX        =   661
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
         Image           =   "FrmSummary.frx":1C56
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   375
         Left            =   5400
         TabIndex        =   12
         ToolTipText     =   "Execute"
         Top             =   120
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
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
         Image           =   "FrmSummary.frx":21F0
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dTo 
         Height          =   315
         Left            =   3720
         TabIndex        =   13
         Top             =   180
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   150339585
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdPrev 
         Height          =   375
         Left            =   5880
         TabIndex        =   14
         ToolTipText     =   "Hide Date End"
         Top             =   150
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
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
         Image           =   "FrmSummary.frx":278A
         cBack           =   -2147483633
      End
      Begin VB.Label Label1 
         Caption         =   "Date Covered:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   210
         Width           =   1335
      End
   End
   Begin lvButton.lvButtons_H cmdShow 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "Show &Form"
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
      Image           =   "FrmSummary.frx":2D24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Print Summary"
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
      Image           =   "FrmSummary.frx":32BE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   3600
      TabIndex        =   8
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
      Image           =   "FrmSummary.frx":3858
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxCancel 
      Height          =   735
      Left            =   6720
      TabIndex        =   15
      Top             =   9240
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1296
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      FormatString    =   "REFERENCE ID.|^VALUE1  |^VALUE2"
   End
   Begin lvButton.lvButtons_H cmdFilter 
      Height          =   375
      Left            =   5400
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "Filter &Records"
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
      Image           =   "FrmSummary.frx":3DF2
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   21
      Tag             =   "XYF"
      Top             =   720
      Visible         =   0   'False
      Width           =   11655
      Begin VB.TextBox txtFind2 
         Height          =   285
         Left            =   9600
         TabIndex        =   22
         Tag             =   "XYF"
         Top             =   210
         Width           =   1935
      End
      Begin VB.ComboBox cboMonth 
         Height          =   315
         Left            =   1440
         TabIndex        =   25
         Tag             =   "XYF"
         Top             =   195
         Width           =   1695
      End
      Begin VB.TextBox txtYear 
         Height          =   315
         Left            =   3135
         TabIndex        =   24
         Tag             =   "XYF"
         Top             =   195
         Width           =   720
      End
      Begin MSComCtl2.UpDown UpdYear 
         Height          =   315
         Left            =   3840
         TabIndex        =   23
         Tag             =   "XYF"
         Top             =   195
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Value           =   2005
         OrigLeft        =   11310
         OrigTop         =   480
         OrigRight       =   11565
         OrigBottom      =   795
         Max             =   2020
         Min             =   2005
         Enabled         =   -1  'True
      End
      Begin lvButton.lvButtons_H cmdFind2 
         Height          =   375
         Left            =   8280
         TabIndex        =   27
         Tag             =   "XY"
         Top             =   165
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
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
         Image           =   "FrmSummary.frx":438C
         cBack           =   -2147483633
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "For the month of"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Tag             =   "XYF"
         Top             =   225
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxPay 
      Height          =   1935
      Left            =   14640
      TabIndex        =   29
      Top             =   3360
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorBkg    =   -2147483633
      FormatString    =   "<OR DATE   |<OR NUMBER |<OR AMOUNT |< REBATE  |< OR TYPE  |< REMARKS   "
   End
   Begin MSFlexGridLib.MSFlexGrid flxLedger 
      Height          =   1695
      Left            =   14040
      TabIndex        =   30
      Top             =   480
      Visible         =   0   'False
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   1
      Cols            =   12
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483633
      GridColor       =   16744576
      GridColorFixed  =   16744576
      ScrollBars      =   2
      SelectionMode   =   1
      BorderStyle     =   0
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxAccts 
      Height          =   1335
      Left            =   14640
      TabIndex        =   31
      Top             =   5400
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2355
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin lvButton.lvButtons_H cmdChkClosed 
      Height          =   375
      Left            =   13320
      TabIndex        =   35
      Tag             =   "xyf"
      Top             =   0
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
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
      Image           =   "FrmSummary.frx":4926
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxAge1 
      Height          =   975
      Left            =   13920
      TabIndex        =   36
      Tag             =   "XY"
      Top             =   6840
      Visible         =   0   'False
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   1720
      _Version        =   393216
      Cols            =   29
      FixedCols       =   0
      BackColorBkg    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   $"FrmSummary.frx":4EC0
   End
   Begin TabDlg.SSTab tabWait 
      Height          =   2295
      Left            =   4680
      TabIndex        =   37
      Top             =   8520
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4048
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "FrmSummary.frx":5016
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProcess"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Image2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Image Image2 
         Height          =   720
         Left            =   360
         Picture         =   "FrmSummary.frx":5032
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "PLEASE WAIT...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1200
         TabIndex        =   39
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label lblProcess 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Processing..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   38
         Top             =   720
         Width           =   3975
      End
   End
   Begin lvButton.lvButtons_H cmdTransmittal 
      Height          =   375
      Left            =   7080
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Transmittal"
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
      Image           =   "FrmSummary.frx":5A2A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdSalesDrive 
      Height          =   375
      Left            =   8760
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Sales Drive"
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
      Image           =   "FrmSummary.frx":5FC4
      cBack           =   -2147483633
   End
   Begin VB.Label lblRemarks 
      Height          =   255
      Left            =   600
      TabIndex        =   19
      Top             =   9240
      Visible         =   0   'False
      Width           =   3735
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   360
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14085
      m_MDI_Height    =   9165
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Form SUMMARY"
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
      Height          =   375
      Index           =   2
      Left            =   9240
      TabIndex        =   0
      Top             =   450
      Width           =   4500
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Form SUMMARY"
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
      Height          =   375
      Index           =   3
      Left            =   9255
      TabIndex        =   1
      Top             =   480
      Width           =   4500
   End
   Begin VB.Menu mnuSum 
      Caption         =   "Summary"
      Visible         =   0   'False
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
      End
      Begin VB.Menu mnuConfirm 
         Caption         =   "Confirm"
      End
   End
   Begin VB.Menu mnuPurchRet 
      Caption         =   "Purchase Return"
      Visible         =   0   'False
      Begin VB.Menu mnuHistory 
         Caption         =   "Show History"
      End
      Begin VB.Menu mnuRetStatus 
         Caption         =   "Update Status"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancelPR 
         Caption         =   "Cancel Purchase Return"
      End
   End
   Begin VB.Menu mnuCancelReason 
      Caption         =   "Cancel Reason"
      Visible         =   0   'False
      Begin VB.Menu mnuViewCancel 
         Caption         =   "View Cancel Details"
      End
   End
   Begin VB.Menu mnuRFP 
      Caption         =   "RFP"
      Visible         =   0   'False
      Begin VB.Menu mnuRFPRec 
         Caption         =   "Receiving"
      End
      Begin VB.Menu mnuRFPCancel 
         Caption         =   "Cancel RFP"
      End
   End
End
Attribute VB_Name = "FrmSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Dim rsSummary As New ADODB.Recordset
Dim rsAccts As New ADODB.Recordset
Dim ActRow As Integer, fRow As Integer, fCol As Integer, TCol As Integer, AccID As Integer, ColPrn As Integer
Dim Qry As String, LoadPriceBCost As Boolean
Public mySrc As String, pRECTYPE As String
Dim rsUPO As New ADODB.Recordset
Dim MRC1Connect As Boolean, MRC2Connect As Boolean, UlasConnect As Boolean, TorilConnect As Boolean, Van2Connect As Boolean
Dim bCol As Integer, aRow As Integer
'******** CODE UPDATED 03/20/2018 **********
Dim CodeDateRange As Integer
'**** END CODE UPDATED 03/20/2018 **********
Dim tmpRs As New ADODB.Recordset

Private Sub btnTCancel_Click()
    ssTabTRec.Visible = False
End Sub

Private Sub btnTOk_Click()
    Dim rsTRFPUpdate As New ADODB.Recordset
    Dim ans As Integer, myChckDate As String
        
    ans = MsgBox("Do you want to Confirm this RFP?", vbQuestion + vbYesNo, "Validation")
    If ans = vbNo Then Exit Sub
    
    If txtTChckDate.Text = "__/__/____" Then
        myChckDate = ""
    Else
        myChckDate = Format(txtTChckDate.Text, "YYYY-MM-DD")
    End If
    
    If flxSum.TextMatrix(ActRow, 3) = "FINAL" Or flxSum.TextMatrix(ActRow, 3) = "FOR APPROVAL" Or flxSum.TextMatrix(ActRow, 3) = "APPROVED WITH CHECK NO." Then
        Qry = "call sp_acctpayinvcng.sp_rfp_treasury_update('" & txtTRFPNO.Text & "','" & Format(txtTRecDate.Text, "YYYY-MM-DD") & "','" & txtTBankAcct.Text & "','" & txtTChckNo.Text & "','" & myChckDate & "','" & txtTCVNo.Text & "'," & Val(UserID) & ",'" & Format(txtTDateDeposit.Text, "YYYY-MM-DD") & "','" & txtTRecBy.Text & "') "
        
        With rsTRFPUpdate
            If .State = adStateOpen Then .Close
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                MsgBox "" & !message & "", vbInformation + vbOKOnly, "" & !level & ""
            End If
        End With
    Else
        MsgBox "Request For Payment RFP# " & flxSum.TextMatrix(ActRow, 2) & " is not yet Final/Approved !", vbInformation + vbOKOnly, "Validation"
        Exit Sub
    End If
    
    ssTabTRec.Visible = False
    Call cmdRefresh_Click
End Sub

Private Sub cmdCancel_Click()
    SSConfirmTab.Visible = False
End Sub

Private Sub cmdOK_Click()
    Dim rsTransfer As New ADODB.Recordset
    Dim ans As Integer, i As Integer, myDateFinal As String, STFID As Long
    Dim Qry As String, CanRemarks As String
        
    ans = MsgBox("Do you want to Confirm ths Transfer?", vbQuestion + vbYesNo, "Validation")
    If ans = vbNo Then Exit Sub
    
    If flxSum.TextMatrix(ActRow, 2) <> "" Then
        MyRef = "TR#" & Val(flxSum.TextMatrix(ActRow, 4))
        'Qry = "SELECT SPITR.SPAREPARTID, SPITR.OUT From SPITR WHERE (((SPITR.REFERENCE)='" & MyRef & "'));"
        'mySQL
        Qry = "SELECT stf.TRANSFERID,stf.DELIVEREDDATE From stocktransfer as stf WHERE stf.TRANSFERID='" & Val(flxSum.TextMatrix(ActRow, 4)) & "';"
        
        With rsTransfer
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    STFID = !TRANSFERID
                    If flxSum.TextMatrix(ActRow, 3) = "" Then
                        LYRcon.Execute "UPDATE STOCKTRANSFER SET STOCKTRANSFER.DELIVEREDDATE = '" & Format(txtConfirmDate.Text, "yyyy-mm-dd") & "', STOCKTRANSFER.UPDATEDBY = " & UserID & ", STOCKTRANSFER.LASTUPDATE = '" & Format(Now, "yyyy-mm-dd") & "' " & _
                            "WHERE STOCKTRANSFER.TRANSFERID=" & STFID & ";"
                    Else
                        MsgBox "Spareparts Transfer Ref.# " & flxSum.TextMatrix(ActRow, 4) & " is already Confirm !", vbInformation + vbOKOnly, "CONFIRM"
                        .Close
                        SSConfirmTab.Visible = False
                        Exit Sub
                     End If
                    .MoveNext
                Loop
            End If
            .Close
        End With
    Else
        MsgBox "Spareparts Transfer Ref.# " & flxSum.TextMatrix(ActRow, 4) & " is not yet Final !", vbInformation + vbOKOnly, "CONFIRM"
        Exit Sub
    End If
    
    MsgBox "Spareparts Transfer Ref.# " & flxSum.TextMatrix(ActRow, 4) & " was Confirm!", vbInformation + vbOKOnly, "CONFIRM"
    SSConfirmTab.Visible = False
    Call cmdRefresh_Click
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo LocalErr
    Dim rsBranch As New Recordset
    Dim branchID As String
    Dim rsBranchPO As New ADODB.Recordset
    Dim Qry As String
    Dim MxMdl As Integer, i As Integer, j As Integer, InvRow As Integer, PbV As Integer, x As Integer
    Dim BlnkInv As Boolean

    With rsBranch
        Qry = "call sp_mrcwarehouse.sp_POConfirm_UpdatePO()"
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            flxBranch.row = 1
            flxBranch.col = 1
            If .RecordCount > 0 Then
                MsgBox " '" & !message & "' ", vbInformation, "'" & !level & "'"
            End If
        .Close
    End With
    
    Call RefreshRecord(1)
    pBar.Visible = False
    Exit Sub
    
LocalErr:
    MsgBox Err.Description
End Sub

Private Sub dFrom_CloseUp()
    FrDate = dFrom.Value
    If cmdNext.Tag = "1" Then
        Call RefreshRecord(1)
    End If
End Sub

Private Sub dTo_CloseUp()
    ToDate = dTo.Value
    If dTo.Value < dFrom.Value Then
        MsgBox "Please choose a date not less than " & dFrom.Value, vbExclamation + vbOKOnly, "Validation Error"
        dTo.Value = dFrom.Value
    End If
End Sub

Private Sub flxPBranch_EnterCell()
    Dim ActRow As Integer
    ActRow = flxBranch.row
    txtSPOID.Text = flxBranch.TextMatrix(ActRow, 3)
End Sub

Private Sub flxBranch_EnterCell()
    aRow = flxBranch.row
End Sub

'Public Sub UpdatePO()
'    Dim rsBranchPO As New ADODB.Recordset
'    Dim Qry As String, MxMdl As Integer, i As Integer, j As Integer, InvRow As Integer, PbV As Integer
'    Dim BlnkInv As Boolean
'
'    pBar.Value = 0
'    pBar.Max = 100
'    txtBID.Text = ""
'    txtSPOID.Text = ""
'    txtbName.Text = ""
'
'    On Error GoTo LocalErr
'    Call ConnectOnline(15) ' MCP DEPOT
'    With rsBranchPO
'        txtBID.Text = 15
'            If txtBID.Text = 15 Then
'                With flxBranch
'                    For i = fRow To .Rows - 1
'                        For j = fCol To bCol - 3
'                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
'                                .Col = 1: .Row = i
'                                .ColSel = bCol - 1: .RowSel = i
'                                .TopRow = i
'                                fRow = IIf(j = bCol - 1, i + 1, i)
'                                fCol = IIf(j = bCol - 1, 1, j + 1)
'                                aRow = i
'                            End If
'                        Next j
'                        fCol = 1
'                    Next i
'                    txtSPOID.Text = .TextMatrix(aRow, 3)
'                    txtbName.Text = .TextMatrix(aRow, 2)
'                End With
'                fRow = 1
'                fCol = 1
'            End If
'
'            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
'                  "From sporder Where sporder.SPORDERID > " & txtSPOID.Text & ";"
'            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
'            If .RecordCount > 0 Then
'                .MoveFirst
'                flxOrder.Rows = 1
'                MxMdl = .RecordCount
'                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
'                pBar.Visible = True
'                Do While Not .EOF
'                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
'                    .MoveNext
'                    pBar.Value = pBar.Value + PbV
'                    DoEvents
'                Loop
'                With flxOrder
'                    For i = 1 To .Rows - 1
'                        LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
'                                "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
'                                "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
'                    Next i
'                End With
'            Else
'                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
'                If pBar.Value > 0 Then
'                    flxOrder.Rows = 1
'                    flxOrder.Rows = 2
'                    pBar.Value = 1000
'                    DoEvents
'                Else
'                    pBar.Value = 0
'                End If
'            End If
'        .Close
'    End With
'    txtBID.Text = ""
'    txtSPOID.Text = ""
'    txtbName.Text = ""
'    OnlineCon.Close
''
''
''    '-----------------------------------------------------------------------
''
''    pBar.Value = 0
''    pBar.Max = 100
''    Call ConnectOnline(13) ' MRC ULAS
''    With rsBranchPO
''        txtBID.Text = 13
''            If txtBID.Text = 13 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                        fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                  "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                    For i = 1 To .Rows - 1
''                        LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                    Next i
''                End With
''            Else
''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''                If pBar.Value > 0 Then
''                    flxOrder.Rows = 1
''                    flxOrder.Rows = 2
''                    DoEvents
''                Else
''                    pBar.Value = 0
''                End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
'''-----------------------------------------------------------------------------------------------------------------------
''    Call ConnectOnline(12) 'MRC1 SOBRECAREY
''    With rsBranchPO
''        txtBID.Text = 12
''            If txtBID.Text = 12 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                        fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                    For i = 1 To .Rows - 1
''                        LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                    Next i
''                End With
''            Else
''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''                If pBar.Value > 0 Then
''                    flxOrder.Rows = 1
''                    flxOrder.Rows = 2
''                    DoEvents
''                Else
''                    pBar.Value = 0
''                End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
'''------------------------------------------------------------------------------------------------------------------
''
''    Call ConnectOnline(16) 'van 2 [LHC403]
''    With rsBranchPO
''        txtBID.Text = 16
''            If txtBID.Text = 16 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                            fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                        For i = 1 To .Rows - 1
''                            LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                    "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                    "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                        Next i
''                End With
''            Else
''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''                If pBar.Value > 0 Then
''                    flxOrder.Rows = 1
''                    flxOrder.Rows = 2
''                    DoEvents
''                Else
''                    pBar.Value = 0
''                End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
'''------------------------------------------------------------------------------------------------------------------
''
''    Call ConnectOnline(22) 'van 56 [AAH1233]
''    With rsBranchPO
''        txtBID.Text = 22
''            If txtBID.Text = 22 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                        fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                    For i = 1 To .Rows - 1
''                        LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                    Next i
''                End With
''            Else
''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''                If pBar.Value > 0 Then
''                    flxOrder.Rows = 1
''                    flxOrder.Rows = 2
''                    DoEvents
''                Else
''                    pBar.Value = 0
''                End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
'''-------------------------------------------------------------------------------------------------------------------------------
''
''    Call ConnectOnline(27) 'van 55 [AAH1250]
''    With rsBranchPO
''        txtBID.Text = 27
''            If txtBID.Text = 27 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                        fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                    For i = 1 To .Rows - 1
''                        LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                    Next i
''                End With
''            Else
''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''                If pBar.Value > 0 Then
''                    flxOrder.Rows = 1
''                    flxOrder.Rows = 2
''                    DoEvents
''                Else
''                    pBar.Value = 0
''                End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
''
'''-----------------------------------------------------------------------------------------------------------------------------
''
''    Call ConnectOnline(28) 'MRC 2 RETAIL
''    With rsBranchPO
''        txtBID.Text = 28
''            If txtBID.Text = 28 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                        fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                    For i = 1 To .Rows - 1
''                        LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                    Next i
''                End With
''            Else
''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''                If pBar.Value > 0 Then
''                    flxOrder.Rows = 1
''                    flxOrder.Rows = 2
''                    DoEvents
''                Else
''                    pBar.Value = 0
''                End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
'''-----------------------------------------------------------------------------------------
''    Call ConnectOnline(21) 'MRC 2 WHOLESALE
''    With rsBranchPO
''        txtBID.Text = 21
''            If txtBID.Text = 21 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                        fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                        For i = 1 To .Rows - 1
''                            LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                    "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                    "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                        Next i
''                End With
''            Else
''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''                If pBar.Value > 0 Then
''                    flxOrder.Rows = 1
''                    flxOrder.Rows = 2
''                    DoEvents
''                Else
''                    pBar.Value = 0
''                End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
'''-------------------------------------------------------------------------
''    Call ConnectOnline(19) 'MRC PANABO
''    With rsBranchPO
''        txtBID.Text = 19
''            If txtBID.Text = 19 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                        fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                    For i = 1 To .Rows - 1
''                        LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                    Next i
''                End With
''            Else
''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''                If pBar.Value > 0 Then
''                    flxOrder.Rows = 1
''                    flxOrder.Rows = 2
''                    DoEvents
''                Else
''                    pBar.Value = 0
''                End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
''
'''-------------------------------------------------------------------------
''    Call ConnectOnline(20) 'HONDA POLOMOLOK
''    With rsBranchPO
''        txtBID.Text = 20
''            If txtBID.Text = 20 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                        fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                    For i = 1 To .Rows - 1
''                        LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                    Next i
''                End With
''            Else
''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''                If pBar.Value > 0 Then
''                    flxOrder.Rows = 1
''                    flxOrder.Rows = 2
''                    DoEvents
''                Else
''                    pBar.Value = 0
''                End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
'''-------------------------------------------------------------------------
''    Call ConnectOnline(26) 'MOTORHUB MARBEL
''    With rsBranchPO
''        txtBID.Text = 26
''            If txtBID.Text = 26 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                        fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                        For i = 1 To .Rows - 1
''                            LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                    "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                    "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                        Next i
''                End With
''            Else
''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''                If pBar.Value > 0 Then
''                    flxOrder.Rows = 1
''                    flxOrder.Rows = 2
''                    DoEvents
''                Else
''                    pBar.Value = 0
''                End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
'''-------------------------------------------------------------------------
''    Call ConnectOnline(23) 'MRC GENSAN
''    With rsBranchPO
''        txtBID.Text = 23
''            If txtBID.Text = 23 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                        fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                        For i = 1 To .Rows - 1
''                            LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                    "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                    "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                        Next i
''                End With
''            Else
''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''                If pBar.Value > 0 Then
''                    flxOrder.Rows = 1
''                    flxOrder.Rows = 2
''                    DoEvents
''                Else
''                    pBar.Value = 0
''                End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
'''-------------------------------------------------------------------------
''    Call ConnectOnline(24) 'HONDA STOTOMAS
''    With rsBranchPO
''        txtBID.Text = 24
''            If txtBID.Text = 24 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                        fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                    For i = 1 To .Rows - 1
''                        LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                    Next i
''                End With
''            Else
''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''                If pBar.Value > 0 Then
''                    flxOrder.Rows = 1
''                    flxOrder.Rows = 2
''                    DoEvents
''                Else
''                    pBar.Value = 0
''                End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
'''-------------------------------------------------------------------------
''    Call ConnectOnline(25) 'HONDA SAMAL
''    With rsBranchPO
''        txtBID.Text = 25
''            If txtBID.Text = 25 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                        fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                    For i = 1 To .Rows - 1
''                        LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                    Next i
''                End With
''            Else
''               MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''               If pBar.Value > 0 Then
''                   flxOrder.Rows = 1
''                   flxOrder.Rows = 2
''                   DoEvents
''               Else
''                   pBar.Value = 0
''               End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
'''-------------------------------------------------------------------------
''    Call ConnectOnline(14) 'SPAREPARTS MOTORHUB TAGUM
''    With rsBranchPO
''        txtBID.Text = 14
''            If txtBID.Text = 14 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                        fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                    For i = 1 To .Rows - 1
''                        LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                    Next i
''                End With
''            Else
''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''                If pBar.Value > 0 Then
''                    flxOrder.Rows = 1
''                    flxOrder.Rows = 2
''                    DoEvents
''                Else
''                    pBar.Value = 0
''                End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
'''-------------------------------------------------------------------------
''    Call ConnectOnline(32) 'MC PARTS DEPOT RIZAL TAGUM
''    With rsBranchPO
''        txtBID.Text = 32
''            If txtBID.Text = 32 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                        fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                    For i = 1 To .Rows - 1
''                        LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                    Next i
''                End With
''            Else
''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''                If pBar.Value > 0 Then
''                    flxOrder.Rows = 1
''                    flxOrder.Rows = 2
''                    DoEvents
''                Else
''                    pBar.Value = 0
''                End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
''
'''-------------------------------------------------------------------------
''    Call ConnectOnline(33) 'Binos  SpareParts TAGUM
''    With rsBranchPO
''        txtBID.Text = 33
''            If txtBID.Text = 33 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                        fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                    For i = 1 To .Rows - 1
''                        LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                    Next i
''                End With
''            Else
''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''                If pBar.Value > 0 Then
''                    flxOrder.Rows = 1
''                    flxOrder.Rows = 2
''                    DoEvents
''                Else
''                    pBar.Value = 0
''                End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
'''-------------------------------------------------------------------------
''    Call ConnectOnline(36) 'MRCVAN5 - Van74[zf9790]
''    With rsBranchPO
''        txtBID.Text = 36
''            If txtBID.Text = 36 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                        fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                    For i = 1 To .Rows - 1
''                        LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                    Next i
''                End With
''            Else
''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''                If pBar.Value > 0 Then
''                    flxOrder.Rows = 1
''                    flxOrder.Rows = 2
''                    DoEvents
''                Else
''                    pBar.Value = 0
''                End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
'''-------------------------------------------------------------
''    Call ConnectOnline(31) 'MRCVAN4 - Van73[ZF9619]
''    With rsBranchPO
''        txtBID.Text = 31
''            If txtBID.Text = 31 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                        fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                    For i = 1 To .Rows - 1
''                        LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                    Next i
''                End With
''            Else
''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''                If pBar.Value > 0 Then
''                    flxOrder.Rows = 1
''                    flxOrder.Rows = 2
''                    DoEvents
''                Else
''                    pBar.Value = 0
''                End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
'''-------------------------------------------------------------
''    Call ConnectOnline(34) 'MRCMATINA
''    With rsBranchPO
''        txtBID.Text = 34
''            If txtBID.Text = 34 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                        fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                    For i = 1 To .Rows - 1
''                        LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                    Next i
''                End With
''            Else
''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''                If pBar.Value > 0 Then
''                    flxOrder.Rows = 1
''                    flxOrder.Rows = 2
''                    DoEvents
''                Else
''                    pBar.Value = 0
''                End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
'''-------------------------------------------------------------------------
''    Call ConnectOnline(30) 'MRC CABANTIAN
''    With rsBranchPO
''        txtBID.Text = 30
''            If txtBID.Text = 30 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                        fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                    For i = 1 To .Rows - 1
''                        LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                    Next i
''                End With
''            Else
''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''                    If pBar.Value > 0 Then
''                        flxOrder.Rows = 1
''                        flxOrder.Rows = 2
''                        DoEvents
''                    Else
''                        pBar.Value = 0
''                    End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
'''-------------------------------------------------------------------------
''    Call ConnectOnline(43) 'MRC KAPALONG
''    With rsBranchPO
''        txtBID.Text = 43
''            If txtBID.Text = 43 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                        fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                    For i = 1 To .Rows - 1
''                        LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                    Next i
''                End With
''            Else
''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''                If pBar.Value > 0 Then
''                    flxOrder.Rows = 1
''                    flxOrder.Rows = 2
''                    DoEvents
''                Else
''                    pBar.Value = 0
''                End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
'''-------------------------------------------------------------------------
''    Call ConnectOnline(42) 'MRC MANKILAM
''    With rsBranchPO
''        txtBID.Text = 42
''            If txtBID.Text = 42 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                        fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                    For i = 1 To .Rows - 1
''                        LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                    Next i
''                End With
''            Else
''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''                If pBar.Value > 0 Then
''                    flxOrder.Rows = 1
''                    flxOrder.Rows = 2
''                    DoEvents
''                Else
''                    pBar.Value = 0
''                End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
'''-------------------------------------------------------------------------
''    Call ConnectOnline(37) 'MRC MOTORHUBULAS
''    With rsBranchPO
''        txtBID.Text = 37
''            If txtBID.Text = 37 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                        fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                    For i = 1 To .Rows - 1
''                        LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                    Next i
''                End With
''            Else
''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''                If pBar.Value > 0 Then
''                    flxOrder.Rows = 1
''                    flxOrder.Rows = 2
''                    DoEvents
''                Else
''                    pBar.Value = 0
''                End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
'''-------------------------------------------------------------------------
''    Call ConnectOnline(46) 'MRC STO TOMAS
''    With rsBranchPO
''        txtBID.Text = 46
''            If txtBID.Text = 46 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                        fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                    For i = 1 To .Rows - 1
''                        LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                    Next i
''                End With
''            Else
''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''                If pBar.Value > 0 Then
''                    flxOrder.Rows = 1
''                    flxOrder.Rows = 2
''                    DoEvents
''                Else
''                    pBar.Value = 0
''                End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
'''-------------------------------------------------------------------------
''    Call ConnectOnline(39) 'MRC PIAGGIO TAGUM
''    With rsBranchPO
''        txtBID.Text = 39
''            If txtBID.Text = 39 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                        fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                    For i = 1 To .Rows - 1
''                        LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                    Next i
''                End With
''            Else
''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''                If pBar.Value > 0 Then
''                    flxOrder.Rows = 1
''                    flxOrder.Rows = 2
''                    DoEvents
''                Else
''                    pBar.Value = 0
''                End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
'''-------------------------------------------------------------------------
''    Call ConnectOnline(45) 'MRC APOKON
''    With rsBranchPO
''        txtBID.Text = 45
''            If txtBID.Text = 45 Then
''                With flxBranch
''                    For i = fRow To .Rows - 1
''                        For j = fCol To bCol - 3
''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                .Col = 1: .Row = i
''                                .ColSel = bCol - 1: .RowSel = i
''                                .TopRow = i
''                                fRow = IIf(j = bCol - 1, i + 1, i)
''                                fCol = IIf(j = bCol - 1, 1, j + 1)
''                                aRow = i
''                            End If
''                        Next j
''                        fCol = 1
''                    Next i
''                    txtSPOID.Text = .TextMatrix(aRow, 3)
''                    txtbName.Text = .TextMatrix(aRow, 2)
''                End With
''                fRow = 1
''                fCol = 1
''            End If
''
''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
''
''            If .RecordCount > 0 Then
''                pBar.Value = 0
''                pBar.Max = 100
''                .MoveFirst
''                flxOrder.Rows = 1
''                MxMdl = .RecordCount
''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
''                pBar.Visible = True
''                Do While Not .EOF
''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
''                    .MoveNext
''                    pBar.Value = pBar.Value + PbV
''                    DoEvents
''                Loop
''
''                With flxOrder
''                    For i = 1 To .Rows - 1
''                        LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
''                                "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
''                                "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
''                    Next i
''                End With
''            Else
''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
''                If pBar.Value > 0 Then
''                    flxOrder.Rows = 1
''                    flxOrder.Rows = 2
''                    DoEvents
''                Else
''                    pBar.Value = 0
''                End If
''            End If
''        .Close
''    End With
''    txtBID.Text = ""
''    txtSPOID.Text = ""
''    txtbName.Text = ""
''    OnlineCon.Close
''
''''-------------------------------------------------------------------------
'''    Call ConnectOnline(45) 'MRC MONKAYO
'''    With rsBranchPO
'''        txtBID.Text = 45
'''            If txtBID.Text = 45 Then
'''                With flxBranch
'''                    For i = fRow To .Rows - 1
'''                        For j = fCol To bCol - 3
'''                            If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
'''                                .Col = 1: .Row = i
'''                                .ColSel = bCol - 1: .RowSel = i
'''                                .TopRow = i
'''                                fRow = IIf(j = bCol - 1, i + 1, i)
'''                                fCol = IIf(j = bCol - 1, 1, j + 1)
'''                                aRow = i
'''                            End If
'''                        Next j
'''                        fCol = 1
'''                    Next i
'''                    txtSPOID.Text = .TextMatrix(aRow, 3)
'''                    txtbName.Text = .TextMatrix(aRow, 2)
'''                End With
'''                fRow = 1
'''                fCol = 1
'''            End If
'''
'''            Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID " & _
'''                "From sporder Where sporder.SPORDERID > " & Val(txtSPOID.Text) & ";"
'''            .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
'''
'''            If .RecordCount > 0 Then
'''                pBar.Value = 0
'''                pBar.Max = 100
'''                .MoveFirst
'''                flxOrder.Rows = 1
'''                MxMdl = .RecordCount
'''                If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
'''                pBar.Visible = True
'''                Do While Not .EOF
'''                    flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID
'''                    .MoveNext
'''                    pBar.Value = pBar.Value + PbV
'''                    DoEvents
'''                Loop
'''
'''                With flxOrder
'''                    For i = 1 To .Rows - 1
'''                        LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID)" & _
'''                                "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
'''                                "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ");"
'''                    Next i
'''                End With
'''            Else
'''                MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
'''                If pBar.Value > 0 Then
'''                    flxOrder.Rows = 1
'''                    flxOrder.Rows = 2
'''                    DoEvents
'''                Else
'''                    pBar.Value = 0
'''                End If
'''            End If
'''        .Close
'''    End With
'''    txtBID.Text = ""
'''    txtSPOID.Text = ""
'''    txtbName.Text = ""
'''    OnlineCon.Close
'
'    Call RefreshRecord(1)
'    pBar.Visible = False
'    Exit Sub
'LocalErr:
'    MsgBox Err.Description
' End Sub

Public Sub savePOdetails(MyBranchID As String)
    Dim rsPOdetails As New ADODB.Recordset
    Dim rsOnlineSave As New ADODB.Recordset
    Dim rsChkExists As New ADODB.Recordset
    Dim chkExistsRefNo As String, chkBCODE As String
    Dim Qry As String, MxMdl As Integer, i As Integer, j As Integer, InvRow As Integer, PbV As Integer

    Qry = "Select BRANCHCODE FROM BRANCH WHERE BRANCHID = " & MyBranchID & " "
    chkBCODE = GetFieldValue(Qry)
    'PAG SAVE SA PO DETAILS SA BRANCH UG E UPDATE ANG STATUS FROM NEW TO VIEWED
    Qry = "call sp_mrcwarehouse.sp_POConfirm_savePODetails('" & chkBCODE & "','" & flxSum.TextMatrix(ActRow, 2) & "'," & flxSum.TextMatrix(ActRow, 12) & ")"
    With rsPOdetails
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            If !level = "Error" Then
                MsgBox " & !MESSAGE & ", vbInformation + vbOKOnly, " & !LEVEL & "
                Exit Sub
            End If
        End If
        .Close
    End With
    
                   
End Sub

Public Sub ShowFilteredRecords()
    If mySrc = "Detailed Outgoing Items" Then
        If dCode = 1 Then
            'Qry = "SELECT productline.PRODINIT, CONCAT(productline.BRAND, ' ', productline.DESCRIPTION) AS MYDESC, productline.MODEL, items.COLOR, items.SERIALNO, invtreg.REFERENCE, DATE_FORMAT(invtreg.TRANSDATE,'%c/%d/%Y'), invtreg.OUT, invtreg.SOURCE, invtreg.REMARKS, productline.CLASSIFICATION " & _
                  "From invtreg INNER JOIN items ON (invtreg.ITEMID = items.ITEMID) INNER JOIN productline ON (items.PRODID = productline.PRODID) " & _
                  txtCriteria.Text & " AND (invtreg.TRANSTYPE = 'TRANSFER' OR invtreg.TRANSTYPE = 'SAL') AND invtreg.TRANSDATE = '" & Format(dFrom.Value, "YYYY-MM-DD") & "' Order By productline.PRODINIT, productline.BRAND, productline.DESCRIPTION, invtreg.REMARKS, invtreg.TRANSDATE;"
            'Qry = "SELECT productline.CLASSIFICATION, productline.BRAND, productline.MODEL, IF(productline.PRODINIT = 'MC' OR productline.PRODINIT = 'FR', '', productline.DESCRIPTION) AS MYDESC, items.COLOR, items.SERIALNO, invtreg.OUT, invtreg.REFERENCE, DATE_FORMAT(invtreg.TRANSDATE,'%c/%d/%Y'), invtreg.SOURCE, invtreg.REMARKS, productline.PRODINIT, items.GRPCODE " & _
                  "From invtreg INNER JOIN items ON (invtreg.ITEMID = items.ITEMID) INNER JOIN productline ON (items.PRODID = productline.PRODID) " & _
                  txtCriteria.Text & " AND (invtreg.TRANSTYPE = 'TRANSFER' OR invtreg.TRANSTYPE = 'SAL') AND invtreg.TRANSDATE = '" & Format(dFrom.Value, "YYYY-MM-DD") & "' Order By items.GRPCODE, productline.PRODINIT, productline.CLASSIFICATION, productline.BRAND, productline.MODEL, items.COLOR, invtreg.TRANSDATE;"
            Qry = "SELECT productline.PRODCODE, productline.MODEL, IF(productline.PRODINIT = 'MC' OR productline.PRODINIT = 'FR', '', productline.DESCRIPTION) AS MYDESC, items.COLOR, items.SERIALNO, invtreg.OUT, invtreg.REFERENCE, DATE_FORMAT(invtreg.TRANSDATE, '%c/%d/%Y') AS FIELD_1, invtreg.SOURCE, invtreg.SRCTYPE, invtreg.REMARKS, productline.PRODINIT, productline.CLASSIFICATION, productline.BRAND " & _
                  "From invtreg INNER JOIN items ON (invtreg.ITEMID = items.ITEMID) INNER JOIN productline ON (items.PRODID = productline.PRODID) " & _
                  txtCriteria.Text & " AND (invtreg.TRANSTYPE = 'TRANSFER' OR invtreg.TRANSTYPE = 'SAL' OR invtreg.TRANSTYPE = 'RETURN' OR invtreg.TRANSTYPE = 'SERVICE') AND invtreg.TRANSDATE = '" & Format(dFrom.Value, "YYYY-MM-DD") & "' Order By PRODINIT, productline.CLASSIFICATION, BRAND, productline.MODEL, invtreg.TRANSDATE, invtreg.REFERENCE"
        Else
            'Qry = "SELECT productline.PRODINIT, CONCAT(productline.BRAND, ' ', productline.DESCRIPTION) AS MYDESC, productline.MODEL, items.COLOR, items.SERIALNO, invtreg.REFERENCE, DATE_FORMAT(invtreg.TRANSDATE,'%c/%d/%Y'), invtreg.OUT, invtreg.SOURCE, invtreg.REMARKS, productline.CLASSIFICATION " & _
                  "From invtreg INNER JOIN items ON (invtreg.ITEMID = items.ITEMID) INNER JOIN productline ON (items.PRODID = productline.PRODID) " & _
                  txtCriteria.Text & " AND (invtreg.TRANSTYPE = 'TRANSFER' OR invtreg.TRANSTYPE = 'SAL') AND (invtreg.TRANSDATE >= '" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND invtreg.TRANSDATE <= '" & Format(dTo.Value, "YYYY-MM-DD") & "') Order By productline.PRODINIT, productline.BRAND, productline.DESCRIPTION, invtreg.REMARKS, invtreg.TRANSDATE;"
            'Qry = "SELECT productline.CLASSIFICATION, productline.BRAND, productline.MODEL, IF(productline.PRODINIT = 'MC' OR productline.PRODINIT = 'FR', '', productline.DESCRIPTION) AS MYDESC, items.COLOR, items.SERIALNO, invtreg.OUT, invtreg.REFERENCE, DATE_FORMAT(invtreg.TRANSDATE,'%c/%d/%Y'), invtreg.SOURCE, invtreg.REMARKS, productline.PRODINIT, items.GRPCODE " & _
                  "From invtreg INNER JOIN items ON (invtreg.ITEMID = items.ITEMID) INNER JOIN productline ON (items.PRODID = productline.PRODID) " & _
                  txtCriteria.Text & " AND (invtreg.TRANSTYPE = 'TRANSFER' OR invtreg.TRANSTYPE = 'SAL') AND invtreg.TRANSDATE >= '" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND invtreg.TRANSDATE <= '" & Format(dTo.Value, "YYYY-MM-DD") & "' Order By items.GRPCODE, productline.PRODINIT, productline.CLASSIFICATION, productline.BRAND, productline.MODEL, items.COLOR, invtreg.TRANSDATE;"
           Qry = "SELECT productline.PRODCODE, productline.MODEL, IF(productline.PRODINIT = 'MC' OR productline.PRODINIT = 'FR', '', productline.DESCRIPTION) AS MYDESC, items.COLOR, items.SERIALNO, invtreg.OUT, invtreg.REFERENCE, DATE_FORMAT(invtreg.TRANSDATE, '%c/%d/%Y') AS FIELD_1, invtreg.SOURCE, invtreg.SRCTYPE, invtreg.REMARKS, productline.PRODINIT, productline.CLASSIFICATION, productline.BRAND " & _
                  "From invtreg INNER JOIN items ON (invtreg.ITEMID = items.ITEMID) INNER JOIN productline ON (items.PRODID = productline.PRODID) " & _
                  txtCriteria.Text & " AND (invtreg.TRANSTYPE = 'TRANSFER' OR invtreg.TRANSTYPE = 'SAL' OR invtreg.TRANSTYPE = 'RETURN' OR invtreg.TRANSTYPE = 'SERVICE') AND invtreg.TRANSDATE >= '" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND invtreg.TRANSDATE <= '" & Format(dTo.Value, "YYYY-MM-DD") & "' Order By PRODINIT, productline.CLASSIFICATION, BRAND, productline.MODEL, invtreg.TRANSDATE, invtreg.REFERENCE;"
        End If
    ElseIf mySrc = "Detailed Receiving Receipt" Then
        If dCode = 1 Then
            'Qry = "SELECT productline.CLASSIFICATION, productline.BRAND, productline.MODEL, IF(productline.PRODINIT = 'MC' OR productline.PRODINIT = 'FR', '', productline.DESCRIPTION) AS MYDESC, items.COLOR, items.SERIALNO, invtreg.REC, invtreg.REFERENCE, DATE_FORMAT(invtreg.TRANSDATE,'%c/%d/%Y'), invtreg.SOURCE, invtreg.REMARKS, productline.PRODINIT, items.GRPCODE " & _
                  "From invtreg INNER JOIN items ON (invtreg.ITEMID = items.ITEMID) INNER JOIN productline ON (items.PRODID = productline.PRODID) " & _
                  txtCriteria.Text & " AND (invtreg.TRANSTYPE = 'REC' OR invtreg.TRANSTYPE = 'DEPOSIT') AND invtreg.TRANSDATE = '" & Format(dFrom.Value, "YYYY-MM-DD") & "' Order By items.GRPCODE, productline.PRODINIT, productline.CLASSIFICATION, productline.BRAND, productline.MODEL, items.COLOR, invtreg.TRANSDATE;"
            Qry = "SELECT productline.PRODCODE, productline.MODEL, IF(productline.PRODINIT = 'MC' OR productline.PRODINIT = 'FR', '', productline.DESCRIPTION) AS MYDESC, items.COLOR, items.SERIALNO, invtreg.REC, invtreg.REFERENCE, DATE_FORMAT(invtreg.TRANSDATE, '%c/%d/%Y') AS FIELD_1, invtreg.SOURCE, invtreg.SRCTYPE, invtreg.REMARKS, productline.PRODINIT, productline.CLASSIFICATION, productline.BRAND " & _
                  "From invtreg INNER JOIN items ON (invtreg.ITEMID = items.ITEMID) INNER JOIN productline ON (items.PRODID = productline.PRODID) " & _
                   txtCriteria.Text & " AND invtreg.TRANSTYPE = 'REC' AND invtreg.TRANSDATE = '" & Format(dFrom.Value, "YYYY-MM-DD") & "' Order By PRODINIT, productline.CLASSIFICATION, BRAND, productline.MODEL, invtreg.TRANSDATE"
        Else
            'Qry = "SELECT productline.CLASSIFICATION, productline.BRAND, productline.MODEL, IF(productline.PRODINIT = 'MC' OR productline.PRODINIT = 'FR', '', productline.DESCRIPTION) AS MYDESC, items.COLOR, items.SERIALNO, invtreg.REC, invtreg.REFERENCE, DATE_FORMAT(invtreg.TRANSDATE,'%c/%d/%Y'), invtreg.SOURCE, invtreg.REMARKS, productline.PRODINIT, items.GRPCODE " & _
                  "From invtreg INNER JOIN items ON (invtreg.ITEMID = items.ITEMID) INNER JOIN productline ON (items.PRODID = productline.PRODID) " & _
                  txtCriteria.Text & " AND (invtreg.TRANSTYPE = 'REC' OR invtreg.TRANSTYPE = 'DEPOSIT') AND invtreg.TRANSDATE >= '" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND invtreg.TRANSDATE <= '" & Format(dTo.Value, "YYYY-MM-DD") & "' Order By items.GRPCODE, productline.PRODINIT, productline.CLASSIFICATION, productline.BRAND, productline.MODEL, items.COLOR, invtreg.TRANSDATE, invtreg.REFERENCE;"
             Qry = "SELECT productline.PRODCODE, productline.MODEL, IF(productline.PRODINIT = 'MC' OR productline.PRODINIT = 'FR', '', productline.DESCRIPTION) AS MYDESC, items.COLOR, items.SERIALNO, invtreg.REC, invtreg.REFERENCE, DATE_FORMAT(invtreg.TRANSDATE, '%c/%d/%Y') AS FIELD_1, invtreg.SOURCE, invtreg.SRCTYPE, invtreg.REMARKS, productline.PRODINIT, productline.CLASSIFICATION, productline.BRAND " & _
                  "From invtreg INNER JOIN items ON (invtreg.ITEMID = items.ITEMID) INNER JOIN productline ON (items.PRODID = productline.PRODID) " & _
                  txtCriteria.Text & " AND invtreg.TRANSTYPE = 'REC' AND invtreg.TRANSDATE >= '" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND invtreg.TRANSDATE <= '" & Format(dTo.Value, "YYYY-MM-DD") & "' Order By PRODINIT, productline.CLASSIFICATION, BRAND, productline.MODEL, invtreg.TRANSDATE"
        End If
    ElseIf mySrc = "Disbursement" Then
        If dCode = 1 Then
            Qry = "SELECT date_format(CVDATE,'%c/%d/%Y') AS CVDATE, CVNUMBER, DISBURSETYPE, ACCTDESC, PAYEE, PARTICULARS, format(AMOUNT,2) AS AMOUNT, REMARKS, PRINTCOPY, VOUCHERID, CUSTOMERID, AGENTID, DOCREF, DOCREF2 From disbursement " & _
                  txtCriteria.Text & " AND CVDATE ='" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND DISBURSETYPE NOT LIKE 'Due From%' AND DISBURSETYPE<>'Withdrawal/Personal' ORDER BY CVDATE ASC, CVNUMBER ASC;"    'AND DISBURSETYPE <> 'General Management Office'
        Else
            Qry = "SELECT date_format(CVDATE,'%c/%d/%Y') AS CVDATE, CVNUMBER, DISBURSETYPE, ACCTDESC, PAYEE, PARTICULARS, format(AMOUNT,2) AS AMOUNT, REMARKS, PRINTCOPY, VOUCHERID, CUSTOMERID, AGENTID, DOCREF, DOCREF2 From disbursement " & _
                  txtCriteria.Text & " AND CVDATE >='" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND CVDATE <='" & Format(dTo.Value, "YYYY-MM-DD") & "' AND DISBURSETYPE NOT LIKE 'Due From%' AND DISBURSETYPE<>'Withdrawal/Personal' ORDER BY CVDATE ASC, CVNUMBER ASC;"    'AND DISBURSETYPE <> 'General Management Office'
        End If
    
    End If
    
    If rsSummary.State = adStateOpen Then rsSummary.Close
    rsSummary.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsSummary.RecordCount = 0 Then
        flxSum.Clear
        flxSum.Rows = 2
        flxSum.Cols = TCol
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxSum.DataSource = rsSummary
    ActRow = 1
    flxSum.Visible = True
    Call RXResizer1_ResizeComplete
    
End Sub



Public Sub RefreshRecord(dCode As Integer)
    Dim i As Integer
    If mySrc = "SP Incoming" Then
        Qry = "call sp_mrcwarehouse.sp_incoming_summary(" & Val(dCode) & ",'" & Format(dFrom.Value, "yyyy-mm-dd") & "','" & Format(dTo.Value, "yyyy-mm-dd") & "','" & pRECTYPE & "')"
    ElseIf mySrc = "Advances for Liquidation" Then
        If dCode = 1 Then
            Qry = "SELECT date_format(advliquidation.AFLDATE,'%c/%d/%Y'), advliquidation.AFLNO, format(advliquidation.AMTREC,2), format(advliquidation.AMTRETURN,2), format(advliquidation.AFLAMOUNT,2), advliquidation.PREPAREDBY, advliquidation.REMARKS, advliquidation.PRINTCOPY, advliquidation.AFLID " & _
                  "From advliquidation Where advliquidation.AFLDATE = '" & Format(dFrom.Value, "YYYY-MM-DD") & "' Order By advliquidation.AFLNO;"
        Else
            Qry = "SELECT date_format(advliquidation.AFLDATE,'%c/%d/%Y'), advliquidation.AFLNO, format(advliquidation.AMTREC,2), format(advliquidation.AMTRETURN,2), format(advliquidation.AFLAMOUNT,2), advliquidation.PREPAREDBY, advliquidation.REMARKS, advliquidation.PRINTCOPY, advliquidation.AFLID " & _
                  "From advliquidation Where advliquidation.AFLDATE >= '" & Format(dFrom.Value, "YYYY-MM-DD") & "' and advliquidation.AFLDATE <= '" & Format(dTo.Value, "YYYY-MM-DD") & "' Order By advliquidation.AFLNO;"
        End If
        
    ElseIf mySrc = "Sales Return" Then
        If dCode = 1 Then '|^DATE RETURNED     |^SALES RET.#           |^INVOICE#|^INVOICE DATE      |<NAME OF CUSTOMER|^SALES RET.THIS MONTH|^SALES RET.-PRIOR MONTH SALES|>TOTAL|<REMARKS |^PRINTED"
             Qry = "SELECT SRETURNSP.RETDATE, SRETURNSP.RETNO, SRETURNSP.INVOICE, SPINVOICE.INVDATE, SPINVOICE.CHARGETO, IF(MONTH(SRETURNSP.RETDATE) = MONTH(SPINVOICE.INVDATE), SRETURNSP.TOTALAMT, '') AS THMOS, " & _
                    "IF(MONTH(SRETURNSP.RETDATE) = MONTH(SPINVOICE.INVDATE),'',SRETURNSP.TOTALAMT) AS PRMOS, SRETURNSP.TOTALAMT, SRETURNSP.REMARKS, SRETURNSP.PRINTCOPY, SRETURNSP.RETURNID " & _
                    "From SRETURNSP INNER JOIN SPINVOICE ON (SRETURNSP.INVOICE = SPINVOICE.SPINVNO) Where SRETURNSP.RETDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' " & _
                    "Order By SRETURNSP.RETDATE, SRETURNSP.RETNO Desc"
        Else
             Qry = "SELECT SRETURNSP.RETDATE, SRETURNSP.RETNO, SRETURNSP.INVOICE, SPINVOICE.INVDATE, SPINVOICE.CHARGETO, IF(MONTH(SRETURNSP.RETDATE) = MONTH(SPINVOICE.INVDATE), SRETURNSP.TOTALAMT, '') AS THMOS, " & _
                    "IF(MONTH(SRETURNSP.RETDATE) = MONTH(SPINVOICE.INVDATE),'',SRETURNSP.TOTALAMT) AS PRMOS, SRETURNSP.TOTALAMT, SRETURNSP.REMARKS, SRETURNSP.PRINTCOPY, SRETURNSP.RETURNID " & _
                    "From SRETURNSP INNER JOIN SPINVOICE ON (SRETURNSP.INVOICE = SPINVOICE.SPINVNO) Where SRETURNSP.RETDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' AND " & _
                    "SRETURNSP.RETDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' " & _
                    "Order By SRETURNSP.RETDATE, SRETURNSP.RETNO Desc"
        End If
    ElseIf mySrc = "Purchase Return" Then
        If dCode = 1 Then
            Qry = "SELECT PURCHASERET.DATERET, PURCHASERET.PURTRNO, PURCHASERET.PACKINGNO, SPRECEIVE.DATEREC, SPRECEIVE.SUPPLIER, PURCHASERET.AMOUNT, PURCHASERET.REMARKS, PURCHASERET.PRINTCOPY, PURCHASERET.PURTRNID " & _
                  "FROM PURCHASERET INNER JOIN SPRECEIVE ON PURCHASERET.PACKINGNO = SPRECEIVE.PACKINGNO WHERE PURCHASERET.DATERET='" & Format(dFrom.Value, "yyyy-mm-dd") & "' ORDER BY PURCHASERET.PURTRNO DESC;"
        Else
            Qry = "SELECT PURCHASERET.DATERET, PURCHASERET.PURTRNO, PURCHASERET.PACKINGNO, SPRECEIVE.DATEREC, SPRECEIVE.SUPPLIER, PURCHASERET.AMOUNT, PURCHASERET.REMARKS, PURCHASERET.PRINTCOPY, PURCHASERET.PURTRNID " & _
                  "FROM PURCHASERET INNER JOIN SPRECEIVE ON PURCHASERET.PACKINGNO = SPRECEIVE.PACKINGNO Where PURCHASERET.DATERET >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And PURCHASERET.DATERET <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' ORDER BY PURCHASERET.PURTRNO DESC;"
        
        End If
    ElseIf mySrc = "Inventory Adjustment" Then
        If dCode = 1 Then
            Qry = "SELECT INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO, INVNTRYADJ.DOCREF, INVNTRYADJ.REASON, INVNTRYADJ.REMARKS, INVNTRYADJ.PREPBY, INVNTRYADJ.APRVBY, INVNTRYADJ.PRINTCOPY, INVNTRYADJ.ADJTYPE " & _
                  "From InvntryADJ Where InvntryADJ.ADJDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And InvntryADJ.ADJTYPE = 'INVENTORY' ORDER BY INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO;"
        Else
            Qry = "SELECT INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO, INVNTRYADJ.DOCREF, INVNTRYADJ.REASON, INVNTRYADJ.REMARKS, INVNTRYADJ.PREPBY, INVNTRYADJ.APRVBY, INVNTRYADJ.PRINTCOPY, INVNTRYADJ.ADJTYPE " & _
                  "From InvntryADJ Where InvntryADJ.ADJDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And InvntryADJ.ADJDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' And InvntryADJ.ADJTYPE = 'INVENTORY' ORDER BY INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO;"
        End If
'******************** CODE UPDATED: 02/12/2019 **********************
    ElseIf mySrc = "WS Inventory Adjustment" Then
        Qry = "call sp_mrcwarehouse.sp_stockadj_summary(" & Val(dCode) & ", '" & Format(dFrom.Value, "yyyy-mm-dd") & "', '" & Format(dTo.Value, "yyyy-mm-dd") & "') "
'**************** END CODE UPDATED: 02/12/2019 **********************
    ElseIf mySrc = "Delivery Receipt" Then
        If dCode = 1 Then
            Qry = "SELECT DATE_FORMAT(invoutgoing.DRDATE, '%c/%d/%Y'), invoutgoing.DRNO, invoutgoing.DELIVERTO, invoutgoing.DOCREFNO, DATE_FORMAT(invoutgoing.DOCREFDATE, '%c/%d/%Y'), invoutgoing.RECBY, invoutgoing.REMARKS, invoutgoing.PRNCOPY, invoutgoing.DRID " & _
                  "From invoutgoing Where invoutgoing.DRDATE = '" & Format(dFrom.Value, "YYYY-MM-DD") & "' Order By invoutgoing.DRDATE DESC, invoutgoing.DRNO"
        Else
            Qry = "SELECT DATE_FORMAT(invoutgoing.DRDATE, '%c/%d/%Y'), invoutgoing.DRNO, invoutgoing.DELIVERTO, invoutgoing.DOCREFNO, DATE_FORMAT(invoutgoing.DOCREFDATE, '%c/%d/%Y'), invoutgoing.RECBY, invoutgoing.REMARKS, invoutgoing.PRNCOPY, invoutgoing.DRID " & _
                  "From invoutgoing Where invoutgoing.DRDATE >= '" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND invoutgoing.DRDATE <= '" & Format(dTo.Value, "YYYY-MM-DD") & "' Order By invoutgoing.DRDATE DESC, invoutgoing.DRNO"
        End If
    ElseIf mySrc = "Item Entry" Then
        Qry = "SELECT items.PARTNUMBER, items.DESCRIPTION, items.QTY, items.UM1, FORMAT(items.COST,2), items.CLASSIFICATION, items.LOCATION, items.SUPPLIER, items.ITEMID " & _
              "From items Order By items.PARTNUMBER, items.DESCRIPTION, items.CLASSIFICATION"
    ElseIf mySrc = "Branch Entry" Then
        Qry = "SELECT branch.BRANCHNAME, branch.ADDRESS, branch.COMPANYGROUP, branch.MANAGER, branch.TELNO, branch.EMAIL, branch.BRANCHID " & _
              "From branch Order By branch.BRANCHNAME;"
    ElseIf mySrc = "Supplier Entry" Then
        Qry = "SELECT supplier.SUPPLIERNAME, supplier.ADDRESS, supplier.CONTACTPERS, supplier.TELNO, supplier.EMAIL, supplier.TIN, supplier.SUPPLIERID " & _
              "From supplier Order by supplier.SUPPLIERNAME"
    ElseIf mySrc = "Inventory" Then
        Qry = "SELECT items.PARTNUMBER, items.DESCRIPTION, items.QTY, items.UM1, FORMAT(items.COST,2), items.CLASSIFICATION, " & _
           "items.LOCATION, items.SUPPLIER, items.REORDERPOINT, items.ITEMID From items Where items.QTY > 0 Order by items.CLASSIFICATION, items.PARTNUMBER;"
    ElseIf mySrc = "Detailed Receiving" Then
        If dCode = 1 Then
            Qry = "SELECT items.PARTNUMBER, items.DESCRIPTION, invtreg.REFERENCE, DATE_FORMAT(invtreg.TRANSDATE, '%c/%d/%Y') AS FIELD_1, invtreg.REC, format(invtreg.UNITPRICE, 2) AS FIELD_2, " & _
                  "format(invtreg.UNITPRICE * invtreg.REC, 2) AS TtlAmt, invtreg.SOURCE, invtreg.SourceDoc From items INNER JOIN invtreg ON (items.ITEMID = invtreg.ITEMID) " & _
                  "Where invtreg.TRANSTYPE = 'REC' AND invtreg.TRANSDATE = '" & Format(dFrom.Value, "YYYY-MM-DD") & "' ORDER BY invtreg.TRANSDATE, invtreg.REFERENCE, items.PARTNUMBER;"
        Else
            Qry = "SELECT items.PARTNUMBER, items.DESCRIPTION, invtreg.REFERENCE, DATE_FORMAT(invtreg.TRANSDATE, '%c/%d/%Y') AS FIELD_1, invtreg.REC, format(invtreg.UNITPRICE, 2) AS FIELD_2, " & _
                  "format(invtreg.UNITPRICE * invtreg.REC, 2) AS TtlAmt, invtreg.SOURCE, invtreg.SourceDoc From items INNER JOIN invtreg ON (items.ITEMID = invtreg.ITEMID) " & _
                  "Where invtreg.TRANSTYPE = 'REC' AND invtreg.TRANSDATE >= '" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND invtreg.TRANSDATE <= '" & Format(dTo.Value, "YYYY-MM-DD") & "' ORDER BY invtreg.TRANSDATE, invtreg.REFERENCE, items.PARTNUMBER;"
        End If
    ElseIf mySrc = "Detailed Outgoing" Then
        If dCode = 1 Then
            Qry = "SELECT items.PARTNUMBER, items.DESCRIPTION, invtreg.REFERENCE, DATE_FORMAT(invtreg.TRANSDATE, '%c/%d/%Y') AS FIELD_1, invtreg.OUT, " & _
                  "format(invtreg.UNITPRICE, 2) AS FIELD_2, format(invtreg.UNITPRICE * invtreg.OUT, 2) AS TtlAmt, invtreg.REMARKS, invtreg.SOURCE, invtreg.SourceDoc " & _
                  "From items INNER JOIN invtreg ON (items.ITEMID = invtreg.ITEMID) " & _
                  "Where invtreg.TRANSTYPE = 'OUT' AND invtreg.TRANSDATE = '" & Format(dFrom.Value, "YYYY-MM-DD") & "' Order By invtreg.TRANSDATE, invtreg.REFERENCE, items.PARTNUMBER;"
        Else
            Qry = "SELECT items.PARTNUMBER, items.DESCRIPTION, invtreg.REFERENCE, DATE_FORMAT(invtreg.TRANSDATE, '%c/%d/%Y') AS FIELD_1, invtreg.OUT, " & _
                  "format(invtreg.UNITPRICE, 2) AS FIELD_2, format(invtreg.UNITPRICE * invtreg.OUT, 2) AS TtlAmt, invtreg.REMARKS, invtreg.SOURCE, invtreg.SourceDoc " & _
                  "From items INNER JOIN invtreg ON (items.ITEMID = invtreg.ITEMID) " & _
                  "Where invtreg.TRANSTYPE = 'OUT' AND invtreg.TRANSDATE >= '" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND invtreg.TRANSDATE <= '" & Format(dTo.Value, "YYYY-MM-DD") & "' Order By invtreg.TRANSDATE, invtreg.REFERENCE, items.PARTNUMBER;"
        End If
    ElseIf mySrc = "Branch Order" Then
        If dCode = 1 Then
            Qry = "SELECT SPORDER.ORDERDATE, SPORDER.REFERENCE, SUPPLIER.SUPPLIERNAME, SPORDER.ORDERTYPE, SPORDER.CONFIRMNO, SPORDER.SUBMITTEDBY, SPORDER.APPROVEDBY, SPORDER.PRINTCOPY, SPORDER.QTYORDER, SPORDER.QTYREC, SPORDER.BACKORDER, SPORDER.BOLASTUPDATE " & _
                  "FROM SPORDER INNER JOIN SUPPLIER ON SPORDER.SUPPLIERID = SUPPLIER.SUPPLIERID Where SPORDER.ORDERDATE = '" & Format(dFrom.Value, "YYYY-MM-DD") & "' ORDER BY SPORDER.ORDERDATE DESC;"
'            Qry = "SELECT sporder.ORDERDATE, sporder.REFERENCE, sporder.`STATUS`, sporder.CONFIRMNO, sporder.QTYPN, sporder.QTYRECPN, sporder.BACKORDERPN, sporder.QTYORDER, sporder.QTYREC, sporder.BACKORDER, sporder.BOLASTUPDATE, sporder.PICKEDBY, sporder.PRINTCOPY " & _
'                  "From sporder Where SPORDER.ORDERDATE = '" & Format(dFrom.Value, "YYYY-MM-DD") & "' ORDER BY SPORDER.ORDERDATE DESC;"
'
        Else
            Qry = "SELECT SPORDER.ORDERDATE, SPORDER.REFERENCE, SUPPLIER.SUPPLIERNAME, SPORDER.ORDERTYPE, SPORDER.CONFIRMNO, SPORDER.SUBMITTEDBY, SPORDER.APPROVEDBY, SPORDER.PRINTCOPY, SPORDER.QTYORDER, SPORDER.QTYREC, SPORDER.BACKORDER, SPORDER.BOLASTUPDATE " & _
                  "FROM SPORDER INNER JOIN SUPPLIER ON SPORDER.SUPPLIERID = SUPPLIER.SUPPLIERID Where SPORDER.ORDERDATE >= '" & Format(dFrom.Value, "YYYY-MM-DD") & "' and SPORDER.ORDERDATE <= '" & Format(dTo.Value, "YYYY-MM-DD") & "' ORDER BY SPORDER.ORDERDATE DESC;"
'            Qry = "SELECT sporder.ORDERDATE, sporder.REFERENCE, sporder.`STATUS`, sporder.CONFIRMNO, sporder.QTYPN, sporder.QTYRECPN, sporder.BACKORDERPN, sporder.QTYORDER, sporder.QTYREC, sporder.BACKORDER, sporder.BOLASTUPDATE, sporder.PICKEDBY, sporder.PRINTCOPY " & _
'                  "From sporder Where SPORDER.ORDERDATE >= '" & Format(dFrom.Value, "YYYY-MM-DD") & "' and SPORDER.ORDERDATE <= '" & Format(dTo.Value, "YYYY-MM-DD") & "' ORDER BY SPORDER.ORDERDATE DESC;"
'
        End If
   ElseIf mySrc = "Branch Purchase Order" Then
            cmdUpdate.Visible = True
            CodeDateRange = 0
            If dCode = 1 Then
                 'Qry = "SELECT spordertaking.SQSERVE, spordertaking.BRANCHNAME, ' ', spordertaking.DATESERVE, spordertaking.DATECONFIRM, '', spordertaking.CONFIRMNO, " & _
                  '  "spordertaking.TOTALAMOUNT, spordertaking.BRANCHID, spordertaking.PRINTPICK, spordertaking.FINALQT " & _
                  '  "From spordertaking Where spordertaking.DATESERVE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' Order By spordertaking.ORDERIDSQ;"
                Qry = "SELECT sporderbranch.BRANCHNAME, sporderbranch.REFERENCE, sporderbranch.ORDERDATE, sporderbranch.`STATUS`, sporderbranch.ORDERTYPE, sporderbranch.QTYPN, sporderbranch.QTYORDER, " & _
                        "sporderbranch.LASTUPDATE, sporderbranch.PICKEDBY, sporderbranch.PRINTPICK, sporderbranch.SPORDERID, sporderbranch.BRANCHSPOID, sporderbranch.BRANCHID " & _
                        "From mrcwarehouse.sporderbranch Where sporderbranch.ORDERDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' ORDER BY sporderbranch.SPORDERID DESC;"
                CodeDateRange = 1
            Else
                ' Qry = "SELECT spordertaking.SQSERVE, spordertaking.BRANCHNAME, ' ', spordertaking.DATESERVE, spordertaking.DATECONFIRM,' ',spordertaking.CONFIRMNO, " & _
                  ' "spordertaking.TOTALAMOUNT, spordertaking.BRANCHID, spordertaking.PRINTPICK, spordertaking.FINALQT " & _
                 '  "From spordertaking Where spordertaking.DATESERVE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' AND spordertaking.DATESERVE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' Order by spordertaking.ORDERIDSQ;"
                    Qry = "SELECT sporderbranch.BRANCHNAME, sporderbranch.REFERENCE, sporderbranch.ORDERDATE, sporderbranch.`STATUS`, sporderbranch.ORDERTYPE, sporderbranch.QTYPN, sporderbranch.QTYORDER, " & _
                        "sporderbranch.LASTUPDATE, sporderbranch.PICKEDBY, sporderbranch.PRINTPICK, sporderbranch.SPORDERID, sporderbranch.BRANCHSPOID, sporderbranch.BRANCHID " & _
                        "From mrcwarehouse.sporderbranch Where sporderbranch.ORDERDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' and sporderbranch.ORDERDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' ORDER BY sporderbranch.SPORDERID DESC;"
                CodeDateRange = 2
            End If
    ElseIf mySrc = "SP PACKING" Then
            If dCode = 1 Then
            '****************** CODE UPDATED 11/03/17 ************************
                Qry = "SELECT stocktransfer.TRANSDATE, stocktransfer.DATEFINAL, Case When ISNULL(stocktransfer.DELIVEREDDATE) THEN '' ELSE date_format(stocktransfer.DELIVEREDDATE, '%Y-%m-%d') END as DELIVEREDDATE, CAST(IFNULL(TRANSNO,TRANSFERID) as Char(60)) as REFERENCE, IFNULL(tr.TTLPN,0) as TTLPN, IFNULL(tr.TTLQTY,0) as TTLQTY, stocktransfer.RECBRANCH, stocktransfer.ADDRESS, stocktransfer.PONO, stocktransfer.`STATUS`, stocktransfer.REMARKS, stocktransfer.CHECKER, stocktransfer.DRIVERNAME " & _
                        "From stocktransfer LEFT JOIN (Select COUNT(spitr.SPAREPARTID) as TTLPN, SUM(`OUT`) as TTLQTY, REPLACE(REPLACE(REPLACE(spitr.REFERENCE,'-MWH',''),'TR#',''),'STF#','') as REFERENCE From spitr Where spitr.TRANSTYPE = 'OUT' and spitr.REMARKS = 'TRANSFER' GROUP BY spitr.REFERENCE) as tr ON tr.REFERENCE = stocktransfer.TRANSFERID Where stocktransfer.TRANSDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' Order By stocktransfer.TRANSFERID DESC, stocktransfer.TRANSDATE Desc"
            Else
                Qry = "SELECT stocktransfer.TRANSDATE, stocktransfer.DATEFINAL, Case When ISNULL(stocktransfer.DELIVEREDDATE) THEN '' ELSE date_format(stocktransfer.DELIVEREDDATE, '%Y-%m-%d') END as DELIVEREDDATE, CAST(IFNULL(TRANSNO,TRANSFERID) as Char(60)) as REFERENCE, IFNULL(tr.TTLPN,0) as TTLPN, IFNULL(tr.TTLQTY,0) as TTLQTY, stocktransfer.RECBRANCH, stocktransfer.ADDRESS, stocktransfer.PONO, stocktransfer.`STATUS`, stocktransfer.REMARKS, stocktransfer.CHECKER, stocktransfer.DRIVERNAME " & _
                        "From stocktransfer LEFT JOIN (Select COUNT(spitr.SPAREPARTID) as TTLPN, SUM(`OUT`) as TTLQTY, REPLACE(REPLACE(REPLACE(spitr.REFERENCE,'-MWH',''),'TR#',''),'STF#','') as REFERENCE From spitr Where spitr.TRANSTYPE = 'OUT' and spitr.REMARKS = 'TRANSFER' GROUP BY spitr.REFERENCE) as tr ON tr.REFERENCE = stocktransfer.TRANSFERID Where stocktransfer.TRANSDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' AND stocktransfer.TRANSDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' Order By stocktransfer.TRANSFERID DESC, stocktransfer.TRANSDATE Desc"
            '************** END CODE UPDATED 11/03/17 ************************
            End If
            
    ElseIf mySrc = "AllocBBO" Then
            If dCode = 1 Then
                Qry = "SELECT backorderalloc.ALLOCID, DATE_FORMAT(backorderalloc.ALLOCDATE, '%c/%d/%Y'), branch.BRANCHNAME, backorderalloc.PREPAREDBY, backorderalloc.APPROVEDBY " & _
                        "From backorderalloc INNER JOIN branch ON (backorderalloc.BRANCHID = branch.BRANCHID) Where backorderalloc.ALLOCDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' Order By backorderalloc.ALLOCID DESC, backorderalloc.ALLOCDATE DESC"
            Else
                 Qry = "SELECT backorderalloc.ALLOCID, DATE_FORMAT(backorderalloc.ALLOCDATE, '%c/%d/%Y'), branch.BRANCHNAME, backorderalloc.PREPAREDBY, backorderalloc.APPROVEDBY " & _
                        "From backorderalloc INNER JOIN branch ON (backorderalloc.BRANCHID = branch.BRANCHID) Where backorderalloc.ALLOCDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' and backorderalloc.ALLOCDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' Order By backorderalloc.ALLOCID DESC, backorderalloc.ALLOCDATE DESC"
 
            End If
            
    ElseIf mySrc = "Cashier Cash Flow" Then
        If dCode = 1 Then
            Qry = "SELECT DATE_FORMAT(cashflow.CFDATE,'%c/%d/%Y'), cashflow.PREPAREDBY, cashflow.APPROVEDBY, cashflow.REMARKS, cashflow.CASHFLOWID " & _
                  "From cashflow Where cashflow.CFDATE = '" & Format(dFrom.Value, "YYYY-MM-DD") & "' Order By cashflow.CFDATE Desc;"
        Else
            Qry = "SELECT DATE_FORMAT(cashflow.CFDATE,'%c/%d/%Y'), cashflow.PREPAREDBY, cashflow.APPROVEDBY, cashflow.REMARKS, cashflow.CASHFLOWID " & _
                  "From cashflow Where cashflow.CFDATE >= '" & Format(dFrom.Value, "YYYY-MM-DD") & "' and cashflow.CFDATE <= '" & Format(dTo.Value, "YYYY-MM-DD") & "' Order By cashflow.CFDATE Desc;"
        End If
    ElseIf mySrc = "Costing" Then
        If dCode = 1 Then
            Qry = "Call sp_acctpayinvcng.sp_acctpayinvcng_summary('" & Format(dFrom.Value, "YYYY-MM-DD") & "','" & Format(dFrom.Value, "YYYY-MM-DD") & "','');"
        Else
        'Call modAPI.sp_acctpayinvcng_summary(dateFrom, dateTo, cboStat.Text, tmpRs)"
            Qry = "Call sp_acctpayinvcng.sp_acctpayinvcng_summary('" & Format(dFrom.Value, "YYYY-MM-DD") & "','" & Format(dTo.Value, "YYYY-MM-DD") & "','');"
        End If
    ElseIf mySrc = "SRP Pricing" Then
        If dCode = 1 Then
            Qry = "Call sp_srppricing.sp_srppricing_summary('" & Format(dFrom.Value, "YYYY-MM-DD") & "','" & Format(dFrom.Value, "YYYY-MM-DD") & "');"
        Else
            Qry = "Call sp_srppricing.sp_srppricing_summary('" & Format(dFrom.Value, "YYYY-MM-DD") & "','" & Format(dTo.Value, "YYYY-MM-DD") & "');"
        End If
'*********************** CODE UPDATED: 08/08/2017 *********************************
    ElseIf mySrc = "SETTLEMENT DISCREPANCY" Then
        If dCode = 1 Then
            Qry = "Call sp_mrcwarehouse.sp_WHSDiscrepancy_Summary_v2('" & sysBranchCode & "', " & dCode & ",'" & Format(dFrom.Value, "YYYY-MM-DD") & "','" & Format(dTo.Value, "YYYY-MM-DD") & "');"
        Else
            Qry = "Call sp_mrcwarehouse.sp_WHSDiscrepancy_Summary_v2('" & sysBranchCode & "', " & dCode & ",'" & Format(dFrom.Value, "YYYY-MM-DD") & "','" & Format(dTo.Value, "YYYY-MM-DD") & "');"
        End If
    ElseIf mySrc = "SDiscrepancy Adjustment" Then
        If dCode = 1 Then
            Qry = "SELECT INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO, INVNTRYADJ.DOCREF, INVNTRYADJ.REASON, INVNTRYADJ.REMARKS, INVNTRYADJ.PREPBY, INVNTRYADJ.APRVBY, INVNTRYADJ.PRINTCOPY " & _
                  "From InvntryADJ Where InvntryADJ.ADJDATE = '" & Format(dFrom.Value, "yyy-mm-dd") & "' And InvntryADJ.ADJTYPE <> 'INVENTORY' ORDER BY INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO;"
        Else
            Qry = "SELECT INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO, INVNTRYADJ.DOCREF, INVNTRYADJ.REASON, INVNTRYADJ.REMARKS, INVNTRYADJ.PREPBY, INVNTRYADJ.APRVBY, INVNTRYADJ.PRINTCOPY " & _
                  "From InvntryADJ Where InvntryADJ.ADJDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And InvntryADJ.ADJDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' And InvntryADJ.ADJTYPE <> 'INVENTORY' ORDER BY INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO;"
        End If
'******************* END CODE UPDATED: 08/08/2017 *********************************
'************************CODE UPDATED BY IAN: 03/05/2018 **********************************
    ElseIf mySrc = "SRP Request form" Then
        If dCode = 1 Then
            Qry = "Call sp_srppricing.sp_branchSRP_summary('" & Format(dFrom.Value, "YYYY-MM-DD") & "','" & Format(dFrom.Value, "YYYY-MM-DD") & "');"
        Else
            Qry = "Call sp_srppricing.sp_branchSRP_summary('" & Format(dFrom.Value, "YYYY-MM-DD") & "','" & Format(dTo.Value, "YYYY-MM-DD") & "');"
        End If
'************************ END CODE UPDATED BY IAN: 03/05/2018 **********************************
'************************CODE UPDATED BY IAN: 06/14/2019 **********************************
    ElseIf mySrc = "PO ADJUSTMENT" Then
        Qry = "call sp_mrcwarehouse.sp_POAdj_summary(" & Val(dCode) & ",'" & Format(dFrom.Value, "YYYY-MM-DD") & "','" & Format(dTo.Value, "YYYY-MM-DD") & "')"
'************************ END CODE UPDATED BY IAN: 06/14/2019 **********************************
'************************CODE UPDATED BY IAN: 06/14/2019 **********************************
    ElseIf mySrc = "RFP" Then
        Qry = "call sp_acctpayinvcng.sp_rfp_summary(" & Val(dCode) & ",'" & Format(dFrom.Value, "YYYY-MM-DD") & "','" & Format(dTo.Value, "YYYY-MM-DD") & "')"
'************************ END CODE UPDATED BY IAN: 06/14/2019 **********************************
'************************CODE UPDATED BY IAN: 06/14/2019 **********************************
    ElseIf mySrc = "RFP RECONCILIATION" Then
        CodeDateRange = Val(dCode)
        Qry = "call sp_acctpayinvcng.sp_rfp_reconciliation_summ(" & Val(dCode) & ",'" & Format(dFrom.Value, "YYYY-MM-DD") & "','" & Format(dTo.Value, "YYYY-MM-DD") & "')"
'************************ END CODE UPDATED BY IAN: 06/14/2019 **********************************
    ElseIf mySrc = "ADJUSTMENT REQUEST" Then
        Qry = "call sp_mrcretails.sp_other_adjustment_summary('" & sysBranchCode & "'," & Val(dCode) & ",'" & Format(dFrom.Value, "yyyy-mm-dd") & "','" & Format(dTo.Value, "yyyy-mm-dd") & "')"
    ElseIf mySrc = "CANCELLATION MODULE" Then
        Qry = "call sp_mrcwarehouse.sp_cancellation_module_summary('" & sysBranchCode & "'," & Val(dCode) & ",'" & Format(dFrom.Value, "yyyy-mm-dd") & "','" & Format(dTo.Value, "yyyy-mm-dd") & "')"
    End If
    If rsSummary.State = adStateOpen Then rsSummary.Close
    rsSummary.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsSummary.RecordCount = 0 Then
        flxSum.Clear
        flxSum.Rows = 2
        flxSum.Cols = TCol
        Call Unload(frm_OnProcess2)
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    If mySrc = "ADJUSTMENT REQUEST" Then
        With rsSummary
            .MoveFirst
            Call RXResizer1_ResizeComplete
            flxSum.Rows = .RecordCount + 1
            If .RecordCount > 0 Then
                For i = 1 To flxSum.Rows - 1
                'oa.req_date, oa.req_status, oa.ticketno, ifnull(oa.controlno,'') as controlno, ifnull(date_format(oa.date_final,'%m/%d/%Y'),'') as date_final,
                'oa.subj_request, emp.EMPLOYEENAME as REQUESTBY, ifnull(oa.aud_apprvd_by,'') as aud_apprvd_by, ifnull(oa.mis_incharge,'') as mis_incharge,
                'ifnull(oa.script_file,'') as script_file
                    flxSum.TextMatrix(i, 1) = !req_date
                    flxSum.TextMatrix(i, 2) = !req_status
                    flxSum.TextMatrix(i, 3) = !ticketno
                    flxSum.TextMatrix(i, 4) = !controlno
                    flxSum.TextMatrix(i, 5) = !adjno
                    flxSum.TextMatrix(i, 6) = !date_final
                    flxSum.TextMatrix(i, 7) = !subj_request
                    flxSum.TextMatrix(i, 8) = !REQUESTBY
                    flxSum.TextMatrix(i, 9) = !aud_apprvd_by
                    flxSum.TextMatrix(i, 10) = mis_incharge
                    flxSum.TextMatrix(i, 11) = script_file
                    .MoveNext
                Next i
            End If
        End With
    Else
        Set flxSum.DataSource = rsSummary
    End If
    ActRow = 1
    txtCriteria.Text = ""
    lblRemarks.Caption = ""
    flxSum.Visible = True
    'If mySrc = "Branch Purchase Order" Then Call DisplayRef
    Call RXResizer1_ResizeComplete
End Sub

Private Sub DisplayRef()
    Dim qryRef As String, r As Integer, Fil As Long, ConRef As String
    Dim rsDRef As New ADODB.Recordset
    
    For Fil = 1 To flxSum.Rows - 1
        ConRef = ""
        qryRef = "SELECT spordrconfirm.REFERENCE From spordrconfirm Where spordrconfirm.SQSERVE =" & flxSum.TextMatrix(Fil, 1) & ";"
        If rsDRef.State = adStateOpen Then rsDRef.Close
        rsDRef.Open qryRef, LYRcon, adOpenKeyset, adLockOptimistic
        If rsDRef.RecordCount > 0 Then
            rsDRef.MoveFirst
                 For r = 1 To rsDRef.RecordCount
                    ConRef = ConRef & rsDRef!REFERENCE & IIf((rsDRef.RecordCount = 1), "", ",")
                    rsDRef.MoveNext
                Next r
                flxSum.TextMatrix(Fil, 3) = IIf((rsDRef.RecordCount = 1), ConRef, Left$(ConRef, Len(ConRef) - 1))
        End If
    Next Fil
    
 End Sub

Private Sub ShowCommands(CmdType As Integer)
    Dim TmpLeft As Integer
    
    Select Case CmdType
        Case 1  'Show Form, Print Summary, Exit Report
            'cmdExit.Left = cmdFilter.Left
            If mySrc = "RFP RECONCILIATION" Then
                cmdShow.Enabled = False
            Else
                cmdShow.Enabled = True
            End If
            'as is
        Case 2  'Show Form, Print Summary, Filter, Exit Report
            cmdExit.Left = cmdFilter.Left + cmdFilter.Width
            cmdFilter.Visible = True
            
        Case 3  ''Show Form, Print Summary, Transmittal, Sales Drive, Filter, Exit Report
            TmpLeft = cmdSalesDrive.Left
            cmdSalesDrive.Left = cmdFilter.Left
            cmdFilter.Left = TmpLeft
            cmdFilter.Visible = True
            cmdSalesDrive.Visible = True
            cmdTransmittal.Visible = True
        Case 4  'Show Form, Print Summary, Sales Drive, Exit Report
            'Change Print Summary to Collection Sheet, Sales Drive to Account Summary
            cmdSalesDrive.Left = cmdFilter.Left
            cmdExit.Left = cmdTransmittal.Left
            cmdPrint.Caption = "Collection Sheet"
            cmdSalesDrive.Caption = "Acct. Summary"
            cmdSalesDrive.Visible = True
        
        Case 5  'Show Form, Print Summary, Sales Drive, Exit Report
            cmdSalesDrive.Left = cmdFilter.Left
            cmdExit.Left = cmdTransmittal.Left
            cmdSalesDrive.Caption = "Collectors"
            cmdSalesDrive.Visible = True
        
    End Select
End Sub


Private Sub cboMonth_Click()
    If cboMonth.ListIndex = -1 Then Exit Sub
    
    If mySrc = "New Account" Or mySrc = "Monthly Collection" Then
        dFrom.Value = DateSerial(Val(txtYear.Text), cboMonth.ListIndex + 1, 1)
        dTo.Value = DateAdd("M", 1, dFrom.Value)
        dTo.Value = DateAdd("D", -1, dTo.Value)
    Else
        dFrom.Value = DateSerial(Val(txtYear.Text), cboMonth.ListIndex + 1, 7)
        dTo.Value = DateSerial(Year(DateAdd("M", 1, dFrom.Value)), Month(DateAdd("M", 1, dFrom.Value)), 6)
    End If
    
    Call RefreshRecord(2)
End Sub

Private Sub cmdChkClosed_Click()
    If mySrc = "Closed Account" Then
        Call ReviewClosedAcct
    ElseIf mySrc = "Deposited Unit" Then
        Call UpdateDepositedCI
    ElseIf mySrc = "Sales Invoice" Then
        Call FixAgentSales
    ElseIf mySrc = "Monthly Commission" Then
        Call FixCommissionTax
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
   ' Marketing.tbMain.Visible = True
End Sub

Private Sub cmdFilter_Click()
    Dim i As Integer
    
    If mySrc = "Detailed Outgoing Items" Then
        frmFilter.MySource = "Detailed Outgoing"
        For i = 0 To 2
             With frmFilter.cboLookIn(i)
               .AddItem "PRODUCT BRAND"
               .AddItem "PRODUCT MODEL"
               .AddItem "CLASSIFICATION"
               .AddItem "PRODUCT STATUS"
               .AddItem "PRODUCT TYPE"
               .AddItem "SERIAL NO."
               .AddItem "TRANSACTION TYPE"
               .AddItem "DELIVERED TO"
             End With
        Next i
        ', items.COLOR, , invtreg.OUT, invtreg.REFERENCE, DATE_FORMAT(invtreg.TRANSDATE, '%c/%d/%Y') AS FIELD_1, , , , productline.PRODINIT, ,  "
        frmFilter.fldNames = "productline.BRAND*productline.MODEL*productline.CLASSIFICATION*invtreg.REMARKS*productline.PRODINIT*items.SERIALNO*invtreg.SRCTYPE*invtreg.SOURCE"
        frmFilter.fldNdx = "" 'start w/ 0
        frmFilter.fldDates = ""
        frmFilter.fldSupplier = 6  'Used no > index to exclude field supplier
        frmFilter.Show vbModal
    ElseIf mySrc = "Detailed Receiving Receipt" Then
        frmFilter.MySource = "Detailed Receiving"
        For i = 0 To 2
             With frmFilter.cboLookIn(i)
               .AddItem "PRODUCT BRAND"
               .AddItem "PRODUCT MODEL"
               .AddItem "CLASSIFICATION"
               .AddItem "PRODUCT STATUS"
               .AddItem "PRODUCT TYPE"
               .AddItem "SERIAL NO."
               .AddItem "PRODUCT SOURCE"
             End With
        Next i
        frmFilter.fldNames = "productline.BRAND*productline.MODEL*productline.CLASSIFICATION*invtreg.REMARKS*productline.PRODINIT*items.SERIALNO*invtreg.SRCTYPE*invtreg.SOURCE"
        frmFilter.fldNdx = "" 'start w/ 0
        frmFilter.fldDates = ""
        frmFilter.fldSupplier = 6  'Used no > index to exclude field supplier
        frmFilter.Show vbModal
    ElseIf mySrc = "Disbursement" Then
        frmFilter.MySource = "Disbursement"
        For i = 0 To 2
             With frmFilter.cboLookIn(i)
               .AddItem "CV NUMBER"
               .AddItem "DISBURSEMENT TYPE"
               .AddItem "ACCOUNT DESCRIPTION"
               .AddItem "PAYEE"
               .AddItem "PARTICULARS"
               .AddItem "DETAILS"
               .AddItem "AMOUNT"
               .AddItem "REMARKS"
             End With
        Next i
        frmFilter.fldNames = "disbursement.CVNUMBER*disbursement.DISBURSETYPE*disbursement.ACCTDESC*disbursement.PAYEE*disbursement.PARTICULARS*disbursement.DETAILS*disbursement.AMOUNT*disbursement.REMARKS"
        frmFilter.fldNdx = "6" 'start w/ 0
        frmFilter.fldDates = ""
        frmFilter.fldSupplier = 99  'Used no > index to exclude field supplier
        frmFilter.Show vbModal
    
    End If
End Sub

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxSum
        For i = fRow To .Rows - 1
            For j = fCol To TCol - 1
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .col = 1: .row = i
                    .ColSel = TCol - 1: .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = TCol - 1, i + 1, i)
                    fCol = IIf(j = TCol - 1, 1, j + 1)
                    ActRow = i
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

Private Sub cmdFind2_Click()
    Dim i As Integer, j As Integer
        
    If txtFind2.Text = "" Then Exit Sub
    With flxSum
        For i = fRow To .Rows - 1
            For j = fCol To TCol - 1
                If InStr(.TextMatrix(i, j), txtFind2.Text) > 0 Then
                    .col = 1: .row = i
                    .ColSel = TCol - 1: .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = TCol - 1, i + 1, i)
                    fCol = IIf(j = TCol - 1, 1, j + 1)
                    ActRow = i
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

Private Sub cmdNext_Click()
    cmdNext.Tag = "2"
    cmdNext.Visible = False
    cmdPrev.Left = cmdNext.Left
    cmdPrev.Top = cmdNext.Top
    cmdPrev.Visible = True
    dTo.Visible = True
    cmdRefresh.Visible = True
    dTo.Value = dFrom.Value
End Sub

Private Sub cmdPrev_Click()
    cmdNext.Tag = "1"
    cmdNext.Visible = True
    cmdPrev.Visible = False
    dTo.Visible = False
    cmdRefresh.Visible = False
    Call RefreshRecord(1)
End Sub

Private Sub ExtractRFP()

On Error GoTo er
    Dim XLapp As Excel.Application
    Dim wbkMVR As Excel.Workbook
    Dim wksLineA1 As Excel.Worksheet
    Dim sSaveAsFileName
    
   
    If flxSum.TextMatrix(1, 1) = "" Then
        MsgBox "No Data to Extract", vbCritical, "Error"
        Exit Sub
    End If
    
            With dlgDEMO
                .CancelError = False
                .InitDir = mstrLastDir
                .Flags = cdlOFNHideReadOnly
                '.FileName = App.Path & "\" & Format(dFrom.Value, "MM_YY") & lblConnect.Tag & IIf(Index <> 3, "_NP", "") & ".xlsx"
                .filename = ""
                .filter = "Microsoft Excel (*.xls)|*.xls|All Files(*.*)|*.*"
                .ShowOpen
                If dlgDEMO.filename = "" Then Exit Sub
                sSaveAsFileName = .filename
            End With
            
            Set XLapp = CreateObject("Excel.Application")
            XLapp.Visible = True
            XLapp.Workbooks.Add
            Set wbkMVR = XLapp.ActiveWorkbook
            Set wksLineA1 = wbkMVR.Sheets(1)
        
            If mySrc = "RFP" Then
                    'PARTNUMBER, DESCRIPTION, REMARKS, INVDATE,
                    'TRANSNO, OUT, PRICE, SUPPLIERNAME, INVTYPE,
                    'CHARGETO, ADDRESS, FRONTLINER, Branch Name, Branch Category
                    wksLineA1.Name = sysBranchCode & " SERVICE RATIO REPORT"
                    wksLineA1.Cells(1, 1).Value = "RFP DATE"
                    wksLineA1.Cells(1, 2).Value = "RFP NO."
                    wksLineA1.Cells(1, 3).Value = "STATUS"
                    wksLineA1.Cells(1, 4).Value = "SUPPLIER"
                    wksLineA1.Cells(1, 5).Value = "TERM"
                    wksLineA1.Cells(1, 6).Value = "RFP TYPE"
                    wksLineA1.Cells(1, 7).Value = "TOTAL INV. AMOUNT"
                    wksLineA1.Cells(1, 8).Value = "TOTAL DEDUCTION"
                    wksLineA1.Cells(1, 9).Value = "TOTAL PAYABLE"
                    wksLineA1.Cells(1, 10).Value = "SPCRR AMOUNT"
                    wksLineA1.Cells(1, 11).Value = "TREASURY RECEIVED"
                    wksLineA1.Cells(1, 12).Value = "BANK ACCOUNT"
                    wksLineA1.Cells(1, 13).Value = "CHECK NO"
                    wksLineA1.Cells(1, 14).Value = "RECONCILIATION"
                                        
                    With flxSum
                        DoEvents
                       With tabWait
                            .Top = (Me.Height - .Height) \ 2
                            .Left = (Me.Width - .Width) \ 2
                            .Caption = "Processing.."
                            .Visible = True
                            .ZOrder
                        End With
                        'cmdSaveExcel.Enabled = False
                        cmdExit.Enabled = False
                        DoEvents
                        'Pbar.Value = 0
                        'Pbar.Max = .Rows - 1
                        For i = 1 To .Rows - 1
                            For j = 1 To 14
'                                If j = 6 Then
'                                    wksLineA1.Cells(i + 1, j).Value = ""
'                                Else
                                    wksLineA1.Cells(i + 1, j).Value = .TextMatrix(i, j)
                                'End If
                            Next j
                            'Pbar.Value = Pbar.Value + 1
                            
                            DoEvents
                        Next i
                    End With
                    XLapp.DisplayAlerts = False
            End If
            
                        
                       
                wbkMVR.SaveAs filename:=sSaveAsFileName, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
                
                  
                wbkMVR.Close SaveChanges:=False
                XLapp.Quit
 
                Set XLapp = Nothing
                Set wbkMVR = Nothing
                Set wksLineA1 = Nothing
                MsgBox "Finish Saving to Excel!", vbInformation + vbOKOnly, MySource
                 tabWait.Top = 12000
                tabWait.Visible = False
                'cmdSaveExcel.Enabled = True
                cmdExit.Enabled = True
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdPrint_Click()
    Dim dCov As String, Qry As String, JOid As String, myCV As String
    Dim AdvLiq As Currency, AFLseries As String, pQRY As String
    Dim i As Integer
           
    If mySrc = "Inventory" Then
        Qry = "SELECT items.PARTNUMBER, items.DESCRIPTION, items.QTY, items.UM1, items.COST, items.LOCATION, items.CLASSIFICATION " & _
              "From items Where items.QTY > 0 Order By items.CLASSIFICATION, items.PARTNUMBER;"
        dCov = "As of " & Format(dFrom.Value, "Mmmm dd, yyyy Dddd")
        frmPrn.QryPrn = Qry
        frmPrn.ParamValue1 = dCov
        frmPrn.ParamValue2 = IIf(lblRemarks.Caption <> "", "Filtered by " & lblRemarks.Caption, "")
        Qry = "SELECT settings.VALUE1 From settings Where settings.DESCRIPTION = 'DEPARTMENT';"
        frmPrn.ParamValue3 = GetFieldValue(Qry)
        frmPrn.MySource = "INVENTORY"
        frmPrn.Show
    ElseIf mySrc = "Disbursement" Then
        If txtCriteria.Text = "" Then
            If Val(cmdNext.Tag) = 1 Then
                Qry = "SELECT SUM(advliquidation.AFLAMOUNT) As sumAFLamt From advliquidation " & _
                      "Where advliquidation.AFLDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' AND advliquidation.PRINTCOPY > 0 AND advliquidation.REMARKS <> 'CANCELLED';"
                AdvLiq = Val(GetFieldValue(Qry))
                Qry = "SELECT advliquidation.AFLNO From advliquidation " & _
                      "Where advliquidation.AFLDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' AND  advliquidation.PRINTCOPY > 0 AND advliquidation.REMARKS <> 'CANCELLED' Order By advliquidation.AFLNO;"
                AFLseries = GetOR(Qry)
                'Qry = "SELECT date_format(CVDATE,'%c/%d/%Y') AS CVDATE, CVNUMBER, DISBURSETYPE, ACCTDESC, PAYEE, PARTICULARS, AMOUNT " & _
                      "From disbursement WHERE (CVDATE ='" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND PRINTCOPY >0) ORDER BY ACCTDESC ASC;"
                Qry = "SELECT date_format(CVDATE,'%c/%d/%Y') AS CVDATE, CVNUMBER, DISBURSETYPE, ACCTDESC, PAYEE, PARTICULARS, AMOUNT From disbursement " & _
                      "WHERE CVDATE ='" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND DISBURSETYPE NOT LIKE 'Due From%' AND DISBURSETYPE<>'Withdrawal/Personal' AND PRINTCOPY>0 ORDER BY ACCTDESC ASC;"  'AND DISBURSETYPE <> 'General Management Office'
                dCov = Format(dFrom.Value, "Dddd, Mmm. dd, yyyy")
                
            Else
                Qry = "SELECT SUM(advliquidation.AFLAMOUNT) As sumAFLamt From advliquidation " & _
                      "Where advliquidation.AFLDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' AND advliquidation.AFLDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' AND advliquidation.PRINTCOPY > 0 AND advliquidation.REMARKS <> 'CANCELLED';"
                AdvLiq = Val(GetFieldValue(Qry))
                Qry = "SELECT advliquidation.AFLNO From advliquidation " & _
                      "Where advliquidation.AFLDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' AND advliquidation.AFLDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' AND advliquidation.PRINTCOPY > 0 AND advliquidation.REMARKS <> 'CANCELLED' Order By advliquidation.AFLNO;"
                AFLseries = GetOR(Qry)
                'Qry = "SELECT date_format(CVDATE,'%c/%d/%Y') AS CVDATE, CVNUMBER, DISBURSETYPE, ACCTDESC, PAYEE, PARTICULARS, AMOUNT " & _
                      "From disbursement WHERE (CVDATE >='" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND CVDATE <='" & Format(dTo.Value, "YYYY-MM-DD") & "' AND PRINTCOPY>0) ORDER BY CVDATE DESC, ACCTDESC ASC;"
                Qry = "SELECT date_format(CVDATE,'%c/%d/%Y') AS CVDATE, CVNUMBER, DISBURSETYPE, ACCTDESC, PAYEE, PARTICULARS, AMOUNT From disbursement " & _
                      "WHERE CVDATE >='" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND CVDATE <='" & Format(dTo.Value, "YYYY-MM-DD") & "' AND DISBURSETYPE NOT LIKE 'Due From%' AND DISBURSETYPE<>'Withdrawal/Personal' AND PRINTCOPY>0 ORDER BY ACCTDESC ASC;"      'AND DISBURSETYPE <> 'General Management Office'
                dCov = Format(dFrom.Value, "mm/dd/yyyy") & " to " & Format(dTo.Value, "mm/dd/yyyy")
            End If
        Else
            If Val(cmdNext.Tag) = 1 Then
                Qry = "SELECT date_format(CVDATE,'%c/%d/%Y') AS CVDATE, CVNUMBER, DISBURSETYPE, ACCTDESC, PAYEE, PARTICULARS, format(AMOUNT,2) AS AMOUNT From disbursement " & _
                      txtCriteria.Text & " AND CVDATE ='" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND DISBURSETYPE NOT LIKE 'Due From%' AND DISBURSETYPE<>'Withdrawal/Personal' ORDER BY CVDATE ASC, CVNUMBER ASC;"    ' AND DISBURSETYPE <> 'General Management Office'
                dCov = Format(dFrom.Value, "Dddd, Mmm. dd, yyyy")
            
            Else
                Qry = "SELECT date_format(CVDATE,'%c/%d/%Y') AS CVDATE, CVNUMBER, DISBURSETYPE, ACCTDESC, PAYEE, PARTICULARS, format(AMOUNT,2) AS AMOUNT From disbursement " & _
                      txtCriteria.Text & " AND CVDATE >='" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND CVDATE <='" & Format(dTo.Value, "YYYY-MM-DD") & "' AND DISBURSETYPE NOT LIKE 'Due From%' AND DISBURSETYPE<>'Withdrawal/Personal' ORDER BY CVDATE ASC, CVNUMBER ASC;"    ' AND DISBURSETYPE <> 'General Management Office'
                dCov = Format(dFrom.Value, "mm/dd/yyyy") & " to " & Format(dTo.Value, "mm/dd/yyyy")
            End If
        
            AdvLiq = 0
            AFLseries = ""
        End If
        frmPrn.QryPrn = Qry
        frmPrn.ParamValue1 = dCov
        Qry = "SELECT EMPLOYEE From transignatory WHERE (TRANSACTION ='PCV (1)' AND DESIGNATION ='DISBURSER');"
        frmPrn.ParamValue2 = GetFieldValue(Qry)
        frmPrn.ParamValue3 = "Total Adv. for Liquidation [" & AFLseries & "]"
        frmPrn.ParamValue4 = AdvLiq
        frmPrn.ParamValue5 = ""
        Qry = "SELECT BRANCH.BRANCHNAME From Branch WHERE (((BRANCH.BRANCHID)=" & Val(Marketing.Tag) & "));"
        frmPrn.ParamValue6 = GetFieldValue(Qry)
        Qry = "SELECT BRANCH.ADDRESS From Branch WHERE (((BRANCH.BRANCHID)=" & Val(Marketing.Tag) & "));"
        frmPrn.ParamValue7 = GetFieldValue(Qry)
        Qry = "SELECT VOUCHERNO From checkvoucher WHERE (CVDATE ='" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND PRINTCOPY > 0);"
        myCV = GetFieldValue(Qry)
        If Val(cmdNext.Tag) = 1 And myCV <> "" Then
            Qry = "SELECT AMOUNT From CHECKvoucher WHERE (CVDATE ='" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND PRINTCOPY > 0);"
            frmPrn.ParamValue8 = "Check Voucher No. " & myCV & " : " & Format(Val(GetFieldValue(Qry)), "#,##0.00")
        Else
            frmPrn.ParamValue8 = ""
        End If
        If txtCriteria.Text <> "" Then frmPrn.ParamValue8 = "Filtered Report: " & lblRemarks.Caption
        Qry = "SELECT EMPLOYEE From transignatory WHERE (TRANSACTION ='PCV (2)' AND DESIGNATION ='MANAGER');"
        frmPrn.ParamValue9 = GetFieldValue(Qry)
        frmPrn.MySource = "DISBURSEMENT"
        frmPrn.Show
        
    ElseIf mySrc = "Job Order" Then
        If Val(cmdNext.Tag) = 1 Then
            Qry = "SELECT svcjoborder.DATEFIN, svcjoborder.JOBORDERNO, svcustomer.CUSTNAME, svcaccount.MODEL, svcjoborder.PAYTYPE, svcjoborder.PAYMENT, svcjoborder.DISCOUNT, svcjoborder.ORNO, svcjoborder.SVCTYPE, svcjoborder.REMARKS, SvcJobOrder.MECHANIC " & _
                  "From SvcJobOrder INNER JOIN svcustomer ON (svcjoborder.CUSTID = svcustomer.CUSTID) INNER JOIN svcaccount ON (svcjoborder.ACCOUNTID = svcaccount.ACCOUNTID) " & _
                  "Where svcjoborder.DATEFIN = '" & Format(dFrom.Value, "YYYY-MM-DD") & "' And svcjoborder.PRINTCOPY > 0 Order By svcjoborder.DATEFIN, SvcJobOrder.JOBORDERNO"
            dCov = "Date Covered: " & Format(dFrom.Value, " MMMM DD, YYYY")
        Else
            Qry = "SELECT svcjoborder.DATEFIN, svcjoborder.JOBORDERNO, svcustomer.CUSTNAME, svcaccount.MODEL, svcjoborder.PAYTYPE, svcjoborder.PAYMENT, svcjoborder.DISCOUNT, svcjoborder.ORNO, svcjoborder.SVCTYPE, svcjoborder.REMARKS, SvcJobOrder.MECHANIC " & _
                  "From SvcJobOrder INNER JOIN svcustomer ON (svcjoborder.CUSTID = svcustomer.CUSTID) INNER JOIN svcaccount ON (svcjoborder.ACCOUNTID = svcaccount.ACCOUNTID) " & _
                  "Where svcjoborder.DATEFIN >= '" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND svcjoborder.DATEFIN <= '" & Format(dTo.Value, "YYYY-MM-DD") & "' And svcjoborder.PRINTCOPY > 0 Order By svcjoborder.DATEFIN, SvcJobOrder.JOBORDERNO"
            dCov = "FROM " & Format(dFrom.Value, "MMMM DD, YYYY") & " to " & Format(dTo.Value, "MMMM DD, yyyy")
        End If
        frmPrn.QryPrn = Qry
        Qry = "SELECT settings.REMARKS From settings Where settings.DESCRIPTION = 'SYSTEM BRANCH';"
        frmPrn.ParamValue1 = GetFieldValue(Qry)
        Qry = "SELECT BRANCH.ADDRESS From Branch WHERE (((BRANCH.BRANCHID)=" & Val(Marketing.Tag) & "));"
        frmPrn.ParamValue2 = GetFieldValue(Qry)
        frmPrn.ParamValue3 = dCov
        Qry = "SELECT concat(employees.FIRSTNAME,' ',employees.LASTNAME) as PrepBy From employees Where employees.EmpId = " & UserID & ";"
        frmPrn.ParamValue4 = GetFieldValue(Qry)
        Qry = "SELECT employees.`POSITION` From employees Where employees.EmpId = " & UserID & ";"
        frmPrn.ParamValue5 = GetFieldValue(Qry)
        Qry = "SELECT transignatory.EMPLOYEE From transignatory Where transignatory.TRANSACTION = 'Job Order Summary (1)';"
        frmPrn.ParamValue6 = GetFieldValue(Qry)
        Qry = "SELECT transignatory.DESIGNATION From transignatory Where transignatory.TRANSACTION = 'Job Order Summary (1)';"
        frmPrn.ParamValue7 = GetFieldValue(Qry)
        Qry = "SELECT transignatory.EMPLOYEE From transignatory Where transignatory.TRANSACTION = 'Job Order Summary (2)';"
        frmPrn.ParamValue8 = GetFieldValue(Qry)
        
        frmPrn.MySource = "Job Order"
        frmPrn.Show
    ElseIf mySrc = "Sales Return" Then
        dCov = "Date Covered: " & Format(dFrom.Value, "Mmm. dd, yyyy") & IIf(Val(cmdNext.Tag) = 2, " to " & Format(dTo.Value, "Mmm. dd, yyyy"), "")
        If Val(cmdNext.Tag) = 1 Then
            '***Access
            'Qry = "SELECT SRETURNSP.RETDATE, SRETURNSP.RETNO, SRETURNSP.INVOICE, SPINVOICE.CHARGETO, SRETURNSP.REASON, SRETURNSP.TOTALAMT, SRETURNSP.REMARKS " & _
                  "FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO Where (((SRETURNSP.RETDATE) = #" & dFrom.Value & "#)) ORDER BY SRETURNSP.RETDATE, SRETURNSP.RETNO;"
            'mySQl
            'Qry = "SELECT SRETURNSP.RETDATE, SRETURNSP.RETNO, SRETURNSP.INVOICE, SPINVOICE.CHARGETO, SRETURNSP.REASON, SRETURNSP.TOTALAMT, SRETURNSP.REMARKS " & _
                  "FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO Where SRETURNSP.RETDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' ORDER BY SRETURNSP.RETDATE, SRETURNSP.RETNO;"
            Qry = "SELECT " & _
                "SRETURNSP.RETDATE, " & _
                "SRETURNSP.RETNO, " & _
                "SRETURNSP.INVOICE, " & _
                "SPINVOICE.INVDATE, " & _
                "SPINVOICE.CHARGETO, " & _
                "IF(MONTH(SRETURNSP.RETDATE) = MONTH(SPINVOICE.INVDATE), SRETURNSP.TOTALAMT, '') AS THMOS, " & _
                "IF(MONTH(SRETURNSP.RETDATE) = MONTH(SPINVOICE.INVDATE),'',SRETURNSP.TOTALAMT) AS PRMOS, " & _
                "SRETURNSP.TOTALAMT, " & _
                "SRETURNSP.REMARKS, " & _
                "SRETURNSP.PRINTCOPY, " & _
                "SRETURNSP.RETURNID " & _
            "From " & _
                "SRETURNSP " & _
                "INNER JOIN SPINVOICE ON (SRETURNSP.INVOICE = SPINVOICE.SPINVNO) " & _
            "Where " & _
                "SRETURNSP.RETDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' " & _
            "Order By " & _
                "SRETURNSP.RETDATE, " & _
                "SRETURNSP.RETNO Desc"

        Else
            '***Access
            'Qry = "SELECT SRETURNSP.RETDATE, SRETURNSP.RETNO, SRETURNSP.INVOICE, SPINVOICE.CHARGETO, SRETURNSP.REASON, SRETURNSP.TOTALAMT, SRETURNSP.REMARKS " & _
                  "FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO Where (((SRETURNSP.RETDATE) >= #" & dFrom.Value & "# And (SRETURNSP.RETDATE) <= #" & dTo.Value & "#)) ORDER BY SRETURNSP.RETDATE, SRETURNSP.RETNO;"
            'mySQl
            'Qry = "SELECT SRETURNSP.RETDATE, SRETURNSP.RETNO, SRETURNSP.INVOICE, SPINVOICE.CHARGETO, SRETURNSP.REASON, SRETURNSP.TOTALAMT, SRETURNSP.REMARKS " & _
                  "FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO Where SRETURNSP.RETDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And SRETURNSP.RETDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' ORDER BY SRET
            Qry = "SELECT " & _
                "SRETURNSP.RETDATE, " & _
                "SRETURNSP.RETNO, " & _
                "SRETURNSP.INVOICE, " & _
                "SPINVOICE.INVDATE, " & _
                "SPINVOICE.CHARGETO, " & _
                "IF(MONTH(SRETURNSP.RETDATE) = MONTH(SPINVOICE.INVDATE), SRETURNSP.TOTALAMT, '') AS THMOS, " & _
                "IF(MONTH(SRETURNSP.RETDATE) = MONTH(SPINVOICE.INVDATE),'',SRETURNSP.TOTALAMT) AS PRMOS, " & _
                "SRETURNSP.TOTALAMT, " & _
                "SRETURNSP.REMARKS, " & _
                "SRETURNSP.PRINTCOPY, " & _
                "SRETURNSP.RETURNID " & _
            "From " & _
                "SRETURNSP " & _
                "INNER JOIN SPINVOICE ON (SRETURNSP.INVOICE = SPINVOICE.SPINVNO) " & _
            "Where " & _
                "SRETURNSP.RETDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' AND " & _
                "SRETURNSP.RETDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' " & _
            "Order By " & _
                "SRETURNSP.RETDATE, " & _
                "SRETURNSP.RETNO Desc"
        End If
        PRNmc.QryPrn = Qry
        PRNmc.ParamValue1 = dCov
        '***Access
        'pQRY = "SELECT SETTINGS.VALUE From SETTINGS " & _
               "WHERE (((SETTINGS.DESCRIPTION)='EMP DEFAULT VALUE') AND ((SETTINGS.REMARKS)='OIC'));"
        'mySQl
        pQRY = "SELECT SETTINGS.VALUE From SETTINGS " & _
               "WHERE SETTINGS.DESCRIPTION='EMP DEFAULT VALUE' AND SETTINGS.REMARKS='OIC';"
        
        PRNmc.ParamValue2 = GetFieldValue(pQRY)
        '***Access
        'pQRY = "SELECT SETTINGS.VALUE From SETTINGS " & _
               "WHERE (((SETTINGS.DESCRIPTION)='EMP DEFAULT VALUE') AND ((SETTINGS.REMARKS)='MANAGER'));"
        'mySQL
        pQRY = "SELECT SETTINGS.VALUE From SETTINGS " & _
               "WHERE SETTINGS.DESCRIPTION='EMP DEFAULT VALUE' AND SETTINGS.REMARKS='MANAGER';"
        
        PRNmc.ParamValue3 = GetFieldValue(pQRY)
        PRNmc.ParamValue4 = "BRANCH MANAGER"
        PRNmc.ParamValue5 = "AREA OPERATION MANAGER"
        PRNmc.MySource = "SALESRETURNSUMMARY"
        PRNmc.Show
    
    ElseIf mySrc = "SP Incoming" Then
                If Val(cmdNext.Tag) = 2 Then
                    dCov = "Date Covered: " & dFrom.Value & " to " & dTo.Value
                Else
                    dCov = "Date Covered: " & dFrom.Value
                End If
                Qry = "call sp_mrcwarehouse.sp_incoming_summary('" & Val(cmdNext.Tag) & "','" & Format(dFrom.Value, "yyyy-mm-dd") & "','" & Format(dTo.Value, "yyyy-mm-dd") & "','" & pRECTYPE & "')"
                PRNmc.MySource = "SPRECEIVESUMM"
                PRNmc.QryPrn = Qry
                PRNmc.ParamValue1 = dCov
                PRNmc.Show
    ElseIf mySrc = "SP PACKING" Then
                    If dCode = 1 Then
                        'PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.TRANSFERID, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.TOTALCOST, STOCKTRANSFER.STATUS " & _
                              "From STOCKTRANSFER Where (((STOCKTRANSFER.TRANSDATE) = #" & dFrom.Value & "#)) ORDER BY STOCKTRANSFER.TRANSFERID;"
                        'mySQL            If dCode = 1 Then
                        PRNmc.QryPrn = "SELECT stocktransfer.TRANSDATE, stocktransfer.DATEFINAL, Case When ISNULL(stocktransfer.DELIVEREDDATE) THEN '' ELSE date_format(stocktransfer.DELIVEREDDATE, '%Y-%m-%d') END as DELIVEREDDATE, CAST(IFNULL(TRANSNO,TRANSFERID) as Char(60)) as REFERENCE, IFNULL(tr.TTLPN,0) as TTLPN, IFNULL(tr.TTLQTY,0) as TTLQTY, stocktransfer.RECBRANCH, (STOCKTRANSFER.TOTALCOST)*1.12 as TOTALCOSTING, STOCKTRANSFER.STATUS, STOCKTRANSFER.REMARKS " & _
                                        "From stocktransfer LEFT JOIN (Select COUNT(spitr.SPAREPARTID) as TTLPN, SUM(`OUT`) as TTLQTY, REPLACE(REPLACE(REPLACE(spitr.REFERENCE,'-MWH',''),'TR#',''),'STF#','') as REFERENCE From spitr Where spitr.TRANSTYPE = 'OUT' and spitr.REMARKS = 'TRANSFER' GROUP BY spitr.REFERENCE) as tr ON tr.REFERENCE = stocktransfer.TRANSFERID Where stocktransfer.TRANSDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' Order By stocktransfer.TRANSFERID DESC, stocktransfer.TRANSDATE Desc;"
                    Else
                        'PRNmc.QryPrn = "SELECT STOCKTRANSFER.TRANSDATE, STOCKTRANSFER.TRANSFERID, STOCKTRANSFER.RECBRANCH, STOCKTRANSFER.TOTALCOST, STOCKTRANSFER.STATUS " & _
                              "From STOCKTRANSFER Where STOCKTRANSFER.TRANSDATE >= #" & dFrom.Value & "# And STOCKTRANSFER.TRANSDATE <= #" & dTo.Value & "#  ORDER BY STOCKTRANSFER.TRANSFERID;"
                        'mySQL
                        PRNmc.QryPrn = "SELECT stocktransfer.TRANSDATE, stocktransfer.DATEFINAL, Case When ISNULL(stocktransfer.DELIVEREDDATE) THEN '' ELSE date_format(stocktransfer.DELIVEREDDATE, '%Y-%m-%d') END as DELIVEREDDATE, CAST(IFNULL(TRANSNO,TRANSFERID) as Char(60)) as REFERENCE, IFNULL(tr.TTLPN,0) as TTLPN, IFNULL(tr.TTLQTY,0) as TTLQTY, stocktransfer.RECBRANCH, (STOCKTRANSFER.TOTALCOST)*1.12 as TOTALCOSTING, STOCKTRANSFER.STATUS, STOCKTRANSFER.REMARKS " & _
                                       "From stocktransfer LEFT JOIN (Select COUNT(spitr.SPAREPARTID) as TTLPN, SUM(`OUT`) as TTLQTY, REPLACE(REPLACE(REPLACE(spitr.REFERENCE,'-MWH',''),'TR#',''),'STF#','') as REFERENCE From spitr Where spitr.TRANSTYPE = 'OUT' and spitr.REMARKS = 'TRANSFER' GROUP BY spitr.REFERENCE) as tr ON tr.REFERENCE = stocktransfer.TRANSFERID Where stocktransfer.TRANSDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' AND stocktransfer.TRANSDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' Order By stocktransfer.TRANSFERID DESC, stocktransfer.TRANSDATE Desc;"
                        
                    End If
                PRNmc.MySource = "TRANSSUMMARY"
                Qry = "SELECT BRANCH.BRANCHNAME From BRANCH WHERE (((BRANCH.SYSTEMBRANCH)=True));"
                PRNmc.ParamValue1 = GetFieldValue(Qry)
                Qry = "SELECT BRANCH.ADDRESS From BRANCH WHERE (((BRANCH.SYSTEMBRANCH)=True));"
                PRNmc.ParamValue2 = GetFieldValue(Qry)
                PRNmc.ParamValue3 = "For the period " & dFrom.Value & " to " & dTo.Value
                PRNmc.Show
                Me.Visible = False
                
   ElseIf mySrc = "SRP Pricing" Then
    '*********************** CODE UPDATED BY IAN: 07/21/2018 ***************************
            Screen.MousePointer = vbHourglass
            Dim rptSRP As New rptSRPSum
        
        
            sql = "Call sp_srppricing.sp_srppricing_summary('" & Format(dFrom.Value, "yyyy-mm-dd") & "','" & Format(dTo.Value, "yyyy-mm-dd") & "');"
            'InputBox sql, sql, sql
            
            
            If tmpRs.State <> 0 Then tmpRs.Close
            Set tmpRs = LYRcon.Execute(sql)
            DoEvents
            rptSRP.DiscardSavedData
            DoEvents
        
            Call rptSRP.Database.Tables(1).SetDataSource(tmpRs)
        
            frmPrint.CRViewer1.ReportSource = rptSRP
            frmPrint.CRViewer1.ViewReport
            frmPrint.MySource = "SRP Updater Summary"
            frmPrint.CRViewer1.Zoom 150
            frmPrint.CRViewer1.EnableExportButton = True
            Screen.MousePointer = vbDefault
            DoEvents
            frmPrint.Caption = "PRINT" & " " & "-" & " " & "[" & "SUMMARY" & "]"
            frmPrint.WindowState = vbMaximized
            frmPrint.Show vbModal
 
    ElseIf mySrc = "Costing" Then
            
            Screen.MousePointer = vbHourglass
            Dim rpt As New rptSummary
        
            'sp_acctpayinvcng.sp_acctpayinvcng_summary(
            '   pDATEFROM       VarChar(60)
            '   ,pDATETO        Varchar(60)
            '   ,pSTATUS        Varchar(10)
            ')
        
            sql = "Call sp_acctpayinvcng.sp_acctpayinvcng_summary('" & Format(dFrom.Value, "yyyy-mm-dd") & "','" & Format(dTo.Value, "yyyy-mm-dd") & "','');"
            'InputBox sql, sql, sql
            
            
            If tmpRs.State <> 0 Then tmpRs.Close
            Set tmpRs = LYRcon.Execute(sql)
            DoEvents
            rpt.DiscardSavedData
            DoEvents
        
            Call rpt.Database.Tables(1).SetDataSource(tmpRs)
        
            frmPrint.CRViewer1.ReportSource = rpt
            frmPrint.CRViewer1.ViewReport
            frmPrint.MySource = "Costing Summary"
            frmPrint.CRViewer1.Zoom 150
            frmPrint.CRViewer1.EnableExportButton = True
            Screen.MousePointer = vbDefault
            DoEvents
            frmPrint.Caption = "PRINT" & " " & "-" & " " & "[" & "SUMMARY" & "]"
            frmPrint.WindowState = vbMaximized
            frmPrint.Show vbModal
'*********************** CODE UPDATED: 03/20/2018 ***************************
    ElseIf mySrc = "Branch Purchase Order" Then
        If CodeDateRange = 1 Then
            frmPrint.QryPrn = "SELECT sporderbranch.BRANCHNAME, sporderbranch.REFERENCE, sporderbranch.ORDERDATE, sporderbranch.`STATUS`, sporderbranch.QTYPN, sporderbranch.QTYORDER, " & _
                        "sporderbranch.LASTUPDATE,sporderbranch.BRANCHSPOID,sporderbranch.BRANCHID, sporderbranch.PICKEDBY, sporderbranch.PRINTPICK,sporderbranch.SPORDERID" & _
                        "From sporderbranch Where sporderbranch.ORDERDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' ORDER BY sporderbranch.SPORDERID DESC;"
        Else
            frmPrint.QryPrn = "SELECT sporderbranch.BRANCHNAME, sporderbranch.REFERENCE, sporderbranch.ORDERDATE, sporderbranch.`STATUS`, sporderbranch.QTYPN, sporderbranch.QTYORDER, " & _
                        "sporderbranch.LASTUPDATE,sporderbranch.BRANCHSPOID,sporderbranch.BRANCHID, sporderbranch.PICKEDBY, sporderbranch.PRINTPICK,sporderbranch.SPORDERID " & _
                        "From sporderbranch Where sporderbranch.ORDERDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' and sporderbranch.ORDERDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' ORDER BY sporderbranch.SPORDERID DESC;"
        End If
        frmPrint.MySource = "Branch Purchase Order"
        frmPrint.ParamValue1 = SysBranch
        frmPrint.ParamValue2 = SysAddress
        frmPrint.ParamValue3 = dFrom.Value & " TO " & dTo.Value
        frmPrint.Show
'******************* END CODE UPDATED: 03/20/2018 ***************************
    ElseIf mySrc = "RFP RECONCILIATION" Then

        Qry = "Call sp_acctpayinvcng.sp_rfp_reconciliation_summ(" & Val(CodeDateRange) & ",'" & Format(dFrom.Value, "yyyy-mm-dd") & "','" & Format(dTo.Value, "yyyy-mm-dd") & "');"

        frmPrint.MySource = "RFP RECONCILIATION"
        frmPrint.QryPrn = Qry
        frmPrint.ParamValue1 = dFrom.Value & " TO " & dTo.Value
        frmPrint.ParamValue2 = SysBranch
        frmPrint.ParamValue3 = SysAddress
        frmPrint.Show
'*********************** CODE UPDATED: 05/06/2020 ***************************
    Else
        'MsgBox "Not yet available", vbInformation + vbOKOnly, "Print"
        Call ExtractRFP
    End If

End Sub

Private Sub cmdRefresh_Click()
On Error GoTo er

    Screen.MousePointer = vbHourglass
    tmrRefresh.Enabled = True

    DoEvents
    Call Load(frm_OnProcess2)
    frm_OnProcess2.Show vbModal
    
Exit Sub
er:
    tmrRefresh.Enabled = False
    Screen.MousePointer = vbDefault
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdRefresh_Click()")
End Sub

Private Sub cmdSalesDrive_Click()

    If mySrc = "DAILY P.R." Then
        If Val(cmdNext.Tag) = 1 Then
            Qry = "SELECT dailypr.ACCOUNTNO, concat(customer.LASTNAME, ', ', customer.FIRSTNAME, ' ', MID(customer.MIDDLENAME, 1, 1),'.') AS CUSTNAME, unitordered.MODEL, SUM(dailypr.PRAMOUNT) As Total " & _
                 "From dailypr INNER JOIN unitordered ON (dailypr.ORDERID = unitordered.ORDERID) INNER JOIN customer ON (unitordered.CUSTOMERID = customer.CUSTOMERID) " & _
                 "WHERE dailypr.PRDATE = '" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND dailypr.REMARKS <> 'CANCELLED' Group By dailypr.ACCOUNTNO, concat(customer.LASTNAME, ', ', customer.FIRSTNAME, ' ', MID(customer.MIDDLENAME, 1, 1),'.'), unitordered.MODEL Order By CustName;"
            dCov = "Date Covered: " & Format(dFrom.Value, " MMMM DD, YYYY")
        
        Else
            Qry = "SELECT dailypr.ACCOUNTNO, concat(customer.LASTNAME, ', ', customer.FIRSTNAME, ' ', MID(customer.MIDDLENAME, 1, 1),'.') AS CUSTNAME, unitordered.MODEL, SUM(dailypr.PRAMOUNT) As Total " & _
                 "From dailypr INNER JOIN unitordered ON (dailypr.ORDERID = unitordered.ORDERID) INNER JOIN customer ON (unitordered.CUSTOMERID = customer.CUSTOMERID) " & _
                 "WHERE dailypr.PRDATE >= '" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND dailypr.PRDATE <= '" & Format(dTo.Value, "YYYY-MM-DD") & "' AND dailypr.REMARKS <> 'CANCELLED' Group By dailypr.ACCOUNTNO, concat(customer.LASTNAME, ', ', customer.FIRSTNAME, ' ', MID(customer.MIDDLENAME, 1, 1),'.'), unitordered.MODEL Order By CustName;"
                dCov = "Date Covered: FROM " & Format(dFrom.Value, "MMMM DD, YYYY") & " to " & Format(dTo.Value, "MMMM DD, yyyy")
        End If
        frmPrn.QryPrn = Qry
        Qry = "SELECT settings.REMARKS From settings Where settings.DESCRIPTION = 'SYSTEM BRANCH';"
        frmPrn.ParamValue1 = GetFieldValue(Qry)
        Qry = "SELECT BRANCH.ADDRESS From Branch WHERE (((BRANCH.BRANCHID)=" & Val(Marketing.Tag) & "));"
        frmPrn.ParamValue2 = GetFieldValue(Qry)
        frmPrn.ParamValue3 = dCov
        Qry = "SELECT concat(employees.FIRSTNAME,' ',employees.LASTNAME) as PrepBy From employees Where employees.EmpId = " & UserID & ";"
        frmPrn.ParamValue4 = GetFieldValue(Qry)
        Qry = "SELECT BRANCH.MANAGER From Branch WHERE (((BRANCH.BRANCHID)=" & Val(Marketing.Tag) & "));"
        frmPrn.ParamValue5 = GetFieldValue(Qry)
        Qry = "SELECT POSITION From employees Where employees.EmpId = " & UserID & ";"
        frmPrn.ParamValue6 = GetFieldValue(Qry)
        frmPrn.MySource = "DAILY SUMMARY"
        frmPrn.Show
    ElseIf mySrc = "Daily Cash Collection" Then
        If Val(cmdNext.Tag) = 1 Then
            Qry = "SELECT CONCAT(employees.LASTNAME, ', ', employees.FIRSTNAME, ' ', MID(employees.MIDDLENAME, 1, 1),'.') AS Collector, " & _
                  "payments.PAYDATE, payments.CUSTOMERNAME, payments.ACCOUNTNO, IF(payments.ORNO <> '', payments.ORNO, payments.CRNO) AS Receipt, payments.CASH + payments.DOWNPAYMENT + payments.INSTALLMENT + payments.OTHERS AS ORAMOUNT " & _
                  "From payments INNER JOIN employees ON (payments.EMPID = employees.EMPID) Where payments.PAYDATE = '" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND employees.`POSITION` <> 'CASHIER' Order By employees.LASTNAME, payments.PAYDATE, payments.CUSTOMERNAME"
        Else
            Qry = "SELECT CONCAT(employees.LASTNAME, ', ', employees.FIRSTNAME, ' ', MID(employees.MIDDLENAME, 1, 1),'.') AS Collector, " & _
                  "payments.PAYDATE, payments.CUSTOMERNAME, payments.ACCOUNTNO, IF(payments.ORNO <> '', payments.ORNO, payments.CRNO) AS Receipt, payments.CASH + payments.DOWNPAYMENT + payments.INSTALLMENT + payments.OTHERS AS ORAMOUNT " & _
                  "From payments INNER JOIN employees ON (payments.EMPID = employees.EMPID) Where payments.PAYDATE >= '" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND payments.PAYDATE <= '" & Format(dTo.Value, "YYYY-MM-DD") & "' AND employees.`POSITION` <> 'CASHIER' Order By employees.LASTNAME, payments.PAYDATE, payments.CUSTOMERNAME"
        End If
        frmPrn.QryPrn = Qry
        Qry = "SELECT settings.REMARKS From settings Where settings.DESCRIPTION = 'SYSTEM BRANCH';"
        frmPrn.ParamValue1 = GetFieldValue(Qry)
        Qry = "SELECT BRANCH.ADDRESS From Branch WHERE (((BRANCH.BRANCHID)=" & Val(Marketing.Tag) & "));"
        frmPrn.ParamValue2 = GetFieldValue(Qry)
        frmPrn.MySource = "COLLECTOR COLLECTION"
        frmPrn.Show
    End If
End Sub

Private Sub cmdShow_Click()
    
    With flxSum
        If .Rows = 2 And .TextMatrix(1, 1) = "" Then
            MsgBox "There is no record to view!", vbInformation + vbOKOnly, "Confirm"
            Exit Sub
        End If
        If mySrc = "SP Incoming" Then
            If pRECTYPE = "WAREHOUSE" Then
                Call SPincoming.ShowRecord(.TextMatrix(ActRow, 12))
            Else
                Call SPincomingWS.ShowRecord(.TextMatrix(ActRow, 12))
            End If
            Unload Me
        ElseIf mySrc = "Advances for Liquidation" Then
            ExpAdvLiq.SrcSumm = True
            ExpAdvLiq.txtAFLid.Text = .TextMatrix(ActRow, 9)
            Call ExpAdvLiq.ShowRecord
            ExpAdvLiq.Show vbModal
        ElseIf mySrc = "Delivery Receipt" Then
            Call Outgoing.ShowRecord(.TextMatrix(ActRow, 9))
            Unload Me
        ElseIf mySrc = "Item Entry" Then
            Call frmItem.ShowRecord(.TextMatrix(ActRow, 9))
            frmItem.Show vbModal
        ElseIf mySrc = "Supplier Entry" Then
            Call frmSupplier.ShowRecord(.TextMatrix(ActRow, 7))
            frmSupplier.Show vbModal
        ElseIf mySrc = "Branch Entry" Then
            Call frmBranch.ShowRecord(.TextMatrix(ActRow, 7))
            frmBranch.Show vbModal
        
        ElseIf mySrc = "Transfer Voucher" Then
            If .TextMatrix(ActRow, 9) = "MC" Then
                Call InvTransferMC.ShowRecord(.TextMatrix(ActRow, 10))
            Else
                Call InvTransfer.ShowRecord(.TextMatrix(ActRow, 10))
            End If
            Unload Me
        ElseIf mySrc = "Spare Parts Receiving" Then
            If pRECTYPE = "WAREHOUSE" Then
                Call SPincoming.ShowRecord(Val(.TextMatrix(ActRow, 9)))
            Else
                Call SPincomingWS.ShowRecord(Val(.TextMatrix(ActRow, 9)))
            End If
            Unload Me
            
        ElseIf mySrc = "Spareparts Invoice" Then
            Call SPInvoice.ShowRecord(Val(.TextMatrix(ActRow, 11)))
            Unload Me
        ElseIf mySrc = "Spareparts Location Adj." Then
            Call SPLocAdj.DisplayLocAdj(Val(.TextMatrix(ActRow, 8)))
            Unload Me
        ElseIf mySrc = "Branch Order" Then
            Call SPordering.DisplaySPOrder(.TextMatrix(ActRow, 2))
            Unload Me
        ElseIf mySrc = "Branch Purchase Order" Then
            cmdUpdate.Visible = True
            If flxSum.TextMatrix(ActRow, 4) = "NEW" Then
               
                ans = MsgBox("Do you want to proceed this action?", vbQuestion + vbYesNo, "Show")
                If ans = vbNo Then Exit Sub
'                If Left(.TextMatrix(ActRow, 2), 1) = "R" Then
'                    Call BranchReplOrder.ShowRequest(.TextMatrix(ActRow, 2), .TextMatrix(ActRow, 11))
'                Else
                    Call savePOdetails(flxSum.TextMatrix(ActRow, 13))
                    Call BranchPO.ShowRequest(.TextMatrix(ActRow, 2), .TextMatrix(ActRow, 12))
'                End If
            Else
'                If Left(.TextMatrix(ActRow, 2), 1) = "R" Then
'                    Call BranchReplOrder.ShowRequest(.TextMatrix(ActRow, 2), .TextMatrix(ActRow, 11))
'                Else
                    Call BranchPO.ShowRequest(.TextMatrix(ActRow, 2), .TextMatrix(ActRow, 12))
'                End If
            
            End If
            Unload Me
        ElseIf mySrc = "SP PACKING" Then
            'Call frmSPtrans.ShowRecord(.TextMatrix(ActRow, 4), .TextMatrix(ActRow, 3))
            Call frmPOReplStockTrans.ShowRecord(.TextMatrix(ActRow, 4), .TextMatrix(ActRow, 3))
            Unload Me
        ElseIf mySrc = "AllocBBO" Then
            Call AllocBBO.ShowRecord(.TextMatrix(ActRow, 1))
            Unload Me
        ElseIf mySrc = "Sales Return" Then
            'Debug
            Call SalesReturn.DisplaySalesReturn(flxSum.TextMatrix(ActRow, 11))
            Unload Me
        ElseIf mySrc = "Purchase Return" Then
             Call PurchaseReturn.DisplayPurchaseReturn(flxSum.TextMatrix(ActRow, 9))
             Unload Me
        ElseIf mySrc = "Cashier Cash Flow" Then
            Call CashFlowCashier.ShowRecord(Val(.TextMatrix(ActRow, 5)))
            Unload Me
        ElseIf mySrc = "Inventory Adjustment" Then
            Call InvntryADJ.DisplayInvAdj(flxSum.TextMatrix(ActRow, 2), flxSum.TextMatrix(ActRow, 9))
            Unload Me
'************************* CODE UPDATED: 02/13/2019 ***********************
        ElseIf mySrc = "WS Inventory Adjustment" Then
            Call WSInvntryADJ.DisplayInvAdj(flxSum.TextMatrix(ActRow, 2), flxSum.TextMatrix(ActRow, 4))
            Unload Me
'**********************END CODE UPDATED: 02/13/2019 ***********************
        ElseIf mySrc = "Costing" Then
            'Call frmAcctPayInvcg.GetRecord(flxSum.TextMatrix(ActRow, 4))
            frmAcctPayInvcg.txtPckngNo.Text = flxSum.TextMatrix(ActRow, 4)
            frmAcctPayInvcg.tmrGetRecord.Enabled = True
            
            DoEvents
            Call Load(frm_OnProcess2)
            frm_OnProcess2.Show vbModal
            Unload Me
            
        ElseIf mySrc = "SRP Pricing" Then
            Dim filter As String
            
            If flxSum.TextMatrix(ActRow, 7) = "MANUAL" Then
                filter = "ME"
            Else
                filter = ""
            End If
            
            frmSRPUpdater.filter = filter
            frmSRPUpdater.formAction = "SUMMARY"
            frmSRPUpdater.SupCode = flxSum.TextMatrix(ActRow, 8)
            frmSRPUpdater.txtPckngNo.Text = flxSum.TextMatrix(ActRow, 3)
            frmSRPUpdater.tmrShowRecord.Enabled = True
            
            DoEvents
            Call Load(frm_OnProcess2)
            frm_OnProcess2.Show vbModal
            Unload Me
            
'******************** CODE UPDATED BY IAN: 03/05/18 ******************************
        ElseIf mySrc = "SRP Request form" Then
            frmSRPREQ.ShowApprovedRecord (flxSum.TextMatrix(ActRow, 1))
            Unload Me
'******************** END CODE UPDATED BY IAN: 03/05/18 ******************************
'********************* CODE UPDATED: 08/14/2017 ******************************
        ElseIf mySrc = "SETTLEMENT DISCREPANCY" Then
            chkRRNO = GetFieldValue("SELECT PACKINGNO FROM SPRECEIVE WHERE PACKINGNO = '" & flxSum.TextMatrix(ActRow, 3) & "'")
            If chkRRNO = "" Then
                tmpSTFNo = Replace(flxSum.TextMatrix(ActRow, 3), "-MWH", "")
                chkSTFNO = GetFieldValue("SELECT TRANSFERID FROM STOCKTRANSFER WHERE TRANSFERID = '" & tmpSTFNo & "'")
                If chkSTFNO <> "" Then
                    Call frmSIMSDiscrepancy.ShowRecord(Val(flxSum.TextMatrix(ActRow, 13)), flxSum.TextMatrix(ActRow, 11), flxSum.TextMatrix(ActRow, 3), flxSum.TextMatrix(ActRow, 14))
                Else
                    MsgBox "Cannot Process Settlement of Discrepancy", vbInformation + vbOKOnly
                    Exit Sub
                End If
            Else
                Call frmSIMSDiscrepancy.ShowRecord(Val(flxSum.TextMatrix(ActRow, 13)), flxSum.TextMatrix(ActRow, 11), flxSum.TextMatrix(ActRow, 3), flxSum.TextMatrix(ActRow, 14))
            End If
            Unload Me
        ElseIf mySrc = "SDiscrepancy Adjustment" Then
            Call SDiscInvtryAdj.DisplayInvAdj(flxSum.TextMatrix(ActRow, 2))
            Unload Me
'***************** END CODE UPDATED: 08/14/2017 ******************************
'********************* CODE UPDATED: 06/14/2019 ******************************
        ElseIf mySrc = "PO ADJUSTMENT" Then
            Call frmPOAdj.ShowRecord(flxSum.TextMatrix(ActRow, 3))
            Unload Me
'***************** END CODE UPDATED: 06/14/2019 ******************************
'********************* CODE UPDATED: 10/04/2019 ******************************
        ElseIf mySrc = "RFP" Then
            Call frmRFPWH.ShowRecord(flxSum.TextMatrix(ActRow, 2))
            Unload Me
'***************** END CODE UPDATED: 10/04/2019 ******************************
        ElseIf mySrc = "ADJUSTMENT REQUEST" Then
            Call frmOtherAdjustment.ShowRecord(flxSum.TextMatrix(ActRow, 3))
            Unload Me
        ElseIf mySrc = "CANCELLATION MODULE" Then
            Call frmCancellation.ShowRecord(flxSum.TextMatrix(ActRow, 2))
            frmCancellation.Show vbModal
        End If
    End With

End Sub

Function GetOR(Qry As String) As String
    Dim rsCashOR As New ADODB.Recordset
    Dim MyBig As String, ORseries As String, LastSer As String
    Dim x As Integer
    
    With rsCashOR
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            ORseries = ""
            x = 1
            Do While Not .EOF
               If x = 1 Then
                    MyBig = .Fields(0)
                    ORseries = .Fields(0)
                Else
                    If Val(.Fields(0)) = Val(MyBig) + 1 Then
                        MyBig = Val(.Fields(0))
                        If Right$(ORseries, 1) <> "-" Then
                            ORseries = ORseries & "-"
                            LastSer = .Fields(0)
                        End If
                    Else
                        If Right$(ORseries, 1) = "-" Then
                            ORseries = ORseries & IIf(Left$(LastSer, 3) <> Left$(Format$(MyBig, "0000"), 3), Format(MyBig, "0000"), Right$(MyBig, 2)) & ", " & .Fields(0)
                            MyBig = .Fields(0)
                        Else
                            MyBig = .Fields(0)
                            ORseries = ORseries & ", " & .Fields(0)
                        End If
                    End If
                End If
                x = x + 1
                .MoveNext
            Loop
            If Right$(ORseries, 1) = "-" Then ORseries = ORseries & MyBig
        End If
        .Close
    End With
    GetOR = ORseries
End Function


Private Sub cmdTransmittal_Click()
    Dim dCov As String, Qry As String
    
    
    If mySrc = "Sales Invoice" Then
        If Val(cmdNext.Tag) = 1 Then
            Qry = "SELECT date_format(salesinvoice.INVDATE,'%c/%d/%Y'), concat(customer.LASTNAME, ', ', customer.FIRSTNAME, ' ', MID(customer.MIDDLENAME, 1, 1),'.') AS CustName, concat(productline.BRAND, ' ', productline.MODEL, ' ', productline.CLASSIFICATION) AS UnitDesc, " & _
                  "salesinvoice.INVOICENO, salesinvoice.DRNUMBER, if(salesinvoice.REMARKS<>'CANCELLED','','CANCELLED') AS MYREM " & _
                  "From SalesInvoice INNER JOIN unitordered ON (salesinvoice.ORDERID = unitordered.ORDERID) INNER JOIN productline ON (unitordered.PRODID = productline.PRODID) INNER JOIN customer ON (unitordered.CUSTOMERID = customer.CUSTOMERID) " & _
                  "WHERE salesinvoice.INVDATE = '" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND salesinvoice.PRNCOPY > 0 Order By salesinvoice.INVDATE DESC, SalesInvoice.DRNUMBER Desc;"
            dCov = "Date covered: " & Format(dFrom.Value, " MMMM DD, YYYY")
        
        Else
            Qry = "SELECT date_format(salesinvoice.INVDATE,'%c/%d/%Y'), concat(customer.LASTNAME, ', ', customer.FIRSTNAME, ' ', MID(customer.MIDDLENAME, 1, 1),'.') AS CustName, concat(productline.BRAND, ' ', productline.MODEL, ' ', productline.CLASSIFICATION) AS UnitDesc, " & _
                  "salesinvoice.INVOICENO, salesinvoice.DRNUMBER, if(salesinvoice.REMARKS<>'CANCELLED','','CANCELLED') AS MYREM " & _
                  "From SalesInvoice INNER JOIN unitordered ON (salesinvoice.ORDERID = unitordered.ORDERID) INNER JOIN productline ON (unitordered.PRODID = productline.PRODID) INNER JOIN customer ON (unitordered.CUSTOMERID = customer.CUSTOMERID) " & _
                  "WHERE salesinvoice.INVDATE >= '" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND salesinvoice.INVDATE <= '" & Format(dTo.Value, "YYYY-MM-DD") & "' AND salesinvoice.PRNCOPY > 0 Order By salesinvoice.INVDATE DESC, SalesInvoice.DRNUMBER Desc;"
            dCov = "For the period " & Format(dFrom.Value, "MMMM DD, YYYY") & " to " & Format(dTo.Value, "MMMM DD, yyyy")
        End If
        frmPrn.QryPrn = Qry
        Qry = "SELECT settings.REMARKS From settings Where settings.DESCRIPTION = 'SYSTEM BRANCH';"
        frmPrn.ParamValue1 = GetFieldValue(Qry)
        Qry = "SELECT BRANCH.ADDRESS From Branch WHERE (((BRANCH.BRANCHID)=" & Val(Marketing.Tag) & "));"
        frmPrn.ParamValue2 = GetFieldValue(Qry)
        frmPrn.ParamValue3 = dCov
        Qry = "SELECT concat(employees.FIRSTNAME,' ',employees.LASTNAME) as PrepBy From employees Where employees.EmpId = " & UserID & ";"
        frmPrn.ParamValue4 = GetFieldValue(Qry)
        Qry = "SELECT BRANCH.MANAGER From Branch WHERE (((BRANCH.BRANCHID)=" & Val(Marketing.Tag) & "));"
        frmPrn.ParamValue5 = GetFieldValue(Qry)
        Qry = "SELECT POSITION From employees Where employees.EmpId = " & UserID & ";"
        frmPrn.ParamValue6 = GetFieldValue(Qry)
        frmPrn.MySource = "SIDR TRANSMITTAL"
        frmPrn.Show
    End If
End Sub

'Private Sub cmdUpdate_Click()
'On Error GoTo LocalErr
'    Dim rsBranch As New Recordset
'    Dim branchID As String
'    Dim rsBranchPO As New ADODB.Recordset
'    Dim Qry As String
'    Dim MxMdl As Integer, i As Integer, j As Integer, InvRow As Integer, PbV As Integer, x As Integer
'    Dim BlnkInv As Boolean
'
'    flxBranch.Cols = bCol
'
'    With rsBranch
'        Qry = "SELECT sporderbranch.BRANCHID, sporderbranch.BRANCHNAME, MAX(sporderbranch.BRANCHSPOID) As MAXSPOID " & _
'                "From sporderbranch Group By sporderbranch.BRANCHID Order By sporderbranch.BRANCHID;"
'        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
'            flxBranch.Rows = 1
'            flxBranch.Rows = 2
'            flxBranch.row = 1
'            flxBranch.Col = 1
'            If .RecordCount > 0 Then
'                .MoveFirst
'                Do While Not .EOF
'                    flxBranch.ColWidth(0) = 300
'                    flxBranch.AddItem vbTab & !branchID & vbTab & !branchName & vbTab & !maxSPOID
'                    .MoveNext
'                Loop
'            End If
'        .Close
'    End With
'
'    For x = 2 To flxBranch.Rows - 1
'        branchID = Val(flxBranch.TextMatrix(x, 1))
'
'        pBar.Value = 0
'        pBar.Max = 100
'        txtBID.Text = ""
'        txtSPOID.Text = ""
'        txtbName.Text = ""
'        txtSPOID.Text = Val(flxBranch.TextMatrix(x, 3))
'        txtbName.Text = flxBranch.TextMatrix(x, 2)
'
'        Call ConnectOnline(branchID)
'        With rsBranchPO
'            txtBID.Text = branchID
''                If txtBID.Text = branchID Then
''                    With flxBranch
''                        For i = fRow To .Rows - 1
''                            For j = fCol To bCol - 3
''                                If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
''                                    .Col = 1: .row = i
''                                    .ColSel = bCol - 1: .RowSel = i
''                                    .TopRow = i
''                                    fRow = IIf(j = bCol - 1, i + 1, i)
''                                    fCol = IIf(j = bCol - 1, 1, j + 1)
''                                    aRow = i
''                                End If
''                            Next j
''                            fCol = 1
''                        Next i
''                        txtSPOID.Text = .TextMatrix(aRow, 3)
''                        txtbName.Text = .TextMatrix(aRow, 2)
''                    End With
''                    fRow = 1
''                    fCol = 1
''                End If
'                Qry = "SELECT sporder.REFERENCE, sporder.ORDERDATE, sporder.`STATUS`, sporder.QTYPN, sporder.QTYORDER, sporder.SPORDERID, sporder.CATEGORY,sporder.ORDERTYPE " & _
'                      "From sporder Where sporder.SPORDERID > " & txtSPOID.Text & " AND sporder.DEALERCODE <> 'MWA';"
'                .Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
'
'                If .RecordCount > 0 Then
'                    .MoveFirst
'                    flxOrder.Rows = 1
'                    MxMdl = .RecordCount
'                    If MxMdl > 10 Then PbV = 1 Else PbV = 100 \ MxMdl
'                    pBar.Visible = True
'                    Do While Not .EOF
'                        flxOrder.AddItem txtbName.Text & vbTab & txtBID.Text & vbTab & !REFERENCE & vbTab & !ORDERDATE & vbTab & !Status & vbTab & !QTYPN & vbTab & !QtyOrder & vbTab & !SPORDERID & vbTab & !Category & vbTab & !ORDERTYPE
'                        .MoveNext
'                        If pBar.Value < 100 Then
'                            pBar.Value = pBar.Value + PbV
'                        End If
'                        DoEvents
'                    Loop
'                    With flxOrder
'                        For i = 1 To .Rows - 1
'
'                            LYRcon.Execute "INSERT INTO sporderbranch(BRANCHSPOID, BRANCHNAME, REFERENCE, ORDERDATE, STATUS, QTYPN, QTYORDER, BRANCHID ,CATEGORY,ORDERTYPE)" & _
'                                    "VALUES(" & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 2) & "', '" & Format(.TextMatrix(i, 3), "yyyy-mm-dd") & "', " & _
'                                    "'" & .TextMatrix(i, 4) & "', " & .TextMatrix(i, 5) & ", " & .TextMatrix(i, 6) & ", " & .TextMatrix(i, 1) & ", '" & .TextMatrix(i, 8) & "', '" & .TextMatrix(i, 9) & "');"
'                        Next i
'                    End With
'                Else
'                    MsgBox "No available latest P.O. for " & txtbName.Text & "!", vbExclamation + vbOKOnly
'                    If pBar.Value > 0 Then
'                        flxOrder.Rows = 1
'                        flxOrder.Rows = 2
'                        pBar.Value = 1000
'                        DoEvents
'                    Else
'                        pBar.Value = 0
'                    End If
'                End If
'            .Close
'        End With
'        txtBID.Text = ""
'        txtSPOID.Text = ""
'        txtbName.Text = ""
'        OnlineCon.Close
'    Next x
'
'    Call RefreshRecord(1)
'    pBar.Visible = False
'    Exit Sub
'
'LocalErr:
'    MsgBox Err.Description
'End Sub

Public Sub branchID()
Dim i As Integer

    With flxBranch

    .SelectionMode = flexSelectionByRow

        For i = 1 To .Rows - 1

            If InStr(1, .TextMatrix(i, 0), txtBID.Text) Then

                .row = i

                .TopRow = i

                .ColSel = .Cols - 1

                .SelectionMode = flexSelectionFree

                Exit Sub

            End If

        Next i

    End With

End Sub

Public Sub findID()
    Dim i As Integer, j As Integer, CCol As Integer
        
    If txtBID.Text = "" Then Exit Sub
    With flxBranch
        
        .Cols = 4
        For i = fRow To .Rows - 1
            For j = fCol To CCol - 1
                If InStr(.TextMatrix(i, j), txtBID.Text) > 0 Then
                    .col = 1: .row = i
                    .ColSel = CCol - 1: .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = CCol - 1, i + 1, i)
                    fCol = IIf(j = CCol - 1, 1, j + 1)
                    ActRow = i
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

Private Sub flxSum_EnterCell()
    ActRow = flxSum.row
    
End Sub

Private Sub flxSum_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'Identify transactions without cancel features and exit procedure
    If flxSum.TextMatrix(ActRow, 1) = "" Or mySrc = "SP Incoming" Or mySrc = "Product" Or mySrc = "Price List" Or mySrc = "Price List Class B" Or mySrc = "Cashier Cash Flow" Or mySrc = "Deposited Unit" Or mySrc = "Redeemed Unit" Then Exit Sub
    If Not IsDate(flxSum.TextMatrix(ActRow, 1)) Then
        Exit Sub
    End If
    If Button = vbRightButton Then
        If mySrc = "Daily Cash Collection" Then
            If InStr(UCase$(flxSum.TextMatrix(ActRow, 3)), "CANCELLED") > 0 Then
                Call PopupMenu(mnuCancelReason)
                
            Else
                mnuCancel.Caption = "Cancel O.R."
                Call PopupMenu(mnuSum)
            
            End If
        ElseIf mySrc = "Purchase Return" Then
            Call PopupMenu(mnuPurchRet)
        ElseIf mySrc = "Spareparts Invoice" Then
            If Val(flxSum.TextMatrix(ActRow, 10)) > 0 Then
                mnuCancel.Caption = "Cancel " & mySrc
                Call PopupMenu(mnuSum)
            End If
        ElseIf mySrc = "RFP" Then
            If flxSum.TextMatrix(ActRow, 2) <> "" Then
                Call PopupMenu(mnuRFP)
            End If
        ElseIf mySrc = "RFP RECONCILIATION" Then
            Exit Sub
        Else
            mnuCancel.Caption = "Cancel " & mySrc
'************CODE UPDATED: 08/04/2017****************
            mnuConfirm.Caption = "Confirm " & mySrc
'********END CODE UPDATED: 08/04/2017****************
            Call PopupMenu(mnuSum)
        End If
    End If
End Sub

Private Sub Form_Load()
On Error GoTo er

    DoEvents
    tmrRefresh.Enabled = False

    tmrAutoLoad.Enabled = True
    
    DoEvents
    Call Load(frm_OnProcess2)
    frm_OnProcess2.Show vbModal
    
    If mySrc = "RFP" Then
        cmdPrint.Caption = "Excel Extract"
    End If
    
Exit Sub
er:
    tmrAutoLoad.Enabled = False
    Call modObjTools.ErrorMessage(Err.Description, "Error!->Form_Load()")
    

End Sub



'********************** CODE UPDATED: 08/11/17 **********************************
Private Sub mnuconfirm_Click()
If mySrc <> "Inventory Adjustment" Then
    SSConfirmTab.Visible = True
    txtConfirmDate.Text = Format(Now, "mm/dd/yyyy")
    With SSConfirmTab
        .Top = (Me.Height - .Height) \ 2
        .Left = (Me.Width - .Width) \ 2
        .Visible = True
        .ZOrder
    End With
Else
    MsgBox "NOT AVAILABLE!", vbInformation + vbOKOnly, "VALIDATION"
End If
End Sub
'****************** END CODE UPDATED: 08/11/17 **********************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    Marketing.sBar.Panels(6).Text = ""
    If rsSummary.State = adStateOpen Then rsSummary.Close
    'LYRcon.Open
End Sub

Private Sub mnuCancel_Click()
    Dim Qry As String, MyAcc As String
    
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=" & AccID & ";"
    MyAcc = GetFieldValue(Qry)
'********** CODE UPDATE 04/28/2018 **************
    If Val(MyAcc) <> 1 Then
'****** END CODE UPDATE 04/28/2018 **************
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    
'    If mySrc = "Transfer Voucher" Then
'        Call CancelTransferVoucher
'    ElseIf mySrc = "Sales Invoice" Then
'        SalesInvCancel.MySource = "CANCEL"
'        SalesInvCancel.txtInvID.Text = flxSum.TextMatrix(ActRow, 9)
'        SalesInvCancel.Tag = ActRow
'        Call SalesInvCancel.ShowRecord
'        SalesInvCancel.Show vbModal
'    ElseIf mySrc = "Delivery Receipt" Then
'        Call CancelDeliveryReceipt
'    ElseIf mySrc = "Daily Cash Collection" Then
'        Call CancelOR
'    ElseIf mySrc = "Disbursement" Then
'        Call CancelDisbursement
'    ElseIf mySrc = "Sales Return" Then
'        Call CancelSalesReturn
'    ElseIf mySrc = "Inventory DR" Then
'        Call CancelInventoryDR
'    ElseIf mySrc = "Service Request" Then
'        Call CancelServiceRequest
'    ElseIf mySrc = "Spareparts Invoice" Then
'        SPInvCancel.MyInv = flxSum.TextMatrix(ActRow, 2)
'        SPInvCancel.CancelInv = True
'        SPInvCancel.iRow = ActRow
'        SPInvCancel.Show vbModal
'    ElseIf mySrc = "Job Order" Then
'        Call CancelJO
'    ElseIf mySrc = "Advances for Liquidation" Then
'        Call CancelAFL
'    ElseIf mySrc = "DMCM" Then
'        'Call CancelDMCM
'    ElseIf mySrc = "DAILY P.R." Then
'        Call CancelDailyPR
'    ElseIf mySrc = "Inventory TRA" Then
'        Call CancelTrustReceipt
'    ElseIf mySrc = "Deed of Redemption" Then
'        Call CancelRedemption
'    ElseIf mySrc = "Certificate of Ownership" Then
'        Call CancelOwnership
'    Else
        MsgBox "NOT AVAILABLE", vbInformation + vbOKOnly, "VALIDATION"
'    End If
End Sub

Private Sub CancelDailyPR()
    Dim myReason As String, MyAcc As String, Qry As String
    
    If InStr(UCase$(flxSum.TextMatrix(ActRow, 6)), "CANCEL") > 0 Then
        MsgBox "P.R. no. " & flxSum.TextMatrix(ActRow, 2) & " was already cancelled!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX = 47;"
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    ans = MsgBox("Do you want to Cancel Daily P.R.?" & vbCrLf & vbCrLf & _
        "Reference Date : " & flxSum.TextMatrix(ActRow, 1) & vbCrLf & "Reference No.: " & flxSum.TextMatrix(ActRow, 2), vbQuestion + vbYesNo, "Cancel")
    If ans = vbNo Then Exit Sub
    myReason = InputBox("Please indicate reason for Cancellation of P.R.", "Reason for Cancel")
    If myReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Sales Return failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    LYRcon.Execute "UPDATE dailypr SET REMARKS = 'CANCELLED' Where dailypr.PRID = " & Val(flxSum.TextMatrix(ActRow, 7)) & ""
    
    LYRcon.Execute "INSERT INTO canceltrans(TRANSNAME, TRANSDATE, REFNO, REASON, CANCELLEDBY, CANCELDATE) " & _
                   "VALUES('DAILY P.R.', '" & Format(Date, "YYYY-MM-DD") & "', '" & flxSum.TextMatrix(ActRow, 2) & "', '" & myReason & "'," & UserID & ", '" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "')"
    
    With flxSum     'Change row color
        .TextMatrix(ActRow, 6) = "CANCELLED"
        .col = 1: .row = ActRow: .ColSel = TCol: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 0, 0)
        .col = 2
    End With
    MsgBox "DAILY P.R. Cancelled!", vbInformation + vbOKOnly
    
    
End Sub

Private Sub CancelServiceRequest()
    Dim myReason As String, MyAcc As String, Qry As String, MyType As String
    Dim ans As Integer, i As Integer, GCode As Integer
    Dim rsPRdtls As New ADODB.Recordset
    
    If UCase$(flxSum.TextMatrix(ActRow, 6)) = "CANCELLED" Then
        MsgBox flxSum.TextMatrix(ActRow, 2) & " was already cancelled!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    If Val(flxSum.TextMatrix(ActRow, 8)) = 0 Then
        MsgBox flxSum.TextMatrix(ActRow, 2) & " is not yet final!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=33;"
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    ans = MsgBox("Do you want to Cancel Service Request?" & vbCrLf & vbCrLf & _
        "Reference Date : " & flxSum.TextMatrix(ActRow, 1) & vbCrLf & "Reference No.: " & flxSum.TextMatrix(ActRow, 2), vbQuestion + vbYesNo, "Cancel")
    If ans = vbNo Then Exit Sub
    myReason = InputBox("Please indicate reason for Cancellation of Service Request", "Reason for Cancel")
    If myReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Sales Return failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    
    LYRcon.Execute "UPDATE invsvcreq SET REMARKS = 'CANCELLED' Where invsvcreq.SVCREQID = " & Val(flxSum.TextMatrix(ActRow, 9)) & ";"
    
    flxCancel.Rows = 1
    Qry = "SELECT invsvcdtls.ITEMID, invsvcdtls.`STATUS` From invsvcdtls Where invsvcdtls.SVCREQID = " & Val(flxSum.TextMatrix(ActRow, 9)) & ";"
    With rsPRdtls
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxCancel.AddItem !ItemID & vbTab & IIf(IsNull(!Status), "", !Status)
                .MoveNext
            Loop
        End If
        .Close
    End With
    'update inventory register
    LYRcon.Execute "UPDATE invtreg SET TRANSTYPE = 'CANCELLED', UPDATEDBY = " & UserID & ", LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
        "Where invtreg.TRANSID = " & Val(flxSum.TextMatrix(ActRow, 9)) & " AND invtreg.SRCTYPE LIKE '%SERVICE REQUEST';"
    For i = 1 To flxCancel.Rows - 1
        'Update item status
        GCode = GetGroupCode(flxCancel.TextMatrix(i, 1))
        LYRcon.Execute "UPDATE items SET `STATUS` = '" & flxCancel.TextMatrix(i, 1) & "', GRPCODE = " & GCode & " Where items.ItemID = " & Val(flxCancel.TextMatrix(i, 0)) & ";"
    Next i
    LYRcon.Execute "INSERT INTO canceltrans(TRANSNAME, TRANSDATE, REFNO, REASON, CANCELLEDBY, CANCELDATE) " & _
                   "VALUES('Service Request', '" & Format(Date, "YYYY-MM-DD") & "', '" & flxSum.TextMatrix(ActRow, 2) & "', '" & myReason & "'," & UserID & ", '" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "')"
    
    With flxSum     'Change row color
        .TextMatrix(ActRow, 6) = "CANCELLED"
        .col = 1: .row = ActRow: .ColSel = TCol: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 0, 0)
        .col = 2
    End With
    MsgBox "Service Request Cancelled!", vbInformation + vbOKOnly
    
End Sub


Private Sub CancelPurchaseReturn()
    Dim myReason As String, MyAcc As String, Qry As String, MyType As String
    Dim ans As Integer, i As Integer, GCode As Integer
    Dim rsPRdtls As New ADODB.Recordset
    
    If UCase$(flxSum.TextMatrix(ActRow, 6)) = "CANCELLED" Then
        MsgBox flxSum.TextMatrix(ActRow, 2) & " was already cancelled!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    If Val(flxSum.TextMatrix(ActRow, 8)) = 0 Then
        MsgBox flxSum.TextMatrix(ActRow, 2) & " is not yet final!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=32;"
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    ans = MsgBox("Do you want to Cancel Purchase Return?" & vbCrLf & vbCrLf & _
        "Reference Date : " & flxSum.TextMatrix(ActRow, 1) & vbCrLf & "Reference No.: " & flxSum.TextMatrix(ActRow, 2), vbQuestion + vbYesNo, "Cancel")
    If ans = vbNo Then Exit Sub
    myReason = InputBox("Please indicate reason for Cancellation of Purchase Return", "Reason for Cancel")
    If myReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Sales Return failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    
    LYRcon.Execute "UPDATE invpureturn SET REASON = 'CANCELLED' Where invpureturn.PURCHRETID = " & Val(flxSum.TextMatrix(ActRow, 9)) & ";"
    
    flxCancel.Rows = 1
    Qry = "SELECT invtreg.ItemID, invtreg.REMARKS From invtreg Where invtreg.TRANSID = " & Val(flxSum.TextMatrix(ActRow, 9)) & " AND invtreg.REFERENCE = '" & flxSum.TextMatrix(ActRow, 2) & "';"
    With rsPRdtls
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxCancel.AddItem !ItemID & vbTab & IIf(IsNull(!Remarks), "", !Remarks)
                .MoveNext
            Loop
        End If
        .Close
    End With
    'update inventory register
    LYRcon.Execute "UPDATE invtreg SET TRANSTYPE = 'CANCELLED', UPDATEDBY = " & UserID & ", LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
        "Where invtreg.TRANSID = " & Val(flxSum.TextMatrix(ActRow, 9)) & " AND invtreg.SRCTYPE Like '%RETURN UNIT';"
    For i = 1 To flxCancel.Rows - 1
        'Update item status
        GCode = GetGroupCode(flxCancel.TextMatrix(i, 1))
        LYRcon.Execute "UPDATE items SET `STATUS` = '" & flxCancel.TextMatrix(i, 1) & "', GRPCODE = " & GCode & " Where items.ItemID = " & Val(flxCancel.TextMatrix(i, 0)) & ";"
    Next i
    LYRcon.Execute "INSERT INTO canceltrans(TRANSNAME, TRANSDATE, REFNO, REASON, CANCELLEDBY, CANCELDATE) " & _
                   "VALUES('PURCHASE RETURN', '" & Format(Date, "YYYY-MM-DD") & "', '" & flxSum.TextMatrix(ActRow, 2) & "', '" & myReason & "'," & UserID & ", '" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "')"
    
    If Val(cmdNext.Tag) = 2 Then
        Call RefreshRecord(2)
    Else
        Call RefreshRecord(1)
    End If
    
    
    MsgBox "Purchase Return Cancelled!", vbInformation + vbOKOnly

End Sub

Private Sub CancelInventoryDR()
    Dim myReason As String, MyAcc As String, Qry As String, MyType As String
    Dim ans As Integer, i As Integer, GCode As Integer
    Dim rsDRdtls As New ADODB.Recordset
    
    If UCase$(flxSum.TextMatrix(ActRow, 6)) = "CANCELLED" Then
        MsgBox "D.R. No. " & flxSum.TextMatrix(ActRow, 2) & " was already cancelled!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    If Val(flxSum.TextMatrix(ActRow, 7)) = 0 Then
        MsgBox "D.R. No. " & flxSum.TextMatrix(ActRow, 2) & " is not yet final!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=31;"
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    ans = MsgBox("Do you want to Cancel Delivery Receipt?" & vbCrLf & vbCrLf & _
        "D.R. Date : " & flxSum.TextMatrix(ActRow, 1) & vbCrLf & "D.R. No: " & flxSum.TextMatrix(ActRow, 2), vbQuestion + vbYesNo, "Cancel")
    If ans = vbNo Then Exit Sub
    myReason = InputBox("Please indicate reason for Cancellation of D.R.", "Reason for Cancel")
    If myReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Sales Return failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    
    LYRcon.Execute "UPDATE SalesDR SET REMARKS = 'CANCELLED' Where SalesDR.SALESDRID = " & Val(flxSum.TextMatrix(ActRow, 8)) & ";"
    
    flxCancel.Rows = 1
    Qry = "SELECT invdrdtls.ITEMID, invdrdtls.REMARKS From invdrdtls Where invdrdtls.SALESDRID = " & Val(flxSum.TextMatrix(ActRow, 8)) & ""
    With rsDRdtls
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxCancel.AddItem !ItemID & vbTab & IIf(IsNull(!Remarks), "", !Remarks)
                .MoveNext
            Loop
        End If
        .Close
    End With
    Qry = "SELECT SalesDR.DRTYPE From SalesDR Where SalesDR.SALESDRID = " & Val(flxSum.TextMatrix(ActRow, 8)) & ";"
    MyType = GetFieldValue(Qry)
    For i = 1 To flxCancel.Rows - 1
        'Update item status
        GCode = GetGroupCode(flxCancel.TextMatrix(i, 1))
        If InStr(MyType, "SERVICE UNIT") > 0 Then
            LYRcon.Execute "UPDATE items SET `STATUS` = '" & flxCancel.TextMatrix(i, 1) & "', GRPCODE = " & GCode & ", REMARKS = '' Where items.ItemID = " & Val(flxCancel.TextMatrix(i, 0)) & ";"
        Else
            LYRcon.Execute "UPDATE items SET `STATUS` = '" & flxCancel.TextMatrix(i, 1) & "', GRPCODE = " & GCode & " Where items.ItemID = " & Val(flxCancel.TextMatrix(i, 0)) & ";"
        End If
        'update inventory register
        LYRcon.Execute "UPDATE invtreg SET TRANSTYPE = 'CANCELLED', UPDATEDBY = " & UserID & ", LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
            "Where invtreg.TRANSID = " & Val(flxSum.TextMatrix(ActRow, 8)) & " AND invtreg.ITEMID = " & Val(flxCancel.TextMatrix(i, 0)) & " AND invtreg.REFERENCE = 'DR#:" & flxSum.TextMatrix(ActRow, 2) & "';"
    Next i
    LYRcon.Execute "INSERT INTO canceltrans(TRANSNAME, TRANSDATE, REFNO, REASON, CANCELLEDBY, CANCELDATE) " & _
                   "VALUES('D.R. INVENTORY', '" & Format(Date, "YYYY-MM-DD") & "', '" & flxSum.TextMatrix(ActRow, 2) & "', '" & myReason & "'," & UserID & ", '" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "')"
    
    With flxSum     'Change row color
        .TextMatrix(ActRow, 6) = "CANCELLED"
        .col = 1: .row = ActRow: .ColSel = TCol: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 0, 0)
        .col = 2
    End With
    MsgBox "Sales Return Cancelled!", vbInformation + vbOKOnly

End Sub


Private Sub CancelAFL()
    Dim myReason As String, MyAcc As String, Qry As String
    Dim ans As Integer, i As Integer
    Dim rsSRdtls As New ADODB.Recordset
    
    If UCase$(flxSum.TextMatrix(ActRow, 7)) = "CANCELLED" Then
        MsgBox "AFL No. " & flxSum.TextMatrix(ActRow, 2) & " was already cancelled!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    If Val(flxSum.TextMatrix(ActRow, 8)) = 0 Then
        MsgBox "AFL No. " & flxSum.TextMatrix(ActRow, 2) & " is not yet final!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=65;"
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    ans = MsgBox("Do you want to Cancel Advances for Liquidation?" & vbCrLf & vbCrLf & _
        "AFL Date : " & flxSum.TextMatrix(ActRow, 1) & vbCrLf & "AFL No: " & flxSum.TextMatrix(ActRow, 2), vbQuestion + vbYesNo, "Cancel")
    If ans = vbNo Then Exit Sub
    myReason = InputBox("Please indicate reason for Cancellation of AFL:", "Reason for Cancel")
    If myReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of AFL failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    
    LYRcon.Execute "UPDATE advliquidation SET REMARKS = 'CANCELLED' Where advliquidation.AFLID = " & Val(flxSum.TextMatrix(ActRow, 9)) & ";"
    LYRcon.Execute "INSERT INTO canceltrans(TRANSNAME, TRANSDATE, REFNO, REASON, CANCELLEDBY, CANCELDATE) " & _
                   "VALUES('AFL', '" & Format(Date, "YYYY-MM-DD") & "', '" & flxSum.TextMatrix(ActRow, 2) & "', '" & myReason & "'," & UserID & ", '" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "')"
    With flxSum     'Change row color
        .TextMatrix(ActRow, 7) = "CANCELLED"
        .col = 1: .row = ActRow: .ColSel = TCol: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 0, 0)
        .col = 2
    End With
    MsgBox "Advances for Liquidaton Cancelled!", vbInformation + vbOKOnly

End Sub

Private Sub CancelSalesReturn()
    Dim myReason As String, MyAcc As String, Qry As String
    Dim ans As Integer, i As Integer
    Dim rsSRdtls As New ADODB.Recordset
    
    If UCase$(flxSum.TextMatrix(ActRow, 5)) = "CANCELLED" Then
        MsgBox "Sales Return No. " & flxSum.TextMatrix(ActRow, 2) & " was already cancelled!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    If Val(flxSum.TextMatrix(ActRow, 6)) = 0 Then
        MsgBox "Sales Return No. " & flxSum.TextMatrix(ActRow, 2) & " is not yet final!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=25;"
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    ans = MsgBox("Do you want to Cancel Sales Return?" & vbCrLf & vbCrLf & _
        "Sales Return Date : " & flxSum.TextMatrix(ActRow, 1) & vbCrLf & "Sales Return No: " & flxSum.TextMatrix(ActRow, 2), vbQuestion + vbYesNo, "Cancel")
    If ans = vbNo Then Exit Sub
    myReason = InputBox("Please indicate reason for Cancellation of Sales Return:", "Reason for Cancel")
    If myReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Sales Return failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    
    LYRcon.Execute "UPDATE invsalret SET REASON = 'CANCELLED' Where invsalret.SRETURNID = " & Val(flxSum.TextMatrix(ActRow, 7)) & ";"
    
    flxCancel.Rows = 1
    Qry = "SELECT * From invsretdtls Where invsretdtls.SRETURNID = " & Val(flxSum.TextMatrix(ActRow, 7)) & ";"
    With rsSRdtls
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxCancel.AddItem !ItemID
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    For i = 1 To flxCancel.Rows - 1
        'Update item status
        LYRcon.Execute "UPDATE items SET `STATUS` = 'SOLD', GRPCODE = 9 Where items.ItemID = " & Val(flxCancel.TextMatrix(i, 0)) & ";"
        'update inventory register
        LYRcon.Execute "UPDATE invtreg SET TRANSTYPE = 'CANCELLED', UPDATEDBY = " & UserID & ", LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
            "Where invtreg.TRANSID = " & Val(flxSum.TextMatrix(ActRow, 7)) & " AND invtreg.ITEMID = " & Val(flxCancel.TextMatrix(i, 0)) & " AND invtreg.SRCTYPE LIKE '%SALES RETURN'"
    Next i
    LYRcon.Execute "INSERT INTO canceltrans(TRANSNAME, TRANSDATE, REFNO, REASON, CANCELLEDBY, CANCELDATE) " & _
                   "VALUES('SALES RETURN', '" & Format(Date, "YYYY-MM-DD") & "', '" & flxSum.TextMatrix(ActRow, 2) & "', '" & myReason & "'," & UserID & ", '" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "')"
    
    With flxSum     'Change row color
        .TextMatrix(ActRow, 5) = "CANCELLED"
        .col = 1: .row = ActRow: .ColSel = TCol: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 0, 0)
        .col = 2
    End With
    MsgBox "Sales Return Cancelled!", vbInformation + vbOKOnly
    
    
End Sub

Private Sub CancelDisbursement()
    Dim ans As Integer, myReason As String, MyAcc As String, Qry As String, DocRef As String, DocRef2 As String
    Dim OrderID As Long, MyBal As Currency, MYTax As Currency
    
    If InStr(UCase$(flxSum.TextMatrix(ActRow, 3)), "CANCELLED") > 0 Then
        MsgBox "Voucher No. " & flxSum.TextMatrix(ActRow, 2) & " was already cancelled!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    
    ans = MsgBox("Do you want to Cancel Voucher?" & vbCrLf & vbCrLf & _
        "PCV Date : " & flxSum.TextMatrix(ActRow, 1) & vbCrLf & "Voucher No: " & flxSum.TextMatrix(ActRow, 2) & vbCrLf & _
        "Received By: " & flxSum.TextMatrix(ActRow, 5) & vbCrLf & _
        "Account Desc.: " & flxSum.TextMatrix(ActRow, 4), vbQuestion + vbYesNo, "Cancel")
    If ans = vbNo Then Exit Sub
    myReason = InputBox("Please indicate reason for Cancellation of Voucher:", "Reason for Cancel")
    If myReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Voucher failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    
    'Update Disbursement
    LYRcon.Execute "UPDATE DISBURSEMENT SET DISBURSEMENT.DISBURSETYPE = 'Cancelled Vouchers', DISBURSEMENT.REMARKS = '" & Left$(myReason, 30) & "', DISBURSEMENT.AMOUNT = 0, DISBURSEMENT.DOCREF= '', DISBURSEMENT.DOCREF2 = '' " & _
                   "WHERE DISBURSEMENT.CVNUMBER='" & flxSum.TextMatrix(ActRow, 2) & "';"
    
    'If Disburse type = Refund ... revert all updates
    If InStr(flxSum.TextMatrix(ActRow, 4), "REFUND") > 0 And InStr(flxSum.TextMatrix(ActRow, 3), "Advances From") > 0 And Val(flxSum.TextMatrix(ActRow, 11)) > 0 Then
        If flxSum.TextMatrix(ActRow, 13) <> "" Then
            DocRef = flxSum.TextMatrix(ActRow, 13)
            Qry = "SELECT payments.OrderID From payments Where payments.ORNO = '" & DocRef & "';"
            OrderID = Val(GetFieldValue(Qry))
            If OrderID > 0 Then 'Get Account No if Exist to update table Payment
                Qry = "SELECT unitordered.ACCOUNTNO From unitordered Where unitordered.OrderID = " & OrderID & ";"
                MyAcc = GetFieldValue(Qry)
                LYRcon.Execute "UPDATE payments SET ACCOUNTNO = '" & MyAcc & "', REMARKS = '' Where payments.ORNO = '" & DocRef & "';"
            End If
        End If
        If flxSum.TextMatrix(ActRow, 14) <> "" Then
            DocRef2 = flxSum.TextMatrix(ActRow, 14)
            Qry = "SELECT payments.OrderID From payments Where payments.ORNO = '" & DocRef2 & "';"
            OrderID = Val(GetFieldValue(Qry))
            If OrderID > 0 Then 'Get Account No if Exist to update table Payment
                Qry = "SELECT unitordered.ACCOUNTNO From unitordered Where unitordered.OrderID = " & OrderID & ";"
                MyAcc = GetFieldValue(Qry)
                LYRcon.Execute "UPDATE payments SET ACCOUNTNO = '" & MyAcc & "', REMARKS = '' Where payments.ORNO = '" & DocRef2 & "';"
            End If
        End If
    End If
    'If Disburse TYpe = Commission
    If InStr(flxSum.TextMatrix(ActRow, 4), "COMMISSION") > 0 And InStr(flxSum.TextMatrix(ActRow, 3), "Operating Expenses") > 0 And Val(flxSum.TextMatrix(ActRow, 12)) > 0 Then
        If flxSum.TextMatrix(ActRow, 13) <> "" Then
            DocRef = flxSum.TextMatrix(ActRow, 13)
            Qry = "SELECT agentorders.Balance From agentorders " & _
                  "Where agentorders.ORDERID = " & DocRef & " AND agentorders.AgentId = " & Val(flxSum.TextMatrix(ActRow, 12)) & ""
            MyBal = Val(GetFieldValue(Qry)) '+ Val(Replace(flxSum.TextMatrix(ActRow, 7), ",", ""))
            Qry = "SELECT disbursement.DocRef2 From disbursement Where disbursement.CVNUMBER = '" & flxSum.TextMatrix(ActRow, 2) & "'"
            DocRef2 = GetFieldValue(Qry)
            If Val(DocRef2) > 1 Then
                MyBal = MyBal + Val(DocRef2)
            Else
                Qry = "SELECT disbursement.TAXAMT From disbursement Where disbursement.CVNUMBER = '" & flxSum.TextMatrix(ActRow, 2) & "'"
                MYTax = Val(GetFieldValue(Qry))
                If MYTax > 1 Then MyBal = MyBal + (MYTax / 0.1)
            End If
            Qry = "SELECT agentorders.CLAIMREF From agentorders " & _
                  "Where agentorders.ORDERID = " & DocRef & " AND agentorders.AgentId = " & Val(flxSum.TextMatrix(ActRow, 12)) & ""
            MyAcc = GetFieldValue(Qry)
            If InStr(MyAcc, ",") > 0 Then
                MyAcc = Replace(MyAcc, flxSum.TextMatrix(ActRow, 2), "")
                MyAcc = Replace(MyAcc, ",", "")
            Else
                MyAcc = Replace(MyAcc, "PCV#:" & flxSum.TextMatrix(ActRow, 2), "")
            End If
            LYRcon.Execute "UPDATE agentorders SET BALANCE = " & MyBal & ", CLAIMREF = '" & MyAcc & "' " & _
                "Where agentorders.ORDERID = " & DocRef & " AND agentorders.AgentId = " & Val(flxSum.TextMatrix(ActRow, 12)) & ";"
        End If
    End If
        
    'Update CANCELLED Trans
    LYRcon.Execute "INSERT INTO canceltrans(TRANSNAME, TRANSDATE, REFNO, REASON, CANCELLEDBY, CANCELDATE) " & _
                   "VALUES('CASH VOUCHER', '" & Format(Date, "YYYY-MM-DD") & "', '" & flxSum.TextMatrix(ActRow, 2) & "', '" & myReason & "'," & UserID & ", '" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "')"
    'Update Grid
    flxSum.TextMatrix(ActRow, 3) = "Cancelled Vouchers"
    flxSum.TextMatrix(ActRow, 8) = myReason
    flxSum.TextMatrix(ActRow, 7) = "0.00"
    With flxSum     'Change row color
        .col = 1: .row = ActRow: .ColSel = 13: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 0, 0)
        .col = 2
    End With
    
    MsgBox "Voucher Cancelled!", vbInformation + vbOKOnly
End Sub

Private Sub CancelOR()
    Dim ans As Integer, i As Integer, myReason As String, Qry As String, SPor As String, CutOff As String
    Dim MyOrd As Long, SPinvID As Long
    
    If InStr(UCase$(flxSum.TextMatrix(ActRow, 4)), "CANCELLED") > 0 Then
        If flxSum.TextMatrix(ActRow, 2) <> "" Then
            MsgBox "O.R. No. " & flxSum.TextMatrix(ActRow, 2) & " was already cancelled!", vbExclamation + vbOKOnly, "Cancel"
            Exit Sub
        Else
            MsgBox "C.R. No. " & flxSum.TextMatrix(ActRow, 3) & " was already cancelled!", vbExclamation + vbOKOnly, "Cancel"
            Exit Sub
        End If
        
    End If
    If flxSum.TextMatrix(ActRow, 2) <> "" Then
        ans = MsgBox("Do you want to Cancel O.R. No. " & flxSum.TextMatrix(ActRow, 2) & "?", vbQuestion + vbYesNo, "Cancel")
    Else
        ans = MsgBox("Do you want to Cancel C.R. No. " & flxSum.TextMatrix(ActRow, 3) & "?", vbQuestion + vbYesNo, "Cancel")
    End If
    If ans = vbNo Then Exit Sub
    myReason = InputBox("Please indicate reason for Cancellation of O.R.:", "Reason for Cancel")
    If myReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of O.R. failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    
    If Val(flxSum.TextMatrix(ActRow, 9)) > 0 Then  'Spare Parts'
        Qry = "SELECT SPInvoice.SPINVOICEID From SPInvoice Where spinvoice.ORNO LIKE '%" & flxSum.TextMatrix(ActRow, 2) & "%';"
        SPinvID = Val(GetFieldValue(Qry))
        If SPinvID > 0 Then
            Qry = "SELECT spinvoice.ORNO From SPInvoice Where SPInvoice.SPINVOICEID = " & SPinvID & ";"
            SPor = GetFieldValue(Qry)
            SPor = Replace(SPor, flxSum.TextMatrix(ActRow, 2), "")
            LYRcon.Execute "UPDATE SPInvoice SET ORNO = '" & SPor & "', Balance = SPInvoice.Balance + " & Val(flxSum.TextMatrix(ActRow, 8)) & " Where SPInvoice.SPINVOICEID = " & SPinvID & ";"
            'LYRcon.Execute "UPDATE SPInvoice SET ORNO = '" & SPor & "' Where SPInvoice.SPINVOICEID = " & SPinvID & ";"
        End If
    ElseIf Val(flxSum.TextMatrix(ActRow, 10)) > 0 Then  'Service Center
        Qry = "SELECT SvcJobOrder.JOBORDERID From SvcJobOrder Where svcjoborder.ORNO = '" & flxSum.TextMatrix(ActRow, 2) & "';"
        SPinvID = Val(GetFieldValue(Qry))
        If SPinvID > 0 Then
            LYRcon.Execute "UPDATE SvcJobOrder SET PAYMENT = svcjoborder.PAYMENT - " & Val(flxSum.TextMatrix(ActRow, 9)) & ", BALANCE = svcjoborder.BALANCE + " & Val(flxSum.TextMatrix(ActRow, 9)) & ", ORNO = '' Where SvcJobOrder.JOBORDERID = " & SPinvID & ";"
        End If
    End If
    
    MyOrd = Val(flxSum.TextMatrix(ActRow, 15))
    LYRcon.Execute "UPDATE payments SET ACCOUNTNO = 'CANCELLED', CASH = 0, DOWNPAYMENT = 0, INSTALLMENT = 0, MOTORSHOP = 0, SRVCENTER = 0, DISCOUNT = 0, OTHERS = 0, ORDERID = 0 " & _
                   "Where payments.PAYMENTID = " & Val(flxSum.TextMatrix(ActRow, 15)) & ";"
    
    
    If Val(Replace(flxSum.TextMatrix(ActRow, 6), ",", "")) > 0 Or Val(Replace(flxSum.TextMatrix(ActRow, 7), ",", "")) > 0 Or Val(Replace(flxSum.TextMatrix(ActRow, 8), ",", "")) > 0 Then
    
        With tabWait
            .Top = (Me.Height - .Height) \ 2
            .Left = (Me.Width - .Width) \ 2
            .Caption = ""
            .Visible = True
            .ZOrder
        End With
        
        cmdShow.Enabled = False
        cmdPrint.Enabled = False
        cmdFilter.Enabled = False
        cmdExit.Enabled = False
        cmdSalesDrive.Enabled = False
        cmdTransmittal.Enabled = False
        
        lblProcess.Caption = "Processing Ledger..."
        DoEvents
        
        'Qry = "SELECT [AGEMONTH] & '/6/' & [AGEYEAR] AS CUTOFF FROM AGINGCURRENT;"
        Qry = "SELECT CONCAT(AGEMONTH, '/6/',AGEYEAR) AS CUTOFF From agingcurrent;"
        CutOff = GetFieldValue(Qry)
        If IsDate(CutOff) Then
            txtMo.Text = Month(DateValue(CutOff))
            txtYr.Text = Year(DateValue(CutOff))
            txtCutOff.Text = DateAdd("m", 1, DateValue(CutOff))
                
            Qry = "SELECT * From agingcurrent Where agingcurrent.OrderID = " & MyOrd & ""
            
            If CtrlNumExist(Qry) Then
                'Print O.R. Here
                Qry = "SELECT UNITORDERED.ACCOUNTNO, UNITORDERED.ORDERID, UNITORDERED.MONTHLYINST, UNITORDERED.FIRSTDUEDATE, UNITORDERED.TERM, UNITORDERED.TERMTYPE, UNITORDERED.DATEPURCHASED, UNITORDERED.ACCTSTATUS, UNITORDERED.INVOICENO, UNITORDERED.INVAMOUNT, UNITORDERED.COLLECTORID, UNITORDERED.AREA, UNITORDERED.TOTALINSTPRICE, UNITORDERED.REGDP, UNITORDERED.MIREBATE, UNITORDERED.SRP, UNITORDERED.LCP, UNITORDERED.ITEMID, UNITORDERED.MODEL, UNITORDERED.SERIALNO " & _
                      "From UNITORDERED WHERE (UNITORDERED.ORDERID = " & MyOrd & ");"
                If rsAccts.State = adStateOpen Then rsAccts.Close
                rsAccts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                Set flxAccts.DataSource = rsAccts
                flxAccts.Refresh
        
                'Update Aging
                If flxAccts.TextMatrix(1, 6) = "D" Then
                    Call DispDailyLedger
                Else
                    Call DisplayLedger
                End If
                lblProcess.Caption = "Processing Aging..."
                DoEvents
                Call ReprocessAging
                
            End If
        End If
        Qry = "SELECT unitordered.DateClosed From unitordered Where unitordered.OrderID = " & MyOrd & ";"
        CutOff = GetFieldValue(Qry)
        If IsDate(CutOff) Then LYRcon.Execute "UPDATE unitordered SET DATECLOSED = NULL, CLOSINGREMARKS = '', ACCTSTATUS = 'OPEN' Where unitordered.OrderID = " & MyOrd & ";"
        
        tabWait.Top = 12000
        tabWait.Visible = False
        cmdShow.Enabled = True
        cmdPrint.Enabled = True
        cmdFilter.Enabled = True
        cmdExit.Enabled = True
        cmdSalesDrive.Enabled = True
        cmdTransmittal.Enabled = True
    End If
    
    flxSum.TextMatrix(ActRow, 4) = "CANCELLED"
    For i = 6 To 12
        flxSum.TextMatrix(ActRow, i) = ""
    Next i
    'Update CANCELLED Trans
    LYRcon.Execute "INSERT INTO canceltrans(TRANSNAME, TRANSDATE, REFNO, REASON, CANCELLEDBY, CANCELDATE) " & _
                   "VALUES('OFFICIAL RECEIPT', '" & Format(Date, "YYYY-MM-DD") & "', '" & IIf(flxSum.TextMatrix(ActRow, 2) <> "", flxSum.TextMatrix(ActRow, 2), flxSum.TextMatrix(ActRow, 3)) & " [" & MyOrd & "]', '" & myReason & "'," & UserID & ", '" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "')"
    
    With flxSum     'Change row color
        .col = 1: .row = ActRow: .ColSel = 14: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 0, 0)
        .col = 2
    End With
    If flxSum.TextMatrix(ActRow, 2) <> "" Then
        MsgBox "Official Reciept " & flxSum.TextMatrix(ActRow, 2) & " Cancelled!", vbInformation + vbOKOnly
    Else
        MsgBox "Collection Reciept " & flxSum.TextMatrix(ActRow, 3) & " Cancelled!", vbInformation + vbOKOnly
    End If
End Sub

Private Sub CancelDMCM()
    Dim ans As Integer, i As Integer, myReason As String, Qry As String, SPor As String, CutOff As String
    Dim MyOrd As Long, DepDt As String
    Dim rsPay As New ADODB.Recordset
    
    If InStr(UCase$(flxSum.TextMatrix(ActRow, 7)), "CANCELLED") > 0 Then
        MsgBox "DMCM No. " & flxSum.TextMatrix(ActRow, 2) & " was already cancelled!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    ans = MsgBox("Do you want to Cancel DMCM No. " & flxSum.TextMatrix(ActRow, 2) & "?", vbQuestion + vbYesNo, "Cancel")
    If ans = vbNo Then Exit Sub
    myReason = InputBox("Please indicate reason for Cancellation of DMCM:", "Reason for Cancel")
    If myReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of DMCM failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    
    '1-REFDATE, 2-REFNO, 3-REFTYPE, 4-ACCOUNTNO, 5-CUSTOMER, 6-REFAMT, 7-REASONS, 8-PRINTCOPY, 9-DMCMID
    LYRcon.Execute "UPDATE ledgeradj SET REASONS = 'CANCELLED', TRANSTYPE = 'TEMP' Where ledgeradj.DMCMID = " & Val(flxSum.TextMatrix(ActRow, 5)) & ";"
    Qry = "SELECT unitordered.OrderID From unitordered Where unitordered.ACCOUNTNO = '" & flxSum.TextMatrix(ActRow, 4) & "';"
    MyOrd = Val(GetFieldValue(Qry))
    If flxSum.TextMatrix(ActRow, 3) = "RECLAIM" Then
        'Display OR. Paid by customer during deposit
        Qry = "SELECT ledgeradj.REFDATE From ledgeradj Where ledgeradj.ORDERID = " & MyOrd & " AND ledgeradj.REFTYPE = 'REPOSSESS' Order By ledgeradj.REFDATE Desc"
        DepDt = GetFieldValue(Qry)
        If IsDate(DepDt) Then
            Qry = "SELECT * From payments Where payments.PAYDATE >= '" & Format(DateValue(DepDt), "yyyy-mm-dd") & "' AND payments.ORDERID = " & Val(txtOrderID.Text) & " AND payments.INSTALLMENT > 0 Order By payments.PAYDATE;"
            With rsPay
                flxCancel.Rows = 1
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        flxCancel.AddItem !ORNO & vbTab & !Installment & vbTab & !PAYMENTID
                        .MoveNext
                    Loop
                End If
                .Close
            End With
            For i = 1 To flxCancel.Rows - 1
                LYRcon.Execute "UPDATE payments SET INSTALLMENT = 0, ACCOUNTNO = '', OTHERS = " & Val(flxCancel.TextMatrix(i, 1)) & ", REMARKS = 'Deposited Unit' Where payments.PAYMENTID = " & Val(flxCancel.TextMatrix(i, 1)) & ";"
            Next i
        End If
    End If
    If MyOrd > 0 Then
        With tabWait
            .Top = (Me.Height - .Height) \ 2
            .Left = (Me.Width - .Width) \ 2
            .Caption = ""
            .Visible = True
            .ZOrder
        End With
        cmdShow.Enabled = False
        cmdPrint.Enabled = False
        cmdFilter.Enabled = False
        cmdExit.Enabled = False
        cmdSalesDrive.Enabled = False
        cmdTransmittal.Enabled = False
        lblProcess.Caption = "Processing Ledger..."
        DoEvents
        
        'Qry = "SELECT [AGEMONTH] & '/6/' & [AGEYEAR] AS CUTOFF FROM AGINGCURRENT;"
        Qry = "SELECT CONCAT(AGEMONTH, '/6/',AGEYEAR) AS CUTOFF From agingcurrent;"
        CutOff = GetFieldValue(Qry)
        If IsDate(CutOff) Then
            txtMo.Text = Month(DateValue(CutOff))
            txtYr.Text = Year(DateValue(CutOff))
            txtCutOff.Text = DateAdd("m", 1, DateValue(CutOff))
                
            Qry = "SELECT * From agingcurrent Where agingcurrent.OrderID = " & MyOrd & ""
            
            If CtrlNumExist(Qry) Then
                'Print DMCM Here
                Qry = "SELECT UNITORDERED.ACCOUNTNO, UNITORDERED.ORDERID, UNITORDERED.MONTHLYINST, UNITORDERED.FIRSTDUEDATE, UNITORDERED.TERM, UNITORDERED.TERMTYPE, UNITORDERED.DATEPURCHASED, UNITORDERED.ACCTSTATUS, UNITORDERED.INVOICENO, UNITORDERED.INVAMOUNT, UNITORDERED.COLLECTORID, UNITORDERED.AREA, UNITORDERED.TOTALINSTPRICE, UNITORDERED.REGDP, UNITORDERED.MIREBATE, UNITORDERED.SRP, UNITORDERED.LCP, UNITORDERED.ITEMID, UNITORDERED.MODEL, UNITORDERED.SERIALNO " & _
                      "From UNITORDERED WHERE (UNITORDERED.ORDERID = " & MyOrd & ");"
                If rsAccts.State = adStateOpen Then rsAccts.Close
                rsAccts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                Set flxAccts.DataSource = rsAccts
                flxAccts.Refresh
        
                'Update Aging
                If flxAccts.TextMatrix(1, 6) = "D" Then
                    Call DispDailyLedger
                Else
                    Call DisplayLedger
                End If
                lblProcess.Caption = "Processing Aging..."
                DoEvents
                Call ReprocessAging
                
            End If
        End If
        Qry = "SELECT unitordered.DateClosed From unitordered Where unitordered.OrderID = " & MyOrd & ";"
        CutOff = GetFieldValue(Qry)
        If IsDate(CutOff) Then LYRcon.Execute "UPDATE unitordered SET DATECLOSED = NULL, CLOSINGREMARKS = '', ACCTSTATUS = 'OPEN' Where unitordered.OrderID = " & MyOrd & ";"
        
        tabWait.Top = 12000
        tabWait.Visible = False
        cmdShow.Enabled = False
        cmdPrint.Enabled = False
        cmdFilter.Enabled = False
        cmdExit.Enabled = False
        cmdSalesDrive.Enabled = False
        cmdTransmittal.Enabled = False
    End If
    
    flxSum.TextMatrix(ActRow, 3) = "CANCELLED"
    For i = 5 To 11
        flxSum.TextMatrix(ActRow, i) = ""
    Next i
    'Update CANCELLED Trans
    LYRcon.Execute "INSERT INTO canceltrans(TRANSNAME, TRANSDATE, REFNO, REASON, CANCELLEDBY, CANCELDATE) " & _
                   "VALUES('OFFICIAL RECEIPT', '" & Format(Date, "YYYY-MM-DD") & "', '" & flxSum.TextMatrix(ActRow, 2) & " [" & MyOrd & "]', '" & myReason & "'," & UserID & ", '" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "')"
    
    With flxSum     'Change row color
        .col = 1: .row = ActRow: .ColSel = 14: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 0, 0)
        .col = 2
    End With
    
    MsgBox "Official Reciept " & flxSum.TextMatrix(ActRow, 2) & " Cancelled!", vbInformation + vbOKOnly
End Sub

Private Sub CancelDeliveryReceipt()
    Dim ans As Integer, i As Integer, myReason As String, Qry As String
    Dim MyOrd As Long
    
    If InStr(UCase$(flxSum.TextMatrix(ActRow, 6)), "CANCELLED") > 0 Then
        MsgBox "D.R. No. " & flxSum.TextMatrix(ActRow, 2) & " was already cancelled!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    If Val(flxSum.TextMatrix(ActRow, 7)) = 0 Then
        MsgBox "D.R. No. " & flxSum.TextMatrix(ActRow, 2) & " is not yet printed!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    
    ans = MsgBox("Do you want to Cancel D.R. No. " & flxSum.TextMatrix(ActRow, 2) & "?", vbQuestion + vbYesNo, "Cancel")
    If ans = vbNo Then Exit Sub
    myReason = InputBox("Please indicate reason for Cancellation of D.R.:", "Reason for Cancel")
    If myReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of D.R. failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    LYRcon.Execute "UPDATE SalesDR SET REMARKS = 'CANCELLED' Where SalesDR.SALESDRID = " & Val(flxSum.TextMatrix(ActRow, 8)) & ";"
    
    Qry = "SELECT SalesDR.ORDERID From SalesDR Where SalesDR.SALESDRID = " & Val(flxSum.TextMatrix(ActRow, 8)) & ";"
    MyOrd = Val(GetFieldValue(Qry))
    If MyOrd > 0 Then
        LYRcon.Execute "UPDATE unitordered SET DRNO = '' Where unitordered.ORDERID = " & MyOrd & ";"
    End If
    'Update Grid
    flxSum.TextMatrix(ActRow, 6) = "CANCELLED"
    'Update CANCELLED Trans
    LYRcon.Execute "INSERT INTO canceltrans(TRANSNAME, TRANSDATE, REFNO, REASON, CANCELLEDBY, CANCELDATE) " & _
                   "VALUES('DELIVERY RECEIPT', '" & Format(Date, "YYYY-MM-DD") & "', '" & flxSum.TextMatrix(ActRow, 2) & "', '" & myReason & "'," & UserID & ", '" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "')"
    
    With flxSum     'Change row color
        .col = 1: .row = ActRow: .ColSel = 8: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 0, 0)
        .col = 2
    End With
    MsgBox "Delivery Reciept " & flxSum.TextMatrix(ActRow, 2) & " Cancelled!", vbInformation + vbOKOnly


End Sub

Private Sub CancelOwnership()
    Dim ans As Integer, i As Integer, myReason As String, Qry As String
    Dim MyOrd As Long

    If InStr(UCase$(flxSum.TextMatrix(ActRow, 7)), "CANCELLED") > 0 Then
        MsgBox "Certificate NO. " & flxSum.TextMatrix(ActRow, 2) & " was already cancelled!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    
    ans = MsgBox("Do you want to Cancel Ownership with Certificate No. " & flxSum.TextMatrix(ActRow, 2) & "?", vbQuestion + vbYesNo, "Cancel")
    If ans = vbNo Then Exit Sub
    myReason = InputBox("Please indicate reason for Cancellation of Ownership:", "Reason for Cancel")
    If myReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Ownership failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    LYRcon.Execute "UPDATE ownership SET ORDERID = 0, REMARKS = 'CANCELLED' Where ownership.CERTID = " & Val(flxSum.TextMatrix(ActRow, 8)) & ";"
    
    'Update Grid
    flxSum.TextMatrix(ActRow, 7) = "CANCELLED"
    'Update CANCELLED Trans
    LYRcon.Execute "INSERT INTO canceltrans(TRANSNAME, TRANSDATE, REFNO, REASON, CANCELLEDBY, CANCELDATE) " & _
                   "VALUES('CERTIFICATE OF OWNERSHIP', '" & Format(Date, "YYYY-MM-DD") & "', '" & flxSum.TextMatrix(ActRow, 2) & "', '" & myReason & "'," & UserID & ", '" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "')"
    
    With flxSum     'Change row color
        .col = 1: .row = ActRow: .ColSel = 8: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 0, 0)
        .col = 2
    End With
    MsgBox "Certificate of Ownership no. " & flxSum.TextMatrix(ActRow, 2) & " Cancelled!", vbInformation + vbOKOnly
    
End Sub


Private Sub CancelRedemption()
    Dim rsITR As New ADODB.Recordset
    Dim ans As Integer, i As Integer, myReason As String, Qry As String
    Dim MyOrd As Long
    
    If InStr(UCase$(flxSum.TextMatrix(ActRow, 7)), "CANCELLED") > 0 Then
        MsgBox "REF. NO. " & flxSum.TextMatrix(ActRow, 2) & " was already cancelled!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    If Val(flxSum.TextMatrix(ActRow, 8)) = 0 Then
        MsgBox "REF. NO. " & flxSum.TextMatrix(ActRow, 2) & " is not yet printed!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    
    ans = MsgBox("Do you want to Cancel Redemption with REF. NO. " & flxSum.TextMatrix(ActRow, 2) & "?", vbQuestion + vbYesNo, "Cancel")
    If ans = vbNo Then Exit Sub
    myReason = InputBox("Please indicate reason for Cancellation of Redemption:", "Reason for Cancel")
    If myReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Redemption failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    LYRcon.Execute "UPDATE InvRedeem SET REMARKS = concat('CANCELLED', ' ', invredeem.REMARKS) Where InvRedeem.REDEEMID = " & Val(flxSum.TextMatrix(ActRow, 9)) & ";"
    'Get items transferred
    Qry = "SELECT invtreg.ITEMID, items.SERIALNO, invtreg.REMARKS, invtreg.OUT From invtreg INNER JOIN items ON (invtreg.ITEMID = items.ITEMID) " & _
          "Where invtreg.TRANSID = " & flxSum.TextMatrix(ActRow, 9) & " AND invtreg.REFERENCE = 'RDM#:" & flxSum.TextMatrix(ActRow, 2) & "';"
    With rsITR
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        flxCancel.Rows = 1
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxCancel.AddItem !ItemID & vbTab & IIf(IsNull(!SERIALNO), "", !SERIALNO) & vbTab & !Remarks & vbTab & !OUT
                .MoveNext
            Loop
        End If
        .Close
    End With
    'Update InvtReg -> TransType
    LYRcon.Execute "UPDATE INVTREG SET INVTREG.TRANSTYPE = 'CANCELLED' WHERE invtreg.REFERENCE = 'RDM#:" & flxSum.TextMatrix(ActRow, 2) & "' AND INVTREG.TRANSID=" & Val(flxSum.TextMatrix(ActRow, 9)) & ";"
    'Update Items Group code and Status
    With flxCancel
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 2) = "" Then .TextMatrix(i, 2) = "BRAND NEW"
            .TextMatrix(i, 4) = GetGroupCode(.TextMatrix(i, 2))
            LYRcon.Execute "UPDATE ITEMS SET ITEMS.STATUS = '" & .TextMatrix(i, 2) & "', ITEMS.GRPCODE = " & Val(.TextMatrix(i, 4)) & " WHERE (ITEMS.ITEMID=" & Val(.TextMatrix(i, 0)) & ");"
            If .TextMatrix(i, 1) = "" Then  'No Serial Applied - > Add Qty
                LYRcon.Execute "UPDATE items SET QTYONHAND = items.QTYONHAND + " & Val(.TextMatrix(i, 3)) & " Where items.ItemID = " & Val(.TextMatrix(i, 0)) & ";"
            End If
        Next i
    End With
    
    'Update Grid
    flxSum.TextMatrix(ActRow, 7) = "CANCELLED " & flxSum.TextMatrix(ActRow, 7)
    'Update CANCELLED Trans
    LYRcon.Execute "INSERT INTO canceltrans(TRANSNAME, TRANSDATE, REFNO, REASON, CANCELLEDBY, CANCELDATE) " & _
                   "VALUES('DEED OF REDEMPTION', '" & Format(Date, "YYYY-MM-DD") & "', '" & flxSum.TextMatrix(ActRow, 2) & "', '" & myReason & "'," & UserID & ", '" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "')"
    
    With flxSum     'Change row color
        .col = 1: .row = ActRow: .ColSel = 9: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 0, 0)
        .col = 2
    End With
    MsgBox "DEED OF REDEMPTION ref. no. " & flxSum.TextMatrix(ActRow, 2) & " Cancelled!", vbInformation + vbOKOnly


End Sub



Private Sub CancelTrustReceipt()
    Dim rsITR As New ADODB.Recordset
    Dim ans As Integer, i As Integer, myReason As String, Qry As String
    Dim MyOrd As Long
    
    If InStr(UCase$(flxSum.TextMatrix(ActRow, 7)), "CANCELLED") > 0 Then
        MsgBox "T.R.A. No. " & flxSum.TextMatrix(ActRow, 2) & " was already cancelled!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    If Val(flxSum.TextMatrix(ActRow, 8)) = 0 Then
        MsgBox "T.R.A. No. " & flxSum.TextMatrix(ActRow, 2) & " is not yet printed!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    
    ans = MsgBox("Do you want to Cancel T.R.A. No. " & flxSum.TextMatrix(ActRow, 2) & "?", vbQuestion + vbYesNo, "Cancel")
    If ans = vbNo Then Exit Sub
    myReason = InputBox("Please indicate reason for Cancellation of T.R.A.:", "Reason for Cancel")
    If myReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of D.R. failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    LYRcon.Execute "UPDATE InvTRA SET REMARKS = 'CANCELLED' Where invtra.TRAID = " & Val(flxSum.TextMatrix(ActRow, 9)) & ";"
    
    'Get items transferred
    Qry = "SELECT invtreg.ITEMID, items.SERIALNO, invtreg.REMARKS, invtreg.OUT From invtreg INNER JOIN items ON (invtreg.ITEMID = items.ITEMID) " & _
          "Where invtreg.TRANSID = " & flxSum.TextMatrix(ActRow, 9) & " AND invtreg.REFERENCE = 'TRA#" & flxSum.TextMatrix(ActRow, 2) & "';"
    With rsITR
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        flxCancel.Rows = 1
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxCancel.AddItem !ItemID & vbTab & IIf(IsNull(!SERIALNO), "", !SERIALNO) & vbTab & !Remarks & vbTab & !OUT
                .MoveNext
            Loop
        End If
        .Close
    End With
    'Update InvtReg -> TransType
    LYRcon.Execute "UPDATE INVTREG SET INVTREG.TRANSTYPE = 'CANCELLED' WHERE invtreg.REFERENCE = 'TRA#" & flxSum.TextMatrix(ActRow, 2) & "' AND INVTREG.TRANSID=" & Val(flxSum.TextMatrix(ActRow, 9)) & ";"
    'Update Items Group code and Status
    With flxCancel
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 2) = "" Then .TextMatrix(i, 2) = "BRAND NEW"
            .TextMatrix(i, 4) = GetGroupCode(.TextMatrix(i, 2))
            LYRcon.Execute "UPDATE ITEMS SET ITEMS.STATUS = '" & .TextMatrix(i, 2) & "', ITEMS.GRPCODE = " & Val(.TextMatrix(i, 4)) & " WHERE (ITEMS.ITEMID=" & Val(.TextMatrix(i, 0)) & ");"
            If .TextMatrix(i, 1) = "" Then  'No Serial Applied - > Add Qty
                LYRcon.Execute "UPDATE items SET QTYONHAND = items.QTYONHAND + " & Val(.TextMatrix(i, 3)) & " Where items.ItemID = " & Val(.TextMatrix(i, 0)) & ";"
            End If
        Next i
    End With
    
    'Update Grid
    flxSum.TextMatrix(ActRow, 7) = "CANCELLED"
    'Update CANCELLED Trans
    LYRcon.Execute "INSERT INTO canceltrans(TRANSNAME, TRANSDATE, REFNO, REASON, CANCELLEDBY, CANCELDATE) " & _
                   "VALUES('TRUST RECEIPT', '" & Format(Date, "YYYY-MM-DD") & "', '" & flxSum.TextMatrix(ActRow, 2) & "', '" & myReason & "'," & UserID & ", '" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "')"
    
    With flxSum     'Change row color
        .col = 1: .row = ActRow: .ColSel = 9: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 0, 0)
        .col = 2
    End With
    MsgBox "TRUST Reciept " & flxSum.TextMatrix(ActRow, 2) & " Cancelled!", vbInformation + vbOKOnly


End Sub


Private Sub CancelTransferVoucher()
    Dim rsITR As New ADODB.Recordset
    Dim ans As Integer, i As Integer, myReason As String, Qry As String
    
    If InStr(UCase$(flxSum.TextMatrix(ActRow, 5)), "CANCELLED") > 0 Then
        MsgBox "Transfer Voucher No. " & flxSum.TextMatrix(ActRow, 2) & " was already cancelled!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    ans = MsgBox("Do you want to Cancel Transfer Voucher?" & vbCrLf & vbCrLf & _
        "Transfer Voucher Date : " & flxSum.TextMatrix(ActRow, 1) & vbCrLf & "Transfer Voucher No: " & flxSum.TextMatrix(ActRow, 2), vbQuestion + vbYesNo, "Cancel")
    If ans = vbNo Then Exit Sub
    myReason = InputBox("Please indicate reason for Cancellation of Transfer Voucher:", "Reason for Cancel")
    If myReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Transfer Voucher failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    'Update InvTransfer->Remarks
    LYRcon.Execute "UPDATE INVTRANSFER SET INVTRANSFER.REMARKS = 'CANCELLED' WHERE INVTRANSFER.TRANSFERID=" & Val(flxSum.TextMatrix(ActRow, 10)) & ";"
    'Get items transferred
    Qry = "SELECT invtreg.ITEMID, items.SERIALNO, invtreg.REMARKS, invtreg.OUT From invtreg INNER JOIN items ON (invtreg.ITEMID = items.ITEMID) " & _
          "Where invtreg.TRANSID = " & flxSum.TextMatrix(ActRow, 10) & " AND invtreg.TRANSTYPE = 'TRANSFER';"
    With rsITR
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        flxCancel.Rows = 1
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxCancel.AddItem !ItemID & vbTab & IIf(IsNull(!SERIALNO), "", !SERIALNO) & vbTab & !Remarks & vbTab & !OUT
                .MoveNext
            Loop
        End If
        .Close
    End With
    'Update InvtReg -> TransType
    LYRcon.Execute "UPDATE INVTREG SET INVTREG.TRANSTYPE = 'CANCELLED' WHERE INVTREG.TRANSTYPE='TRANSFER' AND INVTREG.TRANSID=" & Val(flxSum.TextMatrix(ActRow, 10)) & ";"
    'Update Items Group code and Status
    With flxCancel
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 2) = "" Then .TextMatrix(i, 2) = "BRAND NEW"
            .TextMatrix(i, 4) = GetGroupCode(.TextMatrix(i, 2))
            LYRcon.Execute "UPDATE ITEMS SET ITEMS.STATUS = '" & .TextMatrix(i, 2) & "', ITEMS.GRPCODE = " & Val(.TextMatrix(i, 4)) & " WHERE (ITEMS.ITEMID=" & Val(.TextMatrix(i, 0)) & ");"
            If .TextMatrix(i, 1) = "" Then  'No Serial Applied - > Add Qty
                LYRcon.Execute "UPDATE items SET QTYONHAND = items.QTYONHAND + " & Val(.TextMatrix(i, 3)) & " Where items.ItemID = " & Val(.TextMatrix(i, 0)) & ";"
            End If
        Next i
    End With
    'Update Cancel Trans
    'LYRcon.Execute "INSERT INTO CANCELTRANS ( TRANSNAME, TRANSDATE, REFNO, REASON, CANCELLEDBY, CANCELDATE ) " & _
            "VALUES ('TRANSFER VOUCHER', '" & Format(DateValue(flxSum.TextMatrix(ActRow, 1)), "yyyy-mm-dd") & "', '" & flxSum.TextMatrix(ActRow, 2) & "' , '" & Left$(MyReason, 250) & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
   'Update Grid
    flxSum.TextMatrix(ActRow, 5) = "CANCELLED"
     'Update CANCELLED Trans
    LYRcon.Execute "INSERT INTO canceltrans(TRANSNAME, TRANSDATE, REFNO, REASON, CANCELLEDBY, CANCELDATE) " & _
                   "VALUES('TRANSFER VOUCHER', '" & Format(Date, "YYYY-MM-DD") & "', '" & flxSum.TextMatrix(ActRow, 2) & "', '" & myReason & "'," & UserID & ", '" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "')"
   
    With flxSum     'Change row color
        .col = 1: .row = ActRow: .ColSel = 9: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 0, 0)
        .col = 2
    End With
    MsgBox "Transfer Voucher Cancelled!", vbInformation + vbOKOnly
    
End Sub

Function GetGroupCode(MyStat As String) As Integer
    
    Select Case MyStat
        Case "BRAND NEW"
            GetGroupCode = 1
        Case "CLASS B"
            GetGroupCode = 2
        Case "REPO."
            GetGroupCode = 3
        Case "ACCESSORIES"
            GetGroupCode = 4
        Case "SERVICE UNIT"
            GetGroupCode = 5
        Case "PROMO"
            GetGroupCode = 6
        Case "DEPOSITED"
            GetGroupCode = 7
        Case "PENDING (SOLD)"
            GetGroupCode = 10
        Case "FOR REPAIR"
            GetGroupCode = 13
        Case Else
            GetGroupCode = 0
    End Select
End Function


Private Sub CancelJO()
    Dim ans As Integer, myReason As String
    
    If flxSum.TextMatrix(ActRow, 10) <> "" Then  'OR NO.<>''
        MsgBox "Job Order already Paid! Cancellation not allowed!", vbExclamation + vbOKOnly, "Validation"
        Exit Sub
    End If
    If InStr(UCase$(flxSum.TextMatrix(ActRow, 15)), "CANCELLED") > 0 Then
        MsgBox "Job Order No. " & flxSum.TextMatrix(ActRow, 2) & " was already cancelled!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    If Val(flxSum.TextMatrix(ActRow, 16)) = 0 Then  'Print = 0
        MsgBox "Job Order not yet printed! Cancellation not allowed!", vbExclamation + vbOKOnly, "Validation"
        Exit Sub
    End If
    
    ans = MsgBox("Do you want to Cancel Job Order?" & vbCrLf & vbCrLf & _
        "Job Order Date : " & flxSum.TextMatrix(ActRow, 1) & vbCrLf & "Job Order No: " & flxSum.TextMatrix(ActRow, 2) & vbCrLf & _
        "Customer: " & flxSum.TextMatrix(ActRow, 3) & vbCrLf & _
        "JO# Amount: " & flxSum.TextMatrix(ActRow, 6), vbQuestion + vbYesNo, "Cancel")
    If ans = vbNo Then Exit Sub
    myReason = InputBox("Please indicate reason for Cancellation of Job Order:", "Reason for Cancel")
    If myReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Job Order failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    'Update Job Order
    LYRcon.Execute "UPDATE SVCJOBORDER SET SVCJOBORDER.REMARKS = 'CANCELLED', SVCJOBORDER.TOTALAMT = 0, SVCJOBORDER.BALANCE = 0 WHERE SVCJOBORDER.JOBORDERID=" & flxSum.TextMatrix(ActRow, 17) & ";"
    'Update On account
    'If flxSum.TextMatrix(ActRow, 5) = "ON ACCOUNT" Then
    '    LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.STATUS = 'CANCELLED' WHERE (((ONACCTERMS.JOBORDERID)=" & flxSum.TextMatrix(ActRow, 14) & "));"
    'End If
    'Update Audit Trans
    LYRcon.Execute "INSERT INTO canceltrans(TRANSNAME, TRANSDATE, REFNO, REASON, CANCELLEDBY, CANCELDATE) " & _
                   "VALUES('JOB ORDER', '" & Format(Date, "YYYY-MM-DD") & "', '" & flxSum.TextMatrix(ActRow, 2) & "', '" & myReason & "'," & UserID & ", '" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "')"
   'Update Grid
    flxSum.TextMatrix(ActRow, 15) = "CANCELLED"
    flxSum.TextMatrix(ActRow, 6) = "0.00"
    flxSum.TextMatrix(ActRow, 8) = "0.00"
    With flxSum     'Change row color
        .col = 1: .row = ActRow: .ColSel = 16: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 0, 0)
        .col = 2
    End With
    MsgBox "Job Order Cancelled!", vbInformation + vbOKOnly
    
End Sub

Private Sub CancelSPlocAdj()
    Dim ans As Integer, myReason As String, Qry As String
    Dim rsPrv As New ADODB.Recordset
    
    ans = MsgBox("Do you want to Cancel Adjustment with Control No. " & flxJob.TextMatrix(ActRow, 2), vbQuestion + vbYesNo, "Cancel")
    If ans = vbNo Then Exit Sub
    myReason = InputBox("Please indicate reason for Cancellation of Adjustment:", "Reason for Cancel")
    If myReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Job Order failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    'Update SP Location Adjustment
    LYRcon.Execute "UPDATE TRANSFERLOC SET TRANSFERLOC.REFERENCE = 'CANCELLED' WHERE TRANSFERLOC.TRANSLOCID=" & flxJob.TextMatrix(ActRow, 8) & ";"
    Qry = "SELECT TRANSLOCDTLS.SPAREPARTID, TRANSLOCDTLS.PREVLOC From TRANSLOCDTLS WHERE TRANSLOCDTLS.TRANSLOCID=" & flxJob.TextMatrix(ActRow, 8) & ";"
    
    With rsPrv
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxCancel.Rows = 1
            .MoveFirst
            Do While Not .EOF
                flxCancel.AddItem !SPAREPARTID & vbTab & !PrevLoc
                .MoveNext
            Loop
        End If
        .Close
    End With
    With flxCancel
        For i = 1 To .Rows - 1
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.LOCATION = '" & .TextMatrix(i, 1) & "' WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 0) & ";"
        Next i
    End With
    'Update Audit Trans
    LYRcon.Execute "INSERT INTO canceltrans(TRANSNAME, TRANSDATE, REFNO, REASON, CANCELLEDBY, CANCELDATE) " & _
                   "VALUES('SP LOCATION ADJ.', '" & Format(Date, "YYYY-MM-DD") & "', '" & flxSum.TextMatrix(ActRow, 2) & "', '" & myReason & "'," & UserID & ", '" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "')"
   'Update Grid
    With flxSum     'Change row color
        .TextMatrix(ActRow, 3) = "CANCELLED"
        .col = 1: .row = ActRow: .ColSel = 8: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 0, 0)
        .col = 2
    End With
    MsgBox "Spare Parts Location Adjustment Cancelled!", vbInformation + vbOKOnly
End Sub




Private Sub mnuCancelPR_Click()
'    Call CancelPurchaseReturn
End Sub

Private Sub mnuHistory_Click()
    InvHistory.ItemID = Val(flxSum.TextMatrix(ActRow, 10))
    InvHistory.Show vbModal

End Sub

Private Sub mnuRetStatus_Click()
    Dim rsCount As New ADODB.Recordset
    Dim MyStat As String, Qry As String, MyRec As Integer
    
    MyStat = InputBox("Enter a new Status of " & flxSum.TextMatrix(ActRow, 4) & " Serial#:" & flxSum.TextMatrix(ActRow, 5) & " [Max. 20 characters]", "Update Status")
    If MyStat = "" Then Exit Sub
    
    Qry = "SELECT invtreg.ItemID From invtreg Where invtreg.TRANSID = " & Val(flxSum.TextMatrix(ActRow, 9)) & " AND invtreg.REFERENCE = '" & flxSum.TextMatrix(ActRow, 2) & "';"
    With rsCount
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        MyRec = .RecordCount
        .Close
    End With
    
    If MyRec = 1 Then
        LYRcon.Execute "UPDATE invpureturn SET `STATUS` = '" & Left$(MyStat, 50) & "' Where invpureturn.PURCHRETID = " & Val(flxSum.TextMatrix(ActRow, 9)) & ";"
    Else
        LYRcon.Execute "UPDATE invpretdtls SET REMARKS = '" & Left$(MyStat, 50) & "' Where invpretdtls.PURCHRETID = " & Val(flxSum.TextMatrix(ActRow, 9)) & " AND invpretdtls.ItemID = " & Val(flxSum.TextMatrix(ActRow, 10)) & ";"
        LYRcon.Execute "UPDATE invpureturn SET `STATUS` = '*" & Left$(MyStat, 50) & "' Where invpureturn.PURCHRETID = " & Val(flxSum.TextMatrix(ActRow, 9)) & ";"
    End If
    flxSum.TextMatrix(ActRow, 7) = MyStat
    MsgBox "New Status of " & flxSum.TextMatrix(ActRow, 4) & " Serial#:" & flxSum.TextMatrix(ActRow, 5) & " successfully updated!", vbInformation + vbOKOnly, "Purchase Return"
End Sub



Private Sub mnuRFPRec_Click()
Dim TRFPRecRs As New ADODB.Recordset

If mySrc = "RFP" Then
    txtTRFPNO.Text = flxSum.TextMatrix(ActRow, 2)
    If flxSum.TextMatrix(ActRow, 10) <> "" Then
        Qry = "call sp_acctpayinvcng.sp_load_rfp_treasury('" & flxSum.TextMatrix(ActRow, 2) & "')"
        With TRFPRecRs
            If .State = adStateOpen Then .Close
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                txtTRecBy.Text = IIf(IsNull(!receivedby), "", !receivedby)
                txtTRecDate.Text = Format(!RFPTDATE, "MM/DD/YYYY")
                txtTBankAcct.Text = IIf(IsNull(!BANKACCNT), "", !BANKACCNT)
                txtTChckNo.Text = IIf(IsNull(!CheckNo), "", !CheckNo)
                txtTChckDate.Text = IIf(IsNull(!checkdate), "__/__/____", Format(!checkdate, "MM/DD/YYYY"))
                txtTCVNo.Text = IIf(IsNull(!CVno), "", !CVno)
                txtTDateDeposit.Text = IIf(IsNull(!DATEDEPOSIT), "__/__/____", Format(!DATEDEPOSIT, "MM/DD/YYYY"))
            End If
        End With
    Else
        txtTRecBy.Text = EmployeesName
        txtTRecDate.Text = Format(Now, "mm/dd/yyyy")
        txtTBankAcct.Text = ""
        txtTChckNo.Text = ""
        txtTChckDate.Text = "__/__/____"
        txtTCVNo.Text = ""
        txtTDateDeposit.Text = "__/__/____"
    End If
    
    With ssTabTRec
        .Top = (Me.Height - .Height) \ 2
        .Left = (Me.Width - .Width) \ 2
        .Visible = True
        .ZOrder
    End With
Else
    MsgBox "NOT AVAILABLE!", vbInformation + vbOKOnly, "VALIDATION"
End If
End Sub

Private Sub mnuViewCancel_Click()
    Dim rsCancel As New ADODB.Recordset
    Dim Qry As String, MyID As Integer, MyDate As String, myReason As String, MyUser As String
                
    Qry = "SELECT * From canceltrans Where canceltrans.TRANSDATE = '" & Format(DateValue(flxSum.TextMatrix(ActRow, 1)), "yyyy-mm-dd") & "' AND canceltrans.TRANSNAME = 'OFFICIAL RECEIPT' AND canceltrans.REFNO LIKE '" & flxSum.TextMatrix(ActRow, 2) & "%'"
    With rsCancel
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MyDate = Format(!CancelDate, "mm/dd/yyyy")
            myReason = !REASON
            MyID = !CancelledBy
        End If
        .Close
    End With
    'Display
    Qry = "SELECT CONCAT(employees.FIRSTNAME, ' ', employees.LASTNAME) AS MyEmp From employees Where employees.EMPID = " & MyID & ";"
    MyUser = GetFieldValue(Qry)
    MsgBox "Date Cancelled: " & MyDate & vbCrLf & "Reason for Cancel: " & myReason & vbCrLf & "Cancelled By: " & MyUser, vbInformation + vbOKOnly, "Cancel Info"
    
End Sub


Private Sub RXResizer1_ResizeComplete()
    Dim Qry As String, i As Integer, j As Integer
    
    lblHeading(2).Left = Frame1.Left + (Frame1.Width - lblHeading(2).Width) - 30
    lblHeading(3).Left = Frame1.Left + (Frame1.Width - lblHeading(3).Width)
    
    With flxSum
        If mySrc = "SP Incoming" Then
            .FormatString = "   |^DATE REC.|<PACKING NO.|<SUPPLIER|^PACK DATE |<SHIPPER |^WAYBILLNO |^DATE FINAL |^DATE CONFIRM |<REMARKS|<CHECKER|<PRINT COPY"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.12    'DATE REC
            .ColWidth(2) = (.Width - 300) * 0.12    'PACKING NO
            .ColWidth(3) = (.Width - 300) * 0.17    'SUPPLIER
            .ColWidth(4) = (.Width - 300) * 0.17    'PACK DATE
            .ColWidth(5) = (.Width - 300) * 0.12    'SHIPPER
            .ColWidth(6) = (.Width - 300) * 0.12    'WAYBILLNO
            .ColWidth(7) = (.Width - 300) * 0.12    'DATE FINAL
            .ColWidth(8) = (.Width - 300) * 0.12    'DATE CONFIRM
            .ColWidth(9) = (.Width - 300) * 0.15    'REMARKS
            .ColWidth(10) = (.Width - 300) * 0.15    'CHECKER
            .ColWidth(11) = (.Width - 300) * 0.15    'PRINT COPY
            .ColWidth(12) = 0
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 11)) = 0 Then 'Not yet printed
                    .col = 1: .row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                    .col = 2
                End If
                If InStr(.TextMatrix(i, 9), "CANCELLED") > 0 Then 'Cancelled
                    .col = 1: .row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
                    .col = 2
                End If
            Next i
        ElseIf mySrc = "Advances for Liquidation" Then
            .FormatString = "  |^AFL Date |^AFL No. |>Amt. Rec.  |>Amt. Ret. |>AFL Amount |<Prepared By |<Remarks |^Prn Copy|"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 400) * 0.1
            .ColWidth(2) = (.Width - 400) * 0.1
            .ColWidth(3) = (.Width - 400) * 0.1
            .ColWidth(4) = (.Width - 400) * 0.1
            .ColWidth(5) = (.Width - 400) * 0.1
            .ColWidth(6) = (.Width - 400) * 0.2
            .ColWidth(7) = (.Width - 400) * 0.2
            .ColWidth(8) = (.Width - 400) * 0.08
            .ColWidth(9) = 0
            For i = 1 To .Rows - 1
                If InStr(UCase$(.TextMatrix(i, 7)), "CANCELLED") > 0 Then 'Cancelled
                    .col = 1: .row = i: .ColSel = TCol: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
                    .col = 2
                End If
                If Val(.TextMatrix(i, 8)) = 0 Then
                    .col = 1: .row = i: .ColSel = TCol: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                    .col = 2
                End If
            Next i
        
        ElseIf mySrc = "Sales Return" Then
            .FormatString = "   |^DATE RETURNED|^SALES RET.#|^INVOICE#|^INVOICE DATE|<NAME OF CUSTOMER|^SALES RET.THIS MONTH|^SALES RET.-PRIOR MONTH SALES|>TOTAL|<REMARKS |^PRINTED"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.15  'Date return
            .ColWidth(2) = (.Width - 300) * 0.12  'Sales RET no
            .ColWidth(3) = (.Width - 300) * 0.12  'Invoice No
            .ColWidth(4) = (.Width - 300) * 0.12  'Invoice date
            .ColWidth(5) = (.Width - 300) * 0.25  'Name of Customer
            .ColWidth(6) = (.Width - 300) * 0.2  'Sales Return This Month
            .ColWidth(7) = (.Width - 300) * 0.23  'Sales Return Prior Month Sales
            .ColWidth(8) = (.Width - 300) * 0.06  'Total
            .ColWidth(9) = (.Width - 300) * 0.26  'REMARKS
            .ColWidth(10) = (.Width - 300) * 0.08 'Printed
            .ColWidth(11) = 0
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                .TextMatrix(i, 4) = Format$(.TextMatrix(i, 4), "mm/dd/yyyy")
                .TextMatrix(i, 6) = Format$(.TextMatrix(i, 6), "#,##0.00")
                .TextMatrix(i, 7) = Format$(.TextMatrix(i, 7), "#,##0.00")
                .TextMatrix(i, 8) = Format$(.TextMatrix(i, 8), "#,##0.00")
                If Val(.TextMatrix(i, 10)) = 0 Then  'PRINT COPY
                    .col = 1: .row = i: .ColSel = 10: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .col = 2
                End If
                If .TextMatrix(i, 9) = "CANCELLED" Then 'REMARKS
                    .col = 1: .row = i: .ColSel = 10: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(240, 0, 0)
                    .col = 2
                End If
            Next i
        ElseIf mySrc = "Purchase Return" Then
            .FormatString = "   |^RET. DATE|^RETURN NO.|^PACKING NO.|^PACK DATE|<SUPPLIER|>AMOUNT|<REMARKS |^PRINTED"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1
            .ColWidth(2) = (.Width - 300) * 0.1
            .ColWidth(3) = (.Width - 300) * 0.1
            .ColWidth(4) = (.Width - 300) * 0.1
            .ColWidth(5) = (.Width - 300) * 0.2
            .ColWidth(6) = (.Width - 300) * 0.1
            .ColWidth(7) = (.Width - 300) * 0.2
            .ColWidth(8) = (.Width - 300) * 0.07
            .ColWidth(9) = 0
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                .TextMatrix(i, 4) = Format$(.TextMatrix(i, 4), "mm/dd/yyyy")
                .TextMatrix(i, 6) = Format$(.TextMatrix(i, 6), "#,##0.00")
                If Val(.TextMatrix(i, 8)) = 0 Then  'PRINT COPY
                    .col = 1: .row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .col = 2
                End If
                If .TextMatrix(i, 7) = "CANCELLED" Then 'REMARKS
                    .col = 1: .row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(240, 0, 0)
                    .col = 2
                End If
            Next i
        ElseIf mySrc = "Inventory Adjustment" Then
                             '       1       2           3       4       5           6           7           8
            .FormatString = "   |^ DATE|^ADJ. NO.|<REFERENCE|<REASON|<REMARKS |<PREPARED BY|<APPROVED BY|<PRINT COPY |^ADJTYPE"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.12
            .ColWidth(2) = (.Width - 300) * 0.12
            .ColWidth(3) = (.Width - 300) * 0.12
            .ColWidth(4) = (.Width - 300) * 0.2
            .ColWidth(5) = (.Width - 300) * 0.2
            .ColWidth(6) = (.Width - 300) * 0.2
            .ColWidth(7) = (.Width - 300) * 0.2
            .ColWidth(8) = (.Width - 300) * 0.15
            .ColWidth(9) = (.Width - 300) * 0.1
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                If Val(.TextMatrix(i, 8)) = 0 Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .col = 1: .row = i: .ColSel = 9: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .col = 2
                End If
                If .TextMatrix(i, 5) = "CANCELLED" Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .col = 1: .row = i: .ColSel = 9: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 0, 0)
                    .col = 2
                End If
            Next i
'************** CODE UPDATED: 02/12/2019 *********************
        ElseIf mySrc = "WS Inventory Adjustment" Then
                             '       1       2           3       4       5           6           7           8
            .FormatString = "   |^ DATE |^ADJ. NO.|^STATUS |<ADJTYPE |<REASON|<REMARKS |<PREPARED BY|<APPROVED BY "
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1
            .ColWidth(2) = (.Width - 300) * 0.12
            .ColWidth(3) = (.Width - 300) * 0.1
            .ColWidth(4) = (.Width - 300) * 0.15
            .ColWidth(5) = (.Width - 300) * 0.2
            .ColWidth(6) = (.Width - 300) * 0.2
            .ColWidth(7) = (.Width - 300) * 0.2
            .ColWidth(8) = (.Width - 300) * 0.2
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                If .TextMatrix(i, 3) <> "CANCELLED" And .TextMatrix(i, 3) <> "FINAL" Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .col = 1: .row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .col = 2
                ElseIf .TextMatrix(i, 3) = "CANCELLED" Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .col = 1: .row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 0, 0)
                    .col = 2
                End If
            Next i
'********** END CODE UPDATED: 02/12/2019 *********************

        ElseIf mySrc = "Cashier Cash Flow" Then
            .FormatString = "  |^Cash Flow Date |<Name of Cashier |< Approved By  |<Remarks "
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 400) * 0.15
            .ColWidth(2) = (.Width - 400) * 0.25
            .ColWidth(3) = (.Width - 400) * 0.25
            .ColWidth(4) = (.Width - 400) * 0.3
            .ColWidth(5) = 0
        
        ElseIf mySrc = "Manager's Cash Flow" Then
            .FormatString = "  |^Cash Flow Date |<Prepared By |< Approved By  |<Remarks "
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 400) * 0.15
            .ColWidth(2) = (.Width - 400) * 0.25
            .ColWidth(3) = (.Width - 400) * 0.25
            .ColWidth(4) = (.Width - 400) * 0.3
            .ColWidth(5) = 0
            
        ElseIf mySrc = "Branch Order" Then
            '1-ORDERDATE, 2-REFERENCE, 3-SUPPLIERNAME, 4-ORDERTYPE,5-CONFIRMNO,6-SUBMITTEDBY,7-APPROVEDBY,8-PRINTCOPY,9-QTYORDER, 10-QTYREC,11-BACKORDER,12-BOLASTUPDATE "
            '.FormatString = "   |^ORDER DATE|^RR NO.|<SUPPLIER NAME|^ORDER TYPE|^CONFIRM NO.|<SUBMITTED BY|<APPROVED BY|^PRN COPY"
            .FormatString = "   |^ORDER DATE|^RR NO.|^WAREHOUSE STATUS|^CONFIRM NO.|^TOTAL PN|^REC PN|^BO PN|^TOTAL QTY.|^REC QTY|^BO QTY.|^LEAD TIME|<PICKER|<PRINT COPY"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.08
            .ColWidth(2) = (.Width - 300) * 0.08
            .ColWidth(3) = (.Width - 300) * 0.15
            .ColWidth(4) = (.Width - 300) * 0.08
            .ColWidth(5) = (.Width - 300) * 0.08
            .ColWidth(6) = (.Width - 300) * 0.08
            .ColWidth(7) = (.Width - 300) * 0.08
            .ColWidth(8) = (.Width - 300) * 0.08
            .ColWidth(9) = (.Width - 300) * 0.08
            .ColWidth(10) = (.Width - 300) * 0.08
            .ColWidth(11) = (.Width - 300) * 0.08
            .ColWidth(12) = (.Width - 300) * 0.1
            .ColWidth(12) = (.Width - 300) * 0.1
            .ColWidth(13) = (.Width - 300) * 0.08
            For i = 1 To .Rows - 1
                '.TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                '.TextMatrix(i, 6) = Format$(.TextMatrix(i, 6), "mm/dd/yyyy")
                '.TextMatrix(i, 9) = Format$(.TextMatrix(i, 9), "mm/dd/yyyy")
                If Val(.TextMatrix(i, 13)) = 0 Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .col = 1: .row = i: .ColSel = 13: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .col = 2
                End If
                If .TextMatrix(i, 4) = "CANCELLED" Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .col = 1: .row = i: .ColSel = 13: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 0, 0)
                    .col = 2
                End If
            Next i
            
                
                
        ElseIf mySrc = "Delivery Receipt" Then
            'DRDATE, '%c/%d/%Y'), DRNO, DELIVERTO, DOCREFNO, DATE_FORMAT(DOCREFDATE, '%c/%d/%Y'), RECBY, REMARKS, PRNCOPY, DRID
            .FormatString = "   |^DR DATE|^DR NO.|<DELIVERED TO|^DOC. REF#|^DOC. DATE|<RECEIVED BY|<REMARKS|^PRN COPY"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.08
            .ColWidth(2) = (.Width - 300) * 0.08
            .ColWidth(3) = (.Width - 300) * 0.2
            .ColWidth(4) = (.Width - 300) * 0.08
            .ColWidth(5) = (.Width - 300) * 0.08
            .ColWidth(6) = (.Width - 300) * 0.17
            .ColWidth(7) = (.Width - 300) * 0.18
            .ColWidth(8) = (.Width - 300) * 0.1
            .ColWidth(9) = 0
            For i = 1 To .Rows - 1
                '.TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                '.TextMatrix(i, 6) = Format$(.TextMatrix(i, 6), "mm/dd/yyyy")
                '.TextMatrix(i, 9) = Format$(.TextMatrix(i, 9), "mm/dd/yyyy")
                If Val(.TextMatrix(i, 8)) = 0 Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .col = 1: .row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .col = 2
                End If
                If .TextMatrix(i, 7) = "CANCELLED" Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .col = 1: .row = i: .ColSel = 9: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 0, 0)
                    .col = 2
                End If
            Next i
        
        ElseIf mySrc = "Item Entry" Then
            'SELECT PARTNUMBER, DESCRIPTION,QTY, COST, CLASSIFICATION, LOCATION, SUPPLIER, ITEMID
            .FormatString = "   |<PARTNUMBER|<DESCRIPTION|^QTY.|^UNIT|>COST|<CLASSIFICATION|<LOCATION|^SUPPLIER"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.15
            .ColWidth(2) = (.Width - 300) * 0.2
            .ColWidth(3) = (.Width - 300) * 0.08
            .ColWidth(4) = (.Width - 300) * 0.08
            .ColWidth(5) = (.Width - 300) * 0.08
            .ColWidth(6) = (.Width - 300) * 0.15
            .ColWidth(7) = (.Width - 300) * 0.15
            .ColWidth(8) = (.Width - 300) * 0.15
            .ColWidth(9) = 0
        ElseIf mySrc = "Supplier Entry" Then
            'supplier.SUPPLIERNAME, supplier.ADDRESS, supplier.CONTACTPERS, supplier.TELNO, supplier.EMAIL, supplier.TIN, supplier.SUPPLIERID "
            .FormatString = "   |<SUPPLIER NAME|<ADDRESS|<CONTACT PERSON|<CONTACT NO|<EMAIL|<TIN"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.18
            .ColWidth(2) = (.Width - 300) * 0.18
            .ColWidth(3) = (.Width - 300) * 0.15
            .ColWidth(4) = (.Width - 300) * 0.15
            .ColWidth(5) = (.Width - 300) * 0.15
            .ColWidth(6) = (.Width - 300) * 0.15
            .ColWidth(7) = 0
        ElseIf mySrc = "Branch Entry" Then
            '1-BRANCHNAME, 2-ADDRESS, 3-COMPANYGROUP, 4-MANAGER, 5-TELNO, 6-EMAIL, 7-BRANCHID
            .FormatString = "   |<BRANCH NAME|<ADDRESS|<COMPANY GROUP|<MANAGER|<TEL. NO.|<EMAIL|<TIN"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.18
            .ColWidth(2) = (.Width - 300) * 0.18
            .ColWidth(3) = (.Width - 300) * 0.18
            .ColWidth(4) = (.Width - 300) * 0.18
            .ColWidth(5) = (.Width - 300) * 0.15
            .ColWidth(6) = (.Width - 300) * 0.18
            .ColWidth(7) = 0
        ElseIf mySrc = "Transfer Voucher" Then
            .FormatString = "   |^T.V. DATE|^T.V. NO.|<TRANSFERRED TO|<ADDRESS|<REMARKS |<PREPARED BY|<APPROVED BY|^PRN COPY|^TYPE"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1
            .ColWidth(2) = (.Width - 300) * 0.1
            .ColWidth(3) = (.Width - 300) * 0.2
            .ColWidth(4) = (.Width - 300) * 0.2
            .ColWidth(5) = (.Width - 300) * 0.2
            .ColWidth(6) = (.Width - 300) * 0.15
            .ColWidth(7) = (.Width - 300) * 0.15
            .ColWidth(8) = (.Width - 300) * 0.07
            .ColWidth(9) = (.Width - 300) * 0.07
            .ColWidth(10) = 0
            For i = 1 To .Rows - 1
                '.TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                If Val(.TextMatrix(i, 8)) = 0 Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .col = 1: .row = i: .ColSel = 9: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .col = 2
                End If
                If InStr(.TextMatrix(i, 5), "CANCELLED") > 0 Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .col = 1: .row = i: .ColSel = 9: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 0, 0)
                    .col = 2
                End If
            Next i
        ElseIf mySrc = "Inventory" Then
           '1-PARTNUMBER,2-DESCRIPTION,3-QTY,4-UM1,5-COST,6-CLASSIFICATION,7-LOCATION,8-SUPPLIER,9-REORDERPOINT,10-ITEMID
            .FormatString = "   |<PARTNUMBER |<DESCRIPTION |^QTY|^UNIT M.|>COST |<CLASSIFICATION|<LOCATION|<SUPPLIER|^REORDER POINT"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.12
            .ColWidth(2) = (.Width - 300) * 0.18
            .ColWidth(3) = (.Width - 300) * 0.08
            .ColWidth(4) = (.Width - 300) * 0.1
            .ColWidth(5) = (.Width - 300) * 0.1
            .ColWidth(6) = (.Width - 300) * 0.12
            .ColWidth(7) = (.Width - 300) * 0.12
            .ColWidth(8) = (.Width - 300) * 0.12
            .ColWidth(9) = (.Width - 300) * 0.08
            .ColWidth(10) = 0 '(.Width - 300) * 0.2
        ElseIf mySrc = "Detailed Receiving" Then
            '1-PARTNUMBER, 2-DESCRIPTION, 3-REFERENCE, 4-TRANSDATE, 5-Qty, 6-UNITPRICE, 7-TtlAmt, 8-SOURCE, 9-SourceDo
            .FormatString = "   |<PARTNUMBER  |<DESCRIPTION|^REF. NO.|^REF. DATE |^QTY.|>PRICE|>TOTAL AMT|<RECEIVED FROM |<SOURCE DOC."
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.15
            .ColWidth(2) = (.Width - 300) * 0.2
            .ColWidth(3) = (.Width - 300) * 0.07
            .ColWidth(4) = (.Width - 300) * 0.07
            .ColWidth(5) = (.Width - 300) * 0.07
            .ColWidth(6) = (.Width - 300) * 0.09
            .ColWidth(7) = (.Width - 300) * 0.09
            .ColWidth(8) = (.Width - 300) * 0.18
            .ColWidth(9) = (.Width - 300) * 0.1
        ElseIf mySrc = "Detailed Outgoing" Then
        '1-PARTNUMBER, 2-DESCRIPTION, 3-REFERENCE, 4-TRANSDATE, 5-OUT, 6-UNITPRICE, 7-TtlAmt, 8-receiver, 9-Branch/Dept, 10-SourceDoc
            .FormatString = "   |<PARTNUMBER  |<DESCRIPTION|^REF. NO.|^REF. DATE |^QTY.|>PRICE|>TOTAL AMT|<RECEIVER |<BRANCH/DEPT. |<SOURCE DOC."
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.15
            .ColWidth(2) = (.Width - 300) * 0.2
            .ColWidth(3) = (.Width - 300) * 0.07
            .ColWidth(4) = (.Width - 300) * 0.07
            .ColWidth(5) = (.Width - 300) * 0.07
            .ColWidth(6) = (.Width - 300) * 0.09
            .ColWidth(7) = (.Width - 300) * 0.09
            .ColWidth(8) = (.Width - 300) * 0.18
            .ColWidth(9) = (.Width - 300) * 0.18
            .ColWidth(10) = (.Width - 300) * 0.12
        
        ElseIf mySrc = "Spare Parts Receiving" Then
            .FormatString = "  |^DATE REC. |^PACKING NO. |<RECEIVED FROM |<DELIVERED BY  |<PACKING DATE |<REMARKS  |<CHECKED BY |^PRINT "
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 400) * 0.1
            .ColWidth(2) = (.Width - 400) * 0.1
            .ColWidth(3) = (.Width - 400) * 0.2
            .ColWidth(4) = (.Width - 400) * 0.18
            .ColWidth(5) = (.Width - 400) * 0.12
            .ColWidth(6) = (.Width - 400) * 0.2
            .ColWidth(7) = (.Width - 400) * 0.18
            .ColWidth(8) = (.Width - 400) * 0.07
            .ColWidth(9) = 0
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 8)) = 0 Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .col = 1: .row = i: .ColSel = TCol - 1: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .col = 2
                End If
                If .TextMatrix(i, 6) = "CANCELLED" Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .col = 1: .row = i: .ColSel = TCol - 1: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 0, 0)
                    .col = 2
                End If
            Next i
        ElseIf mySrc = "Spareparts Invoice" Then
            .FormatString = "  |INV. DATE  |^INV. NO.  |<CUSTOMER   |<ADDRESS     |> INV. AMOUNT |<INV. TYPE  |<TYPE OF CUSTOMER|<FRONTLINER |<REMARKS |^PRINT"

            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 400) * 0.1
            .ColWidth(2) = (.Width - 400) * 0.1
            .ColWidth(3) = (.Width - 400) * 0.18
            .ColWidth(4) = (.Width - 400) * 0.2
            .ColWidth(5) = (.Width - 400) * 0.1
            .ColWidth(6) = (.Width - 400) * 0.15
            .ColWidth(7) = (.Width - 400) * 0.15
            .ColWidth(8) = (.Width - 400) * 0.18
            .ColWidth(9) = (.Width - 400) * 0.18
            .ColWidth(10) = (.Width - 400) * 0.08
            .ColWidth(11) = 0
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 10)) = 0 Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .col = 1: .row = i: .ColSel = TCol - 1: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .col = 2
                End If
                If .TextMatrix(i, 9) = "CANCELLED" Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .col = 1: .row = i: .ColSel = TCol - 1: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 0, 0)
                    .col = 2
                End If
            Next i
    
        ElseIf mySrc = "Spareparts Location Adj." Then
            .FormatString = "   |^ DATE|^CONTROL NO.|<REFERENCE|<REASON|<PREPARED BY|<APPROVED BY|^PRINT COPY"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.12
            .ColWidth(2) = (.Width - 300) * 0.12
            .ColWidth(3) = (.Width - 300) * 0.2
            .ColWidth(4) = (.Width - 300) * 0.4
            .ColWidth(5) = (.Width - 300) * 0.15
            .ColWidth(6) = (.Width - 300) * 0.15
            .ColWidth(7) = (.Width - 300) * 0.1
            .ColWidth(8) = 0
            If .Rows >= 2 And .TextMatrix(1, 1) <> "" Then
                For i = 1 To .Rows - 1
                    .TextMatrix(i, 2) = Format$(.TextMatrix(i, 2), "0000")
                    If Val(.TextMatrix(i, 7)) = 0 Then
                        '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                        .col = 1: .row = i: .ColSel = .Cols - 1: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                        .col = 2
                    End If
                    If .TextMatrix(i, 3) = "CANCELLED" Then
                        .col = 1: .row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 0, 0)
                        .col = 2
                    End If
                Next i
            End If
        ElseIf mySrc = "Cash Voucher" Then
            .FormatString = "  |^  CV Date |^Cash Voucher# |< CV Type |>Amount  |< Name of Payee |Particulars  |<Remarks |^Prn Copy|"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1
            .ColWidth(2) = (.Width - 300) * 0.1
            .ColWidth(3) = (.Width - 300) * 0.15
            .ColWidth(4) = (.Width - 300) * 0.1
            .ColWidth(5) = (.Width - 300) * 0.17
            .ColWidth(6) = (.Width - 300) * 0.17
            .ColWidth(7) = (.Width - 300) * 0.17
            .ColWidth(8) = (.Width - 300) * 0.07
            .ColWidth(9) = 0
            For i = 1 To .Rows - 1
                'If Len(.TextMatrix(i, 2)) < 4 Then .TextMatrix(i, 2) = Format(.TextMatrix(i, 2), "0000")
                If InStr(UCase$(.TextMatrix(i, 7)), "CANCELLED") > 0 Then 'Cancelled
                    .col = 1: .row = i: .ColSel = 9: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
                    .col = 2
                End If
                If Val(.TextMatrix(i, 8)) = 0 Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10 ': .Text = lblP.Caption
                    .col = 1: .row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                    .col = 2
                End If
            Next i
        ElseIf mySrc = "Redeemed Unit" Then
            .FormatString = "  |^ DR DATE |^DR NUMBER |< CUSTOMER  |<MODEL  |< ACCOUNT NO.  |^DMCM NO.  |<REMARKS |<SERIAL NO."
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1
            .ColWidth(2) = (.Width - 300) * 0.1
            .ColWidth(3) = (.Width - 300) * 0.2
            .ColWidth(4) = (.Width - 300) * 0.15
            .ColWidth(5) = (.Width - 300) * 0.12
            .ColWidth(6) = (.Width - 300) * 0.12
            .ColWidth(7) = (.Width - 300) * 0.25
            .ColWidth(8) = (.Width - 300) * 0.15
        
        ElseIf mySrc = "DAILY P.R." Then
            .FormatString = "  |^ P.R. DATE |^P.R. NUMBER |< CUSTOMER NAME |< ACCOUNT NO.  |^P.R. AMOUNT  |<REMARKS "
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1
            .ColWidth(2) = (.Width - 300) * 0.1
            .ColWidth(3) = (.Width - 300) * 0.25
            .ColWidth(4) = (.Width - 300) * 0.15
            .ColWidth(5) = (.Width - 300) * 0.12
            .ColWidth(6) = (.Width - 300) * 0.22
            .ColWidth(7) = 0
        ElseIf mySrc = "Inventory TRA" Then
            .FormatString = "  |^ TRA DATE |^TRA NUMBER |^TRA TYPE|< CUSTOMER  |<ADDRESS  |< INVOICE NO.  |<REMARKS |^PRINT"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1
            .ColWidth(2) = (.Width - 300) * 0.1
            .ColWidth(3) = (.Width - 300) * 0.15
            .ColWidth(4) = (.Width - 300) * 0.15
            .ColWidth(5) = (.Width - 300) * 0.15
            .ColWidth(6) = (.Width - 300) * 0.1
            .ColWidth(7) = (.Width - 300) * 0.15
            .ColWidth(8) = (.Width - 300) * 0.07
            .ColWidth(9) = 0
            For i = 1 To .Rows - 1
                'If Len(.TextMatrix(i, 2)) < 4 Then .TextMatrix(i, 2) = Format(.TextMatrix(i, 2), "0000")
                If InStr(UCase$(.TextMatrix(i, 7)), "CANCELLED") > 0 Then 'Cancelled
                    .col = 1: .row = i: .ColSel = 9: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
                    .col = 2
                End If
                If Val(.TextMatrix(i, 8)) = 0 Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10 ': .Text = lblP.Caption
                    .col = 1: .row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                    .col = 2
                End If
            Next i
        ElseIf mySrc = "Deed of Redemption" Then
            .FormatString = "  |^ REF. DATE |^REF. NO. |^DMCM NO.|< CUSTOMER  |<ADDRESS  |^DATE DEPOSIT  |<REMARKS |^PRINT"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1
            .ColWidth(2) = (.Width - 300) * 0.1
            .ColWidth(3) = (.Width - 300) * 0.12
            .ColWidth(4) = (.Width - 300) * 0.16
            .ColWidth(5) = (.Width - 300) * 0.16
            .ColWidth(6) = (.Width - 300) * 0.1
            .ColWidth(7) = (.Width - 300) * 0.16
            .ColWidth(8) = (.Width - 300) * 0.07
            .ColWidth(9) = 0
            For i = 1 To .Rows - 1
                'If Len(.TextMatrix(i, 2)) < 4 Then .TextMatrix(i, 2) = Format(.TextMatrix(i, 2), "0000")
                If InStr(UCase$(.TextMatrix(i, 7)), "CANCELLED") > 0 Then 'Cancelled
                    .col = 1: .row = i: .ColSel = 9: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
                    .col = 2
                End If
                If Val(.TextMatrix(i, 8)) = 0 Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10 ': .Text = lblP.Caption
                    .col = 1: .row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                    .col = 2
                End If
            Next i
        ElseIf mySrc = "Certificate of Ownership" Then
            .FormatString = "  |^DATE ISSUED |^CERTIFICATE NO. |< CUSTOMER NAME  |<ITEM DESCRIPTION  |<ACCOUNT STATUS  |<SERIAL NO. |<REMARKS "
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1
            .ColWidth(2) = (.Width - 300) * 0.15
            .ColWidth(3) = (.Width - 300) * 0.2
            .ColWidth(4) = (.Width - 300) * 0.2
            .ColWidth(5) = (.Width - 300) * 0.15
            .ColWidth(6) = (.Width - 300) * 0.15
            .ColWidth(7) = (.Width - 300) * 0.15
            .ColWidth(8) = 0
            For i = 1 To .Rows - 1
                If InStr(UCase$(.TextMatrix(i, 7)), "CANCELLED") > 0 Then 'Cancelled
                    .col = 1: .row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
                    .col = 2
                End If
            Next i
        ElseIf mySrc = "Monthly Commission" Then
            .FormatString = "  |<NAME OF AGENT |^PCV DATE |^PCV NO.  |<PARTICULARS |>GROSS AMT.|>10% EXPANDED |>NET AMOUNT "
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.2
            .ColWidth(2) = (.Width - 300) * 0.1
            .ColWidth(3) = (.Width - 300) * 0.1
            .ColWidth(4) = (.Width - 300) * 0.28
            .ColWidth(5) = (.Width - 300) * 0.1
            .ColWidth(6) = (.Width - 300) * 0.1
            .ColWidth(7) = (.Width - 300) * 0.1
        ElseIf mySrc = "Branch Purchase Order" Then
            cmdUpdate.Visible = True
            '               0      1          2               3              4         5          6         7                8     9  10    11
            '.FormatString = "   |^SRV NO. |^BRANCH NAME |^ P.O REFERENCE |^DATE SRV|^DATE CONF.|^STATUS|^TRANSFER#|^TOTAL AMOUNT|^BID|^PICK^FQT"
            '               0         1             2                3              4              5          6         7         8
            .FormatString = "   |<BRANCH NAME |^ORDER REFERENCE |^ORDER DATE|^WAREHOUSE STATUS |^ORDER TYPE|^TOTAL PN|^TOTAL QTY|^LEAD TIME|^PICKER|^PRINT"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.15
            .ColWidth(2) = (.Width - 300) * 0.12
            .ColWidth(3) = (.Width - 300) * 0.11
            .ColWidth(4) = (.Width - 300) * 0.13
            .ColWidth(5) = (.Width - 300) * 0.1
            .ColWidth(6) = (.Width - 300) * 0.06
            .ColWidth(7) = (.Width - 300) * 0.06
            .ColWidth(8) = (.Width - 300) * 0.13
            .ColWidth(9) = (.Width - 300) * 0.13
            .ColWidth(10) = (.Width - 300) * 0.08
            .ColWidth(11) = 0
            .ColWidth(12) = 0
            .ColWidth(13) = 0
            .ColWidth(14) = 0
            For i = 1 To .Rows - 1
                 If Val(.TextMatrix(i, 10)) = 0 Then 'NoT Printed for Picking
                      .col = 1: .row = i: .ColSel = 10: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                      .col = 2
                      '.TextMatrix(i, 6) = "PENDING"
                End If
                If Val(.TextMatrix(i, 12)) = 1 Then 'Final for Older Taking
                      .col = 1: .row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 300)
                      .col = 2
                      '.TextMatrix(i, 6) = "PENDING"
                End If
                If Val(.TextMatrix(i, 12)) = 1 Then 'Final for Older Taking
                      .col = 1: .row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = &H80000005
                      .col = 2
                      '.TextMatrix(i, 6) = "FINAL"
                End If
                If .TextMatrix(i, 4) = "CANCELLED" Then 'Not Final for Older Taking
                      .col = 1: .row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 199, 199)
                      .col = 2
                      '.TextMatrix(i, 6) = "CANCELLED"
                End If
                
                 If .TextMatrix(i, 4) = "INTRANSIT" Then 'Not Final for Older Taking
                      .col = 1: .row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(199, 254, 255)
                      .col = 2
                      '.TextMatrix(i, 6) = "CANCELLED"
                End If
                If .TextMatrix(i, 1) = "" And .Rows = 2 Then
                      .col = 1: .row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = &H80000005
                      .col = 2
                     '.TextMatrix(i, 6) = ""
                End If
                .TextMatrix(i, 9) = Format(.TextMatrix(i, 9), "###,###,##0.00")
            Next i
        
        
        ElseIf mySrc = "SP PACKING" Then
            '                0       1        2              3               4        5        6           7                  8        9           10         11        12       13
            .FormatString = "  |^REF DATE |^FINAL DATE |^DELIVERY DATE |^REF.NO. |^TTL PN |^TTL QTY |<RECEIVING BRANCH  |<ADDRESS |^P.O. NUMBER |^STATUS |<REMARKS |<CHECKER |<DRIVER "
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.09
            .ColWidth(2) = (.Width - 300) * 0.09
            .ColWidth(3) = (.Width - 300) * 0.09
            .ColWidth(4) = (.Width - 300) * 0.07
            .ColWidth(5) = (.Width - 300) * 0.07
            .ColWidth(6) = (.Width - 300) * 0.07
            .ColWidth(7) = (.Width - 300) * 0.18
            .ColWidth(8) = (.Width - 300) * 0.28
            .ColWidth(9) = (.Width - 300) * 0.1
            .ColWidth(10) = (.Width - 300) * 0.08
            .ColWidth(11) = (.Width - 300) * 0.1
            .ColWidth(12) = (.Width - 300) * 0.1
            .ColWidth(13) = (.Width - 300) * 0.1
            
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = Format(.TextMatrix(i, 1), "mm/dd/yyyy")
                .TextMatrix(i, 2) = Format(.TextMatrix(i, 2), "mm/dd/yyyy")
                .TextMatrix(i, 3) = Format(.TextMatrix(i, 3), "mm/dd/yyyy")
                If .TextMatrix(i, 9) = "" And .TextMatrix(i, 10) = "FINAL" Then
                    .TextMatrix(i, 10) = "INTRANSIT"
                End If
                If .TextMatrix(i, 9) = "" And .TextMatrix(i, 10) = "PENDING" Then
                    .TextMatrix(i, 10) = "PACKING"
                End If
                If .TextMatrix(i, 10) = "PACKING" Then
                        .col = 1: .row = i: .ColSel = 10: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                        .col = 2
                End If
                If .TextMatrix(i, 10) = "CANCELLED" Then
                        .col = 1: .row = i: .ColSel = 10: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
                        .col = 2
                End If
            Next i
            
         ElseIf mySrc = "AllocBBO" Then
            '                0       1                2                 3                 4          5           6
            .FormatString = "  |^ALLOCATION NO. |^ALLOCATION DATE  |<BRANCH NAME  |<PREPARED BY |<APPROVED BY |<STATUS"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1
            .ColWidth(2) = (.Width - 300) * 0.1
            .ColWidth(3) = (.Width - 300) * 0.2
            .ColWidth(4) = (.Width - 300) * 0.1
            .ColWidth(5) = (.Width - 300) * 0.1
            .ColWidth(6) = (.Width - 300) * 0.1
            For i = 1 To .Rows - 1
                Qry = GetFieldValue("SELECT stocktransfer.STATUS From stocktransfer Where stocktransfer.PONO = 'BBO#:" & Val(.TextMatrix(i, 1)) & "' AND (stocktransfer.`STATUS` = 'INTRANSIT' OR stocktransfer.`STATUS` = 'FINAL');")
                    If Qry = "FINAL" Or Qry = "INTRANSIT" Then
                        'MsgBox "Allocation of BBO#:" & .TextMatrix(i, 1) & " was already INTRANSIT", vbCritical, "EDIT"
                        .TextMatrix(i, 6) = Qry
                    End If
            Next i
            
        ElseIf mySrc = "Costing" Then
            With flxSum
                   .FormatString = "   |<SUPPLIER NAME |<DELIVERD BY |^P.O. No. |^PACKING NO. |<PACKING DATE  |^TERMS |^DUE DATE |^TOTAL PART NO. |^TOTAL QTY |>INV TOTAL PAY. |^STATUS"
                   .ColWidth(0) = 440
                   .ColWidth(1) = 3500 'SUPPLIER NAME
                   .ColWidth(2) = 3000 'DELIVERD BY
                   .ColWidth(3) = 1300 'P.O. No.
                   .ColWidth(4) = 1500 'PACKING NO.
                   .ColWidth(5) = 1700 'PACKING DATE
                   .ColWidth(6) = 1000 'TERMS
                   .ColWidth(7) = 1300 'DUE DATE
                   .ColWidth(8) = 1800 'TOTAL PART NO.
                   .ColWidth(9) = 1500 'TOTAL QTY
                   .ColWidth(10) = 1800 'INVOICE TOTAL PAY.
                   .ColWidth(11) = 1000 'STATUS
            
            
            End With
            
            'fraHoldTop.Width = Me.Width - 480
            'flxSum.Width = Me.Width - 480
            'flxSum.Height = Me.Height - 2865
            'txtFind.Left = fraHoldTop.Width - 2055
            'cmdFind.Left = fraHoldTop.Width - 3375
            'cmdSelect.Top = Me.Height - 1080
            'cmdCancel.Top = Me.Height - 1080
            'cmdCancel.Left = Me.Width - 1455
        ElseIf mySrc = "SRP Pricing" Then
            
            With flxSum
                   .FormatString = "   |^REF. NO. |<SUPPLIER NAME |^PACKING NO. |^TOTAL PART NO. |^DATE FINAL |^STATUS |^ENTRY TYPE "
                   .ColWidth(0) = 440
                   .ColWidth(1) = (.Width - 300) * 0.1  'REF. NO.
                   .ColWidth(2) = (.Width - 300) * 0.3  'SUPPLIER NAME
                   .ColWidth(3) = (.Width - 300) * 0.15 'PACKING NO.
                   .ColWidth(4) = (.Width - 300) * 0.11  'TOTAL PART NO.
                   .ColWidth(5) = (.Width - 300) * 0.11 'DATE FINAL
                   .ColWidth(6) = (.Width - 300) * 0.1 'STATUS
                   .ColWidth(7) = (.Width - 300) * 0.1 'ENTRY TYPE
                   .ColWidth(8) = 0 'Supplier Code
            
            End With
            
            'fraHoldTop.Width = Me.Width - 480
            'flxSum.Width = Me.Width - 480
            'flxSum.Height = Me.Height - 2865
            'txtFind.Left = fraHoldTop.Width - 2055
            'cmdFind.Left = fraHoldTop.Width - 3375
            'cmdSelect.Top = Me.Height - 1080
            'cmdCancel.Top = Me.Height - 1080
            'cmdCancel.Left = Me.Width - 1455
'*********************** CODE UPDATED: 08/08/2017 ********************
        ElseIf mySrc = "SETTLEMENT DISCREPANCY" Then
            '                0          1              2          3            4               5         6         7         8        9         10         11           14
            .FormatString = "  |^S.D DATE  |<BRANCH NAME |<STF NO. |<S.D REF NO. |<REF NO. |<REF DATE |^LACKING |^EXCESS |^DAMAGE |^TOTAL PN |<STATUS |<REMARKS |<AUDIT FINDINGS"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.08    'REF DATE
            .ColWidth(2) = (.Width - 300) * 0.25    'BRANCH NAME
            .ColWidth(3) = (.Width - 300) * 0.08    'STF NO.
            .ColWidth(4) = (.Width - 300) * 0.08    'REF NO.
            .ColWidth(5) = (.Width - 300) * 0.08    'WH REF NO.
            .ColWidth(6) = (.Width - 300) * 0.08    'WH REF DATE
            .ColWidth(7) = (.Width - 300) * 0.08    'LACKING
            .ColWidth(8) = (.Width - 300) * 0.08    'EXCESS
            .ColWidth(9) = (.Width - 300) * 0.08    'DAMAGE
            .ColWidth(10) = (.Width - 300) * 0.08   'TOTAL PN
            .ColWidth(11) = (.Width - 300) * 0.08   'STATUS
            .ColWidth(12) = (.Width - 300) * 0.15   'REMARKS
            .ColWidth(13) = 0 '(.Width - 300) * 0.2    'SDISCID
            .ColWidth(14) = 0 '(.Width - 300) * 0.1 'BCODE
            .ColWidth(15) = (.Width - 300) * 0.2 'AUD FINDINGS
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 11) = "NEW" Then
                        .col = 1: .row = i: .ColSel = 15: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                        .col = 2
                End If
            Next i
        ElseIf mySrc = "SDiscrepancy Adjustment" Then
                             '       1       2           3       4       5           6           7           8
            .FormatString = "   |^ DATE|^ADJ. NO.|<REFERENCE|<REASON|<REMARKS |<PREPARED BY|<APPROVED BY|<PRINT COPY"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.12
            .ColWidth(2) = (.Width - 300) * 0.12
            .ColWidth(3) = (.Width - 300) * 0.12
            .ColWidth(4) = (.Width - 300) * 0.2
            .ColWidth(5) = (.Width - 300) * 0.2
            .ColWidth(6) = (.Width - 300) * 0.2
            .ColWidth(7) = (.Width - 300) * 0.2
            .ColWidth(8) = (.Width - 300) * 0.15
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                If Val(.TextMatrix(i, 8)) = 0 Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .col = 1: .row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .col = 2
                End If
                If .TextMatrix(i, 5) = "CANCELLED" Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .col = 1: .row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 0, 0)
                    .col = 2
                End If
            Next i
'******************* END CODE UPDATED: 08/08/2017 ********************
'*********************** CODE UPDATED: 06/14/2019 ********************
        ElseIf mySrc = "PO ADJUSTMENT" Then
                             '       1       2           3       4       5           6           7           8
            .FormatString = "   |^DATE |^STATUS |^ADJ. NO.|^REFERENCE |<SUPPLIER|^REPLACE |^NEW"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1         'DATE
            .ColWidth(2) = (.Width - 300) * 0.12        'STATUS
            .ColWidth(3) = (.Width - 300) * 0.12        'ADJNO
            .ColWidth(4) = (.Width - 300) * 0.15        'REFERENCE
            .ColWidth(5) = (.Width - 300) * 0.25        'SUPPLIER
            .ColWidth(6) = (.Width - 300) * 0.12         'REPLACE
            .ColWidth(7) = (.Width - 300) * 0.12         'NEW
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 2) = "PENDING" Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .col = 1: .row = i: .ColSel = 7: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .col = 2
                End If
                If .TextMatrix(i, 2) = "CANCELLED" Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .col = 1: .row = i: .ColSel = 7: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 0, 0)
                    .col = 2
                End If
            Next i
'******************* END CODE UPDATED: 06/14/2019 ********************
'*********************** CODE UPDATED: 10/04/2019 ********************
        ElseIf mySrc = "RFP" Then
                             '       1          2         3     4            5    6                7                8            9                 10                11                  12            13                 14
            .FormatString = "   |^RFP DATE |^RFP NO. |^STATUS |<SUPPLIER |^TERM |^RFP TYPE |>TOTAL INV. AMOUNT |>TOTAL DEDUCTION |>TOTAL PAYABLE |>SPCRR AMOUNT|<RECEIVED ON(TREASURY) |^BANK ACCOUNT |^CHECK NO |<RECONCILIATION"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1         'RFP DATE
            .ColWidth(2) = (.Width - 300) * 0.1         'RFP NO.
            .ColWidth(3) = (.Width - 300) * 0.1         'STATUS
            .ColWidth(4) = (.Width - 300) * 0.25        'SUPPLIER
            .ColWidth(5) = (.Width - 300) * 0.1         'TERM
            .ColWidth(6) = (.Width - 300) * 0.1         'RFP TYPE
            .ColWidth(7) = (.Width - 300) * 0.14        'TOTAL AMOUNT
            .ColWidth(8) = (.Width - 300) * 0.14        'TOTAL DEDUCTION
            .ColWidth(9) = (.Width - 300) * 0.14        'TOTAL PAYABLE
            .ColWidth(10) = (.Width - 300) * 0.14        'SPCRR AMOUNT
            .ColWidth(11) = (.Width - 300) * 0.1        'RECEIVED ON(TREASURY)
            .ColWidth(12) = (.Width - 300) * 0.14       'BANK ACCOUNT
            .ColWidth(13) = (.Width - 300) * 0.14       'CHECK NO
            .ColWidth(14) = (.Width - 300) * 0.1       'CHECK NO
            
            For i = 1 To .Rows - 1
                .TextMatrix(i, 7) = Format(.TextMatrix(i, 7), "#,###,###,###,###.00")        'TOTAL AMOUNT
                .TextMatrix(i, 8) = Format(.TextMatrix(i, 8), "#,###,###,###,###.00")        'TOTAL DEDUCTION
                .TextMatrix(i, 9) = Format(.TextMatrix(i, 9), "#,###,###,###,###.00")        'TOTAL PAYABLE
                .TextMatrix(i, 10) = Format(.TextMatrix(i, 10), "#,###,###,###,###.00")       'SPCRR AMOUNT
                If .TextMatrix(i, 3) = "PENDING" Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .col = 1: .row = i: .ColSel = 13: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .col = 2
                End If
                If .TextMatrix(i, 3) = "CANCELLED" Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .col = 1: .row = i: .ColSel = 13: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 0, 0)
                    .col = 2
                End If
            Next i
'******************* END CODE UPDATED: 10/04/2019 ********************
'*********************** CODE UPDATED: 10/04/2019 ********************
        ElseIf mySrc = "RFP RECONCILIATION" Then

                             '       1          2       3        4             5            6          7          8            9           10            11            12          13         14
            .FormatString = "   |^RFP DATE |^TERMS |^ORD DATE |^DATE REC |^AGE(PO TO REC)|^PO NO. |<RFP NO. |<INVOICE NO. |<SPCRR NO. |>PO AMOUNT |>INV. AMOUNT |>SPCRR AMOUNT |>BALANCE |^REMARKS"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1         'RFP DATE
            .ColWidth(2) = (.Width - 300) * 0.08        'TERMS
            .ColWidth(3) = (.Width - 300) * 0.1         'ORD DATE
            .ColWidth(4) = (.Width - 300) * 0.1         'DATE REC
            .ColWidth(5) = (.Width - 300) * 0.1         'AGE
            .ColWidth(6) = (.Width - 300) * 0.12        'PO NO.
            .ColWidth(7) = (.Width - 300) * 0.25        'RFP NO.
            .ColWidth(8) = (.Width - 300) * 0.25        'INVOICE NO.
            .ColWidth(9) = (.Width - 300) * 0.25        'SPCRR NO.
            .ColWidth(10) = (.Width - 300) * 0.14       'PO AMOUNT
            .ColWidth(11) = (.Width - 300) * 0.14       'INV. AMOUNT
            .ColWidth(12) = (.Width - 300) * 0.14       'SPCRR AMOUNT
            .ColWidth(13) = (.Width - 300) * 0.14       'SPCRR AMOUNT
            .ColWidth(14) = (.Width - 300) * 0.1        'REMARKS
            
            For i = 1 To .Rows - 1
                .TextMatrix(i, 10) = Format(.TextMatrix(i, 10), "#,###,###,###,###.00")        'PO AMOUNT
                .TextMatrix(i, 11) = Format(.TextMatrix(i, 11), "#,###,###,###,###.00")        'INV AMOUNT
                .TextMatrix(i, 12) = Format(.TextMatrix(i, 12), "#,###,###,###,###.00")        'SPCRR AMOUNT
                .TextMatrix(i, 13) = Format(.TextMatrix(i, 13), "#,###,###,###,###.00")      'BALANCE
                If .TextMatrix(i, 14) = "UNDELIVERED" Or .TextMatrix(i, 14) = "" Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .col = 1: .row = i: .ColSel = 14: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .col = 2
                End If
                If .TextMatrix(i, 14) = "CANCELLED" Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .col = 1: .row = i: .ColSel = 14: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 0, 0)
                    .col = 2
                End If
            Next i
'******************* END CODE UPDATED: 10/04/2019 ********************
        ElseIf mySrc = "ADJUSTMENT REQUEST" Then
            '                     1        2         3           4              5               6                 7             8                     9               10
            .FormatString = "  |^DATE |^STATUS |^TICKET NO. |^CONTROL NO. |^ADJ. NO. |^DATE FINAL |<SUBJECT OF REQUEST |<REQUEST BY |<AUDIT APPROVED BY |<MIS IN-INCHARGE |<SCRIPT FILENAME"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1         'DATE
            .ColWidth(2) = (.Width - 300) * 0.1         'STATUS
            .ColWidth(3) = (.Width - 300) * 0.1         'TICKET NO.
            .ColWidth(4) = (.Width - 300) * 0.1         'CONTROL NO.
            .ColWidth(5) = (.Width - 300) * 0.1         'ADJ NO.
            .ColWidth(6) = (.Width - 300) * 0.1         'DATE FINAL
            .ColWidth(7) = (.Width - 300) * 0.2         'SUBJECT OF REQUEST
            .ColWidth(8) = (.Width - 300) * 0.15        'REQUEST BY
            .ColWidth(9) = (.Width - 300) * 0.15        'AUDIT APPROVED BY
            .ColWidth(10) = (.Width - 300) * 0.15       'MIS IN-INCHARGE
            .ColWidth(11) = (.Width - 300) * 0.15       'SCRIPT FILENAME
        ElseIf mySrc = "CANCELLATION MODULE" Then
            '                           1             2         3           4                 5          6        7       8      9          10            11
            .FormatString = "  |^DATE CREATED |^CANCEL NO. |^STATUS |<TRANSACTION TYPE |^PROG. VALUE |<REASON |<NAME |>AMOUNT |^DATE |<PREPARED BY |<APPROVED BY"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1         'DATE CREATED
            .ColWidth(2) = (.Width - 300) * 0.1         'CANCEL NO.
            .ColWidth(3) = (.Width - 300) * 0.1         'STATUS
            .ColWidth(4) = (.Width - 300) * 0.1         'TRANSACTION TYPE
            .ColWidth(5) = (.Width - 300) * 0.1         'PROG. VALUE
            .ColWidth(6) = (.Width - 300) * 0.2         'REASON
            .ColWidth(7) = (.Width - 300) * 0.15        'NAME
            .ColWidth(8) = (.Width - 300) * 0.1         'AMOUNT
            .ColWidth(9) = (.Width - 300) * 0.1         'DATE
            .ColWidth(10) = (.Width - 300) * 0.15       'PREPARED BY
            .ColWidth(11) = (.Width - 300) * 0.15       'APPROVED BY
        End If
    End With
    
    cmdFind.Height = cmdRefresh.Height
    cmdFind.Top = cmdRefresh.Top
    txtFind.Top = cmdFind.Top
    LoadPriceBCost = True
End Sub


Private Sub tmrAutoLoad_Timer()
On Error GoTo er

    Dim Qry As String, ShowDateRange As Boolean, i As Integer

    If tmrAutoLoad.Enabled = False Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    tmrAutoLoad.Enabled = False
    
    dFrom.Value = Date
    cmdNext.Tag = "1"
    'identify total columns of trancsaction and Access ID to identify user accessibility
    If mySrc = "Item Entry" Then
        TCol = 8: AccID = 19
        dFrom.Visible = False
        dTo.Visible = False
        cmdNext.Visible = False
        cmdPrev.Visible = False
        cmdRefresh.Visible = False
        Label1.Visible = False
        ShowDateRange = False
        Call ShowCommands(1)
    ElseIf mySrc = "Supplier Entry" Then
        TCol = 7: AccID = 19
        dFrom.Visible = False
        dTo.Visible = False
        cmdNext.Visible = False
        cmdPrev.Visible = False
        cmdRefresh.Visible = False
        Label1.Visible = False
        ShowDateRange = False
        Call ShowCommands(1)
    ElseIf mySrc = "Advances for Liquidation" Then
        TCol = 9: ColPrn = 0: AccID = 65
        dFrom = Date
        dTo.Value = Date
        cmdNext.Tag = "1"
    
    ElseIf mySrc = "Branch Entry" Then
        TCol = 8: AccID = 19
        dFrom.Visible = False
        dTo.Visible = False
        cmdNext.Visible = False
        cmdPrev.Visible = False
        cmdRefresh.Visible = False
        Label1.Visible = False
        ShowDateRange = False
        Call ShowCommands(1)
    ElseIf mySrc = "Sales Return" Then
        TCol = 9: ColPrn = 0: AccID = 29
        Call cmdNext_Click
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        dTo.Value = Date
        ShowDateRange = True
        Call ShowCommands(1)
    ElseIf mySrc = "Purchase Return" Then
        TCol = 9: ColPrn = 0: AccID = 24
        Call cmdNext_Click
        If Day(Date) <= 7 Then
            dFrom.Value = DateAdd("m", -1, DateSerial(Year(Date), Month(Date), 1))
        Else
            dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        End If
    ElseIf mySrc = "Inventory Adjustment" Then
        TCol = 9: ColPrn = 9: AccID = 55
        Call cmdNext_Click
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        dTo.Value = Date
        ShowDateRange = True
        Call ShowCommands(1)
'***************** CODE UPDATED: 02/12/2019 *******************
    ElseIf mySrc = "WS Inventory Adjustment" Then
        TCol = 9: ColPrn = 9: AccID = 55
        Call cmdNext_Click
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        dTo.Value = Date
        ShowDateRange = True
        Call ShowCommands(1)
'************* END CODE UPDATED: 02/12/2019 *******************
    ElseIf mySrc = "SP Incoming" Then
        TCol = 9: ColPrn = 8: AccID = 52
        Call cmdNext_Click
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        dTo.Value = Date
        ShowDateRange = True
        cmdChkClosed.Visible = True
        Call ShowCommands(1)
    ElseIf mySrc = "Branch Order" Then
        TCol = 8: ColPrn = 7: AccID = 53
        Call cmdNext_Click
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        dTo.Value = Date
        ShowDateRange = True
        Call ShowCommands(1)
    ElseIf mySrc = "Inventory" Then
        TCol = 9: AccID = 19
        dFrom.Visible = False
        dTo.Visible = False
        cmdNext.Visible = False
        cmdPrev.Visible = False
        cmdRefresh.Visible = False
        Label1.Visible = False
        ShowDateRange = False
        Call ShowCommands(1)
    ElseIf mySrc = "Detailed Receiving" Then
        TCol = 10: ColPrn = 0
        Call cmdNext_Click
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        dTo.Value = Date
        ShowDateRange = True
        Call ShowCommands(2)
    ElseIf mySrc = "Detailed Outgoing" Then
        TCol = 10: ColPrn = 0: AccID = 108
        Call cmdNext_Click
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        dTo.Value = Date
        ShowDateRange = True
        Call ShowCommands(2)
    ElseIf mySrc = "Daily Cash Collection" Then
        TCol = 14: ColPrn = 0: AccID = 51
        dFrom = Date
        dTo.Value = Date
        cmdNext.Tag = "1"
        ShowDateRange = False
        Call ShowCommands(5)
    ElseIf mySrc = "Branch Purchase Order" Then
        TCol = 14: ColPrn = 0
        'AccID =0
        bCol = 4
        'dFrom = Date
        Call cmdNext_Click
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        dTo.Value = Date
        ShowDateRange = True
        cmdUpdate.Visible = True
    ElseIf mySrc = "SP PACKING" Then
        TCol = 14: ColPrn = 0
        'dFrom = Date
        Call cmdNext_Click
'        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
'        dTo.Value = Date
         
        If Day(Date) <= 7 Then
            dFrom.Value = DateAdd("m", -1, DateSerial(Year(Date), Month(Date), 1))
        Else
            dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        End If
        dTo.Value = Date
        ShowDateRange = True
        SSConfirmTab.Visible = False
        
    ElseIf mySrc = "AllocBBO" Then
        TCol = 7: ColPrn = 0
        'dFrom = Date
        Call cmdNext_Click
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        dTo.Value = Date
        ShowDateRange = True
    ElseIf mySrc = "Costing" Then
        TCol = 12: ColPrn = 0
        Call cmdNext_Click
        dTo.Value = Date
        dFrom.Value = DateAdd("D", -30, Date)   ' DateSerial(Year(Date), Month(Date), 1)
        ShowDateRange = True
    
    ElseIf mySrc = "SRP Pricing" Then
        TCol = 8: ColPrn = 0
        Call cmdNext_Click
        dTo.Value = Date
        dFrom.Value = DateAdd("D", -30, Date)    ' DateSerial(Year(Date), Month(Date), 1)
        ShowDateRange = True

'************************CODE UPDATED BY IAN: 03/05/2018 **********************************
    ElseIf mySrc = "SRP Request form" Then
        TCol = 8: ColPrn = 0
        Call cmdNext_Click
        dTo.Value = Date
        dFrom.Value = DateAdd("D", -30, Date)    ' DateSerial(Year(Date), Month(Date), 1)
        ShowDateRange = True
'************************END CODE BY IAN: 03/05/2018 **********************************
    ElseIf mySrc = "Closed Account" Then
        TCol = 8: ColPrn = 0: AccID = 95
        For i = 1 To 12
            cboMonth.AddItem MonthName(i)
        Next i
        UpdYear.Value = Year(Date)
        txtYear.Text = Year(Date)
        If Day(Date) < 7 Then          'Date Starts with 7, Ends with 6 of the ff. month
            dFrom.Value = DateSerial(Year(DateAdd("M", -1, Date)), Month(DateAdd("M", -1, Date)), 7)
            dTo.Value = DateSerial(Year(Date), Month(Date), 6)
        Else
            dFrom.Value = DateSerial(Year(Date), Month(Date), 6)
            dTo.Value = DateSerial(Year(DateAdd("M", 1, Date)), Month(DateAdd("M", 1, Date)), 7)
        End If
        cboMonth.ListIndex = Month(dFrom.Value) - 1
        ShowDateRange = True
        Call ShowCommands(1)
        cmdShow.Enabled = False
        cmdChkClosed.Visible = True
'*********************** CODE UPDATED: 08/07/2017 **********************
    ElseIf mySrc = "SETTLEMENT DISCREPANCY" Then
        TCol = 16: ColPrn = 0
        'dFrom = Date
        Call cmdNext_Click
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        dTo.Value = Date
        ShowDateRange = True
    ElseIf mySrc = "SDiscrepancy Adjustment" Then
        TCol = 8: ColPrn = 8: AccID = 55
        Call cmdNext_Click
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        dTo.Value = Date
        ShowDateRange = True
        Call ShowCommands(1)
'******************* END CODE UPDATED: 08/07/2017 **********************
'*********************** CODE UPDATED: 06/14/2019 **********************
    ElseIf mySrc = "PO ADJUSTMENT" Then
        TCol = 8: ColPrn = 8
        Call cmdNext_Click
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        dTo.Value = Date
        ShowDateRange = True
        Call ShowCommands(1)
    ElseIf mySrc = "RFP" Then
        TCol = 15: ColPrn = 15
        Call cmdNext_Click
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        dTo.Value = Date
        ShowDateRange = True
        Call ShowCommands(1)
    ElseIf mySrc = "RFP RECONCILIATION" Then
        TCol = 14: ColPrn = 14
        Call cmdNext_Click
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        dTo.Value = Date
        ShowDateRange = True
        Call ShowCommands(1)
'******************* END CODE UPDATED: 06/14/2019 **********************
    ElseIf mySrc = "ADJUSTMENT REQUEST" Then
        TCol = 12: ColPrn = 12
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        dTo = Date
        ShowDateRange = True
    ElseIf mySrc = "CANCELLATION MODULE" Then
        TCol = 12: ColPrn = 12
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        dTo = Date
        ShowDateRange = True
    End If
    If FrDate <> DateValue("1/1/1977") Then
        dFrom.Value = FrDate
        dTo.Value = ToDate
    Else
        FrDate = dFrom.Value
        ToDate = dTo.Value
    End If
    If ShowDateRange Then
        Call RefreshRecord(2)
    Else
        Call RefreshRecord(1)
    End If
'    Marketing.sBar.Panels(6).Text = MySrc & " Summary"
    lblHeading(2).Caption = mySrc & " Summary"
    lblHeading(3).Caption = mySrc & " Summary"
    Me.Caption = mySrc & " Summary"
    'Qry = "SELECT settings.VALUE1 From settings Where settings.DESCRIPTION = 'DEPARTMENT';"
    'lblHeading(0).Caption = GetFieldValue(Qry)
    'lblHeading(1).Caption = lblHeading(0).Caption
 '   Marketing.tbMain.Visible = False
    ActRow = 1
    
    
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    
    Screen.MousePointer = Default
    
Exit Sub
er:
    tmrAutoLoad.Enabled = False
    Screen.MousePointer = vbDefault
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    Call modObjTools.ErrorMessage(Err.Description, "Error->tmrAutoLoad_Timer()")
End Sub


Private Sub tmrRefresh_Timer()
On Error GoTo er
    tmrRefresh.Enabled = False
    Screen.MousePointer = vbHourglass
    Warehouse.Enabled = False
    DoEvents
    DoEvents
    Call RefreshRecord(2)
    DoEvents
    DoEvents
    
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    
    DoEvents
    DoEvents
    Warehouse.Enabled = True
    Screen.MousePointer = vbDefault
    
Exit Sub
er:
    dFrom.Enabled = True
    'frmMain.Enabled = True
    Screen.MousePointer = vbDefault
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    Call modObjTools.ErrorMessage(Err.Description, "Error->tmrRefresh_Timer()")
End Sub

Private Sub txtFind_Change()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub txtFind2_Change()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFind2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind2_Click
End Sub


Private Sub updYear_Change()
    txtYear.Text = UpdYear.Value
    Call cboMonth_Click
End Sub

Private Sub UpdateDepositedCI()
    Dim Qry As String, CIname As String, i As Integer, CIid As Long, OrdId As Long
    
    pBar.Value = 0
    pBar.Max = flxSum.Rows
    pBar.Visible = True
    With flxSum
        For i = 1 To .Rows - 1
            Qry = "SELECT CONCAT(employees.LASTNAME, ' ', MID(employees.FIRSTNAME, 1, 1),'.') AS CI From " & _
                  "unitordered INNER JOIN employees ON (unitordered.CI = employees.EMPID) Where unitordered.ACCOUNTNO = '" & .TextMatrix(i, 3) & "';"
            CIname = GetFieldValue(Qry)
            If CIname <> "" Then
                LYRcon.Execute "UPDATE depositedunit SET PLATENO = '" & CIname & "' Where depositedunit.DEPOSITID = " & Val(.TextMatrix(i, 9)) & ";"
            Else
                Qry = "SELECT unitordered.ORDERID From unitordered Where unitordered.ACCOUNTNO = '" & .TextMatrix(i, 3) & "';"
                OrdId = Val(GetFieldValue(Qry))
                Qry = "SELECT reference.ADDRESS From reference Where reference.RNAME = 'C.I.' AND reference.CONTACTNOS = " & OrdId & ";"
                CIname = GetFieldValue(Qry)
                If CIname <> "" Then
                    LYRcon.Execute "UPDATE depositedunit SET PLATENO = '" & CIname & "' Where depositedunit.DEPOSITID = " & Val(.TextMatrix(i, 9)) & ";"
                End If
            End If
            pBar.Value = pBar.Value + 1
            DoEvents
        Next i
        pBar.Visible = False
        MsgBox "C.I. successfully updated in Deposit Agreement!", vbInformation + vbOKOnly, "Summary"
    End With
                
End Sub

Private Sub FixCommissionTax()
    Dim ZeroTax As String, NoTax As String, msgDisc As String, Qry As String, ValMsg As String
    Dim Gross As Currency, MYTax As Currency, i As Integer
    
    
    With flxSum
        msgDisc = ""
        NoTax = ""
        ZeroTax = ""
        pBar.Visible = True
        pBar.Value = 0
        pBar.Max = .Rows
        For i = 1 To .Rows - 1
            If Val(Replace(.TextMatrix(i, 6), ",", "")) > 0 Then 'deducted with tax
                Gross = (Val(Replace(.TextMatrix(i, 5), ",", "")) * 0.1)
                MYTax = Val(Replace(.TextMatrix(i, 6), ",", ""))
                If Gross <> MYTax Then
                    Qry = "SELECT disbursement.VOUCHERID From disbursement Where disbursement.CVNUMBER = '" & .TextMatrix(i, 3) & "' AND disbursement.DETAILS LIKE '%TAX%';"
                    If Val(GetFieldValue(Qry)) > 1 Then
                        Gross = Val(Replace(.TextMatrix(i, 7), ",", "")) / 0.9
                        MYTax = Gross * 0.1
                        If Gross - MYTax = Val(Replace(.TextMatrix(i, 7), ",", "")) Then
                            LYRcon.Execute "UPDATE disbursement SET TAXAMT = " & MYTax & " Where disbursement.CVNUMBER = '" & .TextMatrix(i, 3) & "' AND disbursement.DETAILS LIKE '%TAX%';"
                            .TextMatrix(i, 5) = Format(Gross, "#,##0.00")
                            .TextMatrix(i, 6) = Format(MYTax, "#,##0.00")
                            
                        Else
                            msgDisc = msgDisc & IIf(msgDisc <> "", ", ", "") & .TextMatrix(i, 3)
                        End If
                    Else
                        NoTax = NoTax & IIf(NoTax <> "", ", ", "") & .TextMatrix(i, 3)
                    End If
                End If
            Else
                ZeroTax = ZeroTax & IIf(ZeroTax <> "", ", ", "") & .TextMatrix(i, 3)
            End If
            pBar.Value = pBar.Value + 1
            DoEvents
        Next i
    End With
    pBar.Visible = False
    ValMsg = ""
    ValMsg = IIf(msgDisc <> "", "Wrong Tax Computation: " & msgDisc & vbCrLf, "")
    ValMsg = ValMsg & IIf(NoTax <> "", vbCrLf & "No Tax Deducted: " & NoTax, "")
    ValMsg = ValMsg & IIf(ZeroTax <> "", vbCrLf & "No Tax Recorded: " & ZeroTax, "")
    
    If ValMsg <> "" Then MsgBox ValMsg, vbExclamation + vbOKOnly, "Warning"
        
    
End Sub

Private Sub FixAgentSales()
    Dim rsSales As New ADODB.Recordset
    Dim Qry As String, MyAcct As String
    
    Qry = "SELECT salesinvoice.ORDERID, salesinvoice.INVDATE, unitordered.ACCOUNTNO, salesinvoice.REMARKS, SalesInvoice.PRNCOPY " & _
    "From unitordered INNER JOIN salesinvoice ON (unitordered.ORDERID = salesinvoice.ORDERID) " & _
    "Where salesinvoice.INVDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' AND salesinvoice.INVDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' Order By SalesInvoice.InvDate"
    
    pBar.Value = 0
    pBar.Visible = True
    With rsSales
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            pBar.Max = .RecordCount
            .MoveFirst
            Do While Not .EOF
                MyAcct = IIf(IsNull(!ACCOUNTNO), "", !ACCOUNTNO)
                If MyAcct <> "" Then
                    LYRcon.Execute "UPDATE agentorders SET SALESDATE = '" & Format(!InvDate, "YYYY-MM-DD") & "' Where agentorders.OrderID = " & !OrderID & ";"
                Else
                    LYRcon.Execute "UPDATE agentorders SET SALESDATE = '" & Format(!InvDate, "YYYY-MM-DD") & "', `STATUS` = 'Pending' Where agentorders.OrderID = " & !OrderID & ";"
                End If
                .MoveNext
                pBar.Value = pBar.Value + 1
                DoEvents
            Loop
        End If
        .Close
    End With
    
    pBar.Visible = False
    MsgBox "Finish checking Sales!", vbInformation + vbOKOnly, "Sales Invoice"
    
    
End Sub


Private Sub ReviewClosedAcct()
    Dim Qry As String, CutOff As String, DateClosed As String
    Dim MyID As Long, i As Integer, j As Integer, Review As Boolean
    
    If flxSum.Rows = 2 And flxSum.TextMatrix(1, 1) = "" Then Exit Sub
    
    Qry = "SELECT CONCAT(AGEMONTH,'/6/',AGEYEAR) AS CUTOFF From agingcurrent;"
    CutOff = GetFieldValue(Qry)
    If Not IsDate(CutOff) Then
        MsgBox "Aging not yet available! No Cut-Off date identified!", vbExclamation + vbOKOnly, "Closed Accounts"
        Exit Sub
    End If
    pBar.Value = 0
    pBar.Max = flxSum.Rows
    pBar.Visible = True
    With flxSum
        For i = 1 To .Rows - 1
            If Not IsDate(.TextMatrix(i, 5)) Then
                Review = True
            Else
                If DateValue(.TextMatrix(i, 5)) <= DateValue(CutOff) Then
                    Review = True
                Else
                    Review = False
                End If
            End If
            If Review Then
                Qry = "SELECT unitordered.DATECLOSED From unitordered Where unitordered.ACCOUNTNO = '" & .TextMatrix(i, 1) & "';"
                DateClosed = GetFieldValue(Qry)
                If IsDate(DateClosed) Then
                    Qry = "SELECT unitordered.ORDERID From unitordered Where unitordered.ACCOUNTNO = '" & .TextMatrix(i, 1) & "';"
                    MyID = Val(GetFieldValue(Qry))
                    LYRcon.Execute "DELETE From agingcurrent Where agingcurrent.OrderID = " & MyID & ";"
                    LYRcon.Execute "UPDATE unitordered SET ACCTSTATUS = 'CLOSED' Where unitordered.ACCOUNTNO = '" & .TextMatrix(i, 1) & "';"
                Else
                    txtMo.Text = Month(DateValue(CutOff))
                    txtYr.Text = Year(DateValue(CutOff))
                    txtCutOff.Text = DateAdd("m", 1, DateValue(CutOff))
                    Qry = "SELECT UNITORDERED.ACCOUNTNO, UNITORDERED.ORDERID, UNITORDERED.MONTHLYINST, UNITORDERED.FIRSTDUEDATE, UNITORDERED.TERM, UNITORDERED.TERMTYPE, UNITORDERED.DATEPURCHASED, UNITORDERED.ACCTSTATUS, UNITORDERED.INVOICENO, UNITORDERED.INVAMOUNT, UNITORDERED.COLLECTORID, UNITORDERED.AREA, UNITORDERED.TOTALINSTPRICE, UNITORDERED.REGDP, UNITORDERED.MIREBATE, UNITORDERED.SRP, UNITORDERED.LCP, UNITORDERED.ITEMID, UNITORDERED.MODEL, UNITORDERED.SERIALNO " & _
                          "From UNITORDERED WHERE (UNITORDERED.ACCOUNTNO='" & .TextMatrix(i, 1) & "');"
                    If rsAccts.State = adStateOpen Then rsAccts.Close
                    rsAccts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                    Set flxAccts.DataSource = rsAccts
                    flxAccts.Refresh
                    If flxAccts.TextMatrix(1, 6) = "D" Then
                        Call DispDailyLedger
                    Else
                        Call DisplayLedger
                    End If
                    For j = flxLedger.Rows - 1 To 0 Step -1
                        If IsDate(flxLedger.TextMatrix(j, 5)) Then
                            DateClosed = flxLedger.TextMatrix(j, 5)
                            Exit For
                        End If
                    Next j
                    If IsDate(DateClosed) Then
                        If DateValue(DateClosed) <= DateValue(CutOff) Then
                            Qry = "SELECT unitordered.ORDERID From unitordered Where unitordered.ACCOUNTNO = '" & .TextMatrix(i, 1) & "';"
                            MyID = Val(GetFieldValue(Qry))
                            LYRcon.Execute "DELETE From agingcurrent Where agingcurrent.OrderID = " & MyID & ";"
                        End If
                        LYRcon.Execute "UPDATE unitordered SET ACCTSTATUS = 'CLOSED', unitordered.DATECLOSED = '" & Format$(DateClosed, "yyyy-mm-dd") & "', unitordered.CLOSINGREMARKS = 'CLOSED ACCT.'  Where unitordered.ACCOUNTNO = '" & .TextMatrix(i, 1) & "';"
                    End If
                End If
            End If
            pBar.Value = pBar.Value + 1
            DoEvents
        Next i
        Call RefreshRecord(2)
        pBar.Visible = False
        MsgBox "Closed Account successfully checked!", vbInformation + vbOKOnly, "Summary"
    End With
End Sub

Private Sub DisplayLedger()
    Dim rsPay As New ADODB.Recordset
    Dim rsAdj As New ADODB.Recordset
    Dim Qry As String, ORType As String, sAdvArr As String, MyAcct As String
    Dim ORAmt As Currency, DP As Currency, TtlCredits As Currency, Balance As Currency, cArrAdv As Currency, InsMos As Currency, TtlAmtDue As Currency, Rebate As Currency
    Dim DateLastBill As Date, CDueDate As Date, MonDue As Date, tmpDate As Date, FinalDateBill As Date
    Dim i As Integer, cRow As Integer, BillNo As Integer, DispBillOnly As Boolean, ExpAcct As Boolean, XpRdm As Boolean
    Dim TtlDP As Currency, RowLastDP As Integer, j As Integer
    
    MyAcct = flxAccts.TextMatrix(1, 1)
    Qry = "SELECT PAYMENTS.PAYDATE, PAYMENTS.ORNO, PAYMENTS.DOWNPAYMENT, PAYMENTS.INSTALLMENT, PAYMENTS.DISCOUNT, PAYMENTS.ACCOUNTNO, PAYMENTS.CASH, PAYMENTS.CRNO " & _
          "From PAYMENTS Where (PAYMENTS.ACCOUNTNO = '" & MyAcct & "') ORDER BY PAYMENTS.PAYDATE, PAYMENTS.ORNO;"
    flxPay.Rows = 1
    XpRdm = False
    With rsPay
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxPay.Rows = 1
            Do While Not .EOF
                DP = IIf(IsNull(!DownPayment), 0, !DownPayment)
                'TtlDP = TtlDP + DP
                TtlDP = TtlDP + DP + !discount + IIf(!Installment > 0 And TtlDP < Val(flxAccts.TextMatrix(1, 14)), !Installment, 0)
                If TtlDP < Val(flxAccts.TextMatrix(1, 14)) And DP = 0 And !Installment > 0 Then
                    ORAmt = IIf(IsNull(!Installment), 0, !Installment)
                    ORType = "DP"
                    TtlDP = TtlDP + ORAmt
                ElseIf DP > 0 And !Installment > 0 Then
                    ORAmt = DP + !Installment
                    ORType = "DP"
                ElseIf DP > 0 Then
                    ORAmt = DP
                    ORType = "DP"
                ElseIf !Installment > 0 Then
                    ORAmt = IIf(IsNull(!Installment), 0, !Installment)
                    ORType = "INST"
                ElseIf !Cash > 0 Then
                    ORAmt = IIf(IsNull(!Cash), 0, !Cash)
                    ORType = "C.O.D."
                Else
                    ORAmt = 0
                    ORType = ""
                End If
                If ORAmt > 0 And ORType <> "" Then
                    flxPay.AddItem Format(!PAYDATE, "mm/dd/yy") & vbTab & IIf(!ORNO <> "", !ORNO, !CRNO) & vbTab & ORAmt & vbTab & IIf(IsNull(!discount), "", !discount) & vbTab & ORType
                End If
                .MoveNext
            Loop
        End If
        .Close
    End With
    Qry = "SELECT LEDGERADJ.REFDATE, LEDGERADJ.REFNO, LEDGERADJ.REFAMT, LEDGERADJ.TRANSTYPE From LEDGERADJ " & _
          "WHERE (LEDGERADJ.ACCOUNTNO='" & MyAcct & "' AND LEDGERADJ.TRANSTYPE<>'temp') ORDER BY LEDGERADJ.REFDATE;"
    With rsAdj
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            RowLastDP = 0
            Do While Not .EOF
                For i = 1 To flxPay.Rows - 1
                    If !TRANSTYPE = "REDEEM" Then
                        flxPay.AddItem Format(!RefDate, "mm/dd/yyyy") & vbTab & !RefNo & vbTab & !RefAmt & vbTab & vbTab & !TRANSTYPE, RowLastDP    'IIf(RowLastDP > flxPay.Rows - 1, flxPay.Rows - 1, RowLastDP)
                        Exit For
                    Else
                        If DateValue(flxPay.TextMatrix(i, 0)) > !RefDate Then
                            flxPay.AddItem Format(!RefDate, "mm/dd/yyyy") & vbTab & !RefNo & vbTab & !RefAmt & vbTab & vbTab & !TRANSTYPE, i
                            Exit For
                        End If
                    End If
                Next i
                If !TRANSTYPE = "DEPOSIT" Then RowLastDP = i + 1     'NEW ENTRY
                If i = flxPay.Rows Then flxPay.AddItem Format(!RefDate, "mm/dd/yyyy") & vbTab & !RefNo & vbTab & !RefAmt & vbTab & vbTab & !TRANSTYPE, i
                .MoveNext
            Loop
        End If
        .Close
    End With
    If flxPay.Rows = 1 Then
        flxLedger.Rows = 0
        If flxAccts.TextMatrix(1, 6) = "M" Then
            If Val(flxAccts.TextMatrix(1, 14)) = 0 Then
                CDueDate = DateValue(flxAccts.TextMatrix(1, 4))
                flxLedger.AddItem "1" & vbTab & Format(CDueDate, "mm/dd/yyyy") & vbTab & Format(Val(flxAccts.TextMatrix(1, 3)), "#,##0.00") & vbTab & "MI"
                flxLedger.TextMatrix(flxLedger.Rows - 1, 9) = "0"
                flxLedger.TextMatrix(flxLedger.Rows - 1, 10) = Format(Val(flxAccts.TextMatrix(1, 13)), "#,##0.00")
                flxLedger.TextMatrix(flxLedger.Rows - 1, 11) = Val(flxAccts.TextMatrix(1, 3))
            Else
                CDueDate = DateValue(flxAccts.TextMatrix(1, 7))
                flxLedger.AddItem vbTab & Format(CDueDate, "mm/dd/yyyy") & vbTab & Format(Val(flxAccts.TextMatrix(1, 14)), "#,##0.00") & vbTab & "DP"
                flxLedger.TextMatrix(flxLedger.Rows - 1, 9) = "0"
                flxLedger.TextMatrix(flxLedger.Rows - 1, 10) = Format(Val(flxAccts.TextMatrix(1, 13)), "#,##0.00")
                flxLedger.TextMatrix(flxLedger.Rows - 1, 11) = Val(flxAccts.TextMatrix(1, 14))
            End If
        ElseIf flxAccts.TextMatrix(1, 6) = "C" Then
            CDueDate = DateValue(flxAccts.TextMatrix(1, 7))
            Balance = GetCODAmt(MyAcct)
            flxLedger.AddItem vbTab & Format(CDueDate, "mm/dd/yyyy") & vbTab & Format(Balance, "#,##0.00") & vbTab & "COD"
            flxLedger.TextMatrix(flxLedger.Rows - 1, 9) = "0"
            flxLedger.TextMatrix(flxLedger.Rows - 1, 10) = Format(Balance, "#,##0.00")
            flxLedger.TextMatrix(flxLedger.Rows - 1, 11) = Balance
        End If
        Exit Sub
    End If
    If flxAccts.TextMatrix(1, 6) = "C" Then 'COD post payment
        flxLedger.Rows = 0
        TtlCredits = 0
        For i = 1 To flxPay.Rows - 1
            If i = 1 Then
                Balance = Val(Replace(flxAccts.TextMatrix(1, 10), ",", "")) - Val(flxPay.TextMatrix(i, 2))
            Else
                Balance = Val(Replace(flxLedger.TextMatrix(i - 2, 10), ",", "")) - Val(flxPay.TextMatrix(i, 2))
            End If
            TtlCredits = TtlCredits + Val(flxPay.TextMatrix(i, 2))
            flxLedger.AddItem vbTab & vbTab & vbTab & "C.O.D." & vbTab & vbTab & flxPay.TextMatrix(i, 0) & vbTab & flxPay.TextMatrix(i, 1) & vbTab & Format(flxPay.TextMatrix(i, 2), "#,##0.00") & vbTab & vbTab & Format(TtlCredits, "#,##0.00") & vbTab & Format(Balance, "#,##0.00")
        Next i
        Exit Sub
    End If
    If flxAccts.TextMatrix(1, 6) = "M" Then
        DateLastBill = DateAdd("M", Val(flxAccts.TextMatrix(1, 5)), flxAccts.TextMatrix(1, 4))
        FinalDateBill = DateLastBill
    ElseIf flxAccts.TextMatrix(1, 6) = "D" Then
        DateLastBill = DateAdd("D", Val(flxAccts.TextMatrix(1, 5)), flxAccts.TextMatrix(1, 4))
    End If
    ExpAcct = False
    If DateDiff("m", DateValue(flxAccts.TextMatrix(1, 4)), DateValue(txtCutOff.Text)) > Val(flxAccts.TextMatrix(1, 5)) Then ExpAcct = True
    If Val(flxAccts.TextMatrix(1, 5)) > 1 Then   'Term
        CDueDate = DateValue(flxAccts.TextMatrix(1, 4))  '1st due date
        MonDue = DateSerial(Val(txtYr.Text), Val(txtMo.Text), 6)  'IIf(Day(CDueDate) <= Day(Date), Day(CDueDate), Day(Date)))
        If DateLastBill > DateValue(txtCutOff.Text) Then DateLastBill = DateValue(txtCutOff.Text)
        If DateLastBill < CDueDate Then DateLastBill = CDueDate
        cRow = 1: BillNo = 1: DP = Val(flxAccts.TextMatrix(1, 14)): cArrAdv = 0
        flxLedger.Rows = 0
        'flxLedger.ColWidth(11) = flxLedger.Width * 0.08
        TtlCredits = 0: cArrAdv = 0: Balance = Val(flxAccts.TextMatrix(1, 13))
        Do While CDueDate <= DateLastBill Or (BillNo <= Val(flxAccts.TextMatrix(1, 5)) And CDueDate > FinalDateBill And XpRdm) Or ExpAcct
            If CDueDate = DateValue(flxAccts.TextMatrix(1, 4)) Then  'Current Due date = 1st due date:    DP and Adv 1st Monthly
                Do While DateValue(flxPay.TextMatrix(cRow, 0)) < CDueDate 'DateSerial(Year(CDueDate), Month(CDueDate), 6)
                    ORAmt = Val(flxPay.TextMatrix(cRow, 2))
                    Rebate = Val(flxPay.TextMatrix(cRow, 3))
                    If flxPay.TextMatrix(cRow, 4) = "DP" Then
                        InsMos = DP 'IIf(DP > 0 And flxLedger.Rows = 1, DP, DP - Val(Replace(flxLedger.TextMatrix(flxLedger.Rows - 1, 4), ",", "")) - (Val(Replace(flxLedger.TextMatrix(flxLedger.Rows - 1, 7), ",", "")) + Val(Replace(flxLedger.TextMatrix(flxLedger.Rows - 1, 8), ",", ""))))
                        TtlAmtDue = InsMos '+ IIf(InStr(flxLedger.TextMatrix(flxLedger.Rows - 1, 3), "DP") > 0, 0, CArrAdv)
                        If DP > 0 Then
                            sAdvArr = IIf(TtlAmtDue - (ORAmt + Rebate) > 0, "Prtl DP", "Full DP")
                        Else
                            sAdvArr = ""
                            TtlAmtDue = cArrAdv 'new 6/18/2011
                        End If
                        DP = IIf(TtlAmtDue - (ORAmt + Rebate) > 0, TtlAmtDue - (ORAmt + Rebate), 0)
                    Else
                        InsMos = 0
                        TtlAmtDue = cArrAdv 'IIf(flxLedger.Rows = 1, 0, Val(Replace(flxLedger.TextMatrix(flxLedger.Rows - 1, 4), ",", "")) - (Val(Replace(flxLedger.TextMatrix(flxLedger.Rows - 1, 7), ",", "")) + Val(Replace(flxLedger.TextMatrix(flxLedger.Rows - 1, 8), ",", ""))))
                        If flxPay.TextMatrix(cRow, 4) = "INST" Then
                            If DP > 0 Then
                                InsMos = DP 'IIf(DP > 0 And flxLedger.Rows = 1, DP, DP - Val(Replace(flxLedger.TextMatrix(flxLedger.Rows - 1, 4), ",", "")) - (Val(Replace(flxLedger.TextMatrix(flxLedger.Rows - 1, 7), ",", "")) + Val(Replace(flxLedger.TextMatrix(flxLedger.Rows - 1, 8), ",", ""))))
                                TtlAmtDue = InsMos + cArrAdv '+ IIf(InStr(flxLedger.TextMatrix(flxLedger.Rows - 1, 3), "DP") > 0, 0, CArrAdv)
                                sAdvArr = IIf(TtlAmtDue - (ORAmt + Rebate) > 0, "Prtl DP", "Full DP")
                                DP = IIf(DP > 0, DP - (ORAmt + Rebate), 0)
                            Else
                                sAdvArr = IIf(sAdvArr <> "M.I.", "M.I.", "")
                            End If
                        Else
                            sAdvArr = flxPay.TextMatrix(cRow, 4)
                            If DP > 0 Then
                                If cRow = 1 Then TtlAmtDue = DP
                                DP = DP - (ORAmt + Rebate)
                            Else
                                DP = 0
                            End If
                        End If
                    End If
                    TtlCredits = TtlCredits + ORAmt + Rebate
                    Balance = Balance - (ORAmt + Rebate)
                    cArrAdv = TtlAmtDue - (ORAmt + Rebate)
                    flxLedger.AddItem vbTab & IIf(flxPay.TextMatrix(cRow, 4) = "DP", flxPay.TextMatrix(cRow, 0), "") & vbTab & IIf(InsMos > 0, Format(InsMos, "#,##0.00"), "") & vbTab & sAdvArr & vbTab & IIf(TtlAmtDue > 0, IIf(sAdvArr = "DMCM", "", Format(TtlAmtDue, "#,##0.00")), "") & vbTab & flxPay.TextMatrix(cRow, 0) & vbTab & flxPay.TextMatrix(cRow, 1) & vbTab & Format(ORAmt, "#,##0.00") & vbTab & IIf(Rebate <> 0, Format(Rebate, "#,##0.00"), "") & vbTab & Format(TtlCredits, "#,##0.00") & vbTab & Format(Balance, "#,##0.00") & vbTab & cArrAdv
                    cRow = cRow + 1
                    If cRow > flxPay.Rows - 1 Then Exit Do
                Loop
            End If
            If cRow > flxPay.Rows - 1 Then
                DispBillOnly = True
                ExpAcct = False
            Else
                DispBillOnly = False
            End If
            If BillNo = 1 And TtlCredits = 0 Then
                InsMos = Val(flxAccts.TextMatrix(1, 3)) + DP 'to be continue... Add DP kng BillNo=1 and No Payments Made
            Else
                InsMos = Val(flxAccts.TextMatrix(1, 3))  'to be continue... Add DP kng BillNo=1 and No Payments Made
            End If
            If flxAccts.TextMatrix(1, 6) = "D" Then
                MonDue = DateAdd("m", 1, CDueDate)
                InsMos = InsMos * Day(DateAdd("d", -1, DateSerial(Year(MonDue), Month(MonDue), 1)))
            End If
            If Balance = 0 And cRow >= flxPay.Rows - 1 Then Exit Do
            sAdvArr = IIf(cArrAdv <> 0, FormatNumber(cArrAdv, 2, vbTrue, vbTrue), "")
            If BillNo <= Val(flxAccts.TextMatrix(1, 5)) And CDueDate <= DateValue(txtCutOff.Text) Then      'New Code Added(4/21/2009): And CDueDate <= DateValue(txtCutOff.Text)   /* Para Dili mo bill kng wla pay cutoff
            'If DateValue(flxAccts.TextMatrix(1, 4)) <= DateLastBill Then 'DateValue(txtCutOff.Text) Then
                TtlAmtDue = InsMos + cArrAdv
                cArrAdv = TtlAmtDue
                flxLedger.AddItem BillNo & vbTab & Format(CDueDate, "MM/DD/YY") & vbTab & Format(InsMos, "#,##0.00") & vbTab & sAdvArr & vbTab & FormatNumber(TtlAmtDue, 2, vbTrue, vbTrue) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & cArrAdv
            End If
            If flxLedger.Rows > 10 Then flxLedger.TopRow = flxLedger.Rows - 1
            If Not DispBillOnly Then
                Do While DateValue(flxPay.TextMatrix(cRow, 0)) <= DateAdd("m", 1, DateSerial(Year(CDueDate), Month(CDueDate), 6)) Or BillNo > Val(flxAccts.TextMatrix(1, 5)) 'DateAdd("m", 1, DateSerial(Year(CDueDate), Month(CDueDate), 6))
                    ORAmt = Val(flxPay.TextMatrix(cRow, 2))
                    Rebate = Val(flxPay.TextMatrix(cRow, 3))
                    TtlCredits = TtlCredits + ORAmt + Rebate
                    Balance = Balance - (ORAmt + Rebate)
                    If flxPay.TextMatrix(cRow, 4) = "INST" Or flxPay.TextMatrix(cRow, 4) = "DP" Then
                        cArrAdv = TtlAmtDue - (ORAmt + Rebate)
                        If flxLedger.TextMatrix(flxLedger.Rows - 1, 5) = "" Then
                            flxLedger.TextMatrix(flxLedger.Rows - 1, 5) = flxPay.TextMatrix(cRow, 0)
                            flxLedger.TextMatrix(flxLedger.Rows - 1, 6) = flxPay.TextMatrix(cRow, 1)
                            flxLedger.TextMatrix(flxLedger.Rows - 1, 7) = Format(ORAmt, "#,##0.00")
                            flxLedger.TextMatrix(flxLedger.Rows - 1, 8) = IIf(Rebate <> 0, Format(Rebate, "#,##0.00"), "")
                            flxLedger.TextMatrix(flxLedger.Rows - 1, 9) = Format(TtlCredits, "#,##0.00")
                            flxLedger.TextMatrix(flxLedger.Rows - 1, 10) = Format(Balance, "#,##0.00")
                            flxLedger.TextMatrix(flxLedger.Rows - 1, 11) = cArrAdv
                        Else
                            flxLedger.AddItem vbTab & vbTab & vbTab & vbTab & vbTab & flxPay.TextMatrix(cRow, 0) & vbTab & flxPay.TextMatrix(cRow, 1) & vbTab & Format(ORAmt, "#,##0.00") & vbTab & IIf(Rebate <> 0, Format(Rebate, "#,##0.00"), "") & vbTab & Format(TtlCredits, "#,##0.00") & vbTab & Format(Balance, "#,##0.00") & vbTab & cArrAdv
                        End If
                        TtlAmtDue = cArrAdv
                    Else
                        If flxPay.TextMatrix(cRow, 4) = "REDEEM" Then   'NEW ENTRY
                            TtlCredits = flxPay.TextMatrix(cRow, 2)
                            Balance = Val(flxAccts.TextMatrix(1, 13)) - TtlCredits
                            flxLedger.AddItem vbTab & vbTab & vbTab & flxPay.TextMatrix(cRow, 4) & vbTab & vbTab & flxPay.TextMatrix(cRow, 0) & vbTab & flxPay.TextMatrix(cRow, 1) & vbTab & Format(ORAmt, "#,##0.00") & vbTab & IIf(Rebate <> 0, Format(Rebate, "#,##0.00"), "") & vbTab & Format(TtlCredits, "#,##0.00") & vbTab & Format(Balance, "#,##0.00") & vbTab & cArrAdv
                        Else
                            If flxPay.TextMatrix(cRow, 4) = "DEPOSIT" Then 'NEW ENTRY UPTO THIS LINE
                                If cRow + 1 <= flxPay.Rows - 1 Then
                                    If flxPay.TextMatrix(cRow + 1, 4) = "REDEEM" Then
                                        j = DateDiff("d", DateValue(flxPay.TextMatrix(cRow, 0)), DateValue(flxPay.TextMatrix(cRow + 1, 0)))
                                        XpRdm = True
                                        If j >= 15 Then
                                            If DateValue(flxPay.TextMatrix(cRow + 1, 0)) < DateSerial(Year(DateValue(flxPay.TextMatrix(cRow + 1, 0))), Month(DateValue(flxPay.TextMatrix(cRow + 1, 0))), 25) Then
                                                j = DateDiff("M", DateValue(flxPay.TextMatrix(cRow, 0)), DateValue(flxPay.TextMatrix(cRow + 1, 0)))
                                                If j > 1 Or Day(DateValue(flxPay.TextMatrix(cRow + 1, 0))) <= 6 Then
                                                    tmpDate = DateAdd("M", -1, DateValue(flxPay.TextMatrix(cRow + 1, 0)))
                                                    CDueDate = DateSerial(Year(tmpDate), Month(tmpDate), Day(DateValue(flxAccts.TextMatrix(1, 4))))
                                                Else
                                                    If Day(DateValue(flxPay.TextMatrix(cRow, 0))) <= 6 Then    'BELOW CUTOFF ANG DATE OF DEPOSIT <=6 OG Dili mo bill kng wla pay cutoff
                                                        CDueDate = DateAdd("M", 1, CDueDate)
                                                    Else
                                                        CDueDate = DateSerial(Year(DateValue(flxPay.TextMatrix(cRow + 1, 0))), Month(DateValue(flxPay.TextMatrix(cRow + 1, 0))), Day(DateValue(flxAccts.TextMatrix(1, 4))))
                                                    End If
                                                End If
                                            Else
                                                CDueDate = DateSerial(Year(DateValue(flxPay.TextMatrix(cRow + 1, 0))), Month(DateValue(flxPay.TextMatrix(cRow + 1, 0))), Day(DateValue(flxAccts.TextMatrix(1, 4))))
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                cArrAdv = cArrAdv - (ORAmt + Rebate)
                                TtlAmtDue = cArrAdv
                            End If
                            flxLedger.AddItem vbTab & vbTab & vbTab & flxPay.TextMatrix(cRow, 4) & vbTab & vbTab & flxPay.TextMatrix(cRow, 0) & vbTab & flxPay.TextMatrix(cRow, 1) & vbTab & Format(ORAmt, "#,##0.00") & vbTab & IIf(Rebate <> 0, Format(Rebate, "#,##0.00"), "") & vbTab & Format(TtlCredits, "#,##0.00") & vbTab & Format(Balance, "#,##0.00") & vbTab & cArrAdv
                        End If
                    End If
                    cRow = cRow + 1
                    If cRow > flxPay.Rows - 1 Then
                        ExpAcct = False
                        Exit Do
                    End If
                Loop
            End If
            CDueDate = DateAdd("M", 1, CDueDate)
            'If cRow > flxPay.Rows - 1 Then Exit Do
            BillNo = BillNo + 1
        Loop
    End If
End Sub

Private Sub DispDailyLedger()
    Dim rsPay As New ADODB.Recordset
    Dim rsAdj As New ADODB.Recordset
    Dim Qry As String, ORType As String, sAdvArr As String, MyAcct As String
    Dim TtlDP As Currency, ORAmt As Currency, DP As Currency, TtlCredits As Currency, Balance As Currency, cArrAdv As Currency, InsMos As Currency, TtlAmtDue As Currency, Rebate As Currency
    Dim DateLastBill As Date, CDueDate As Date, MonDue As Date, FinalDateBill As Date
    Dim DispBillOnly As Boolean, ExpAcct As Boolean, Billy As Boolean, XpRdm As Boolean
    Dim k As Integer, RowLastDP As Integer, cRow As Integer, BillNo As Integer, StartBN As Integer, myCut As Integer, PrevBN As Integer
    
    'Call ShowPleaseWait(True)
    MyAcct = flxAccts.TextMatrix(1, 1)
    Qry = "SELECT PAYMENTS.PAYDATE, PAYMENTS.ORNO, PAYMENTS.DOWNPAYMENT, PAYMENTS.INSTALLMENT, PAYMENTS.DISCOUNT, PAYMENTS.ACCOUNTNO, PAYMENTS.CRNO " & _
          "From PAYMENTS Where (((PAYMENTS.ACCOUNTNO) = '" & MyAcct & "')) ORDER BY PAYMENTS.PAYDATE;"
    flxPay.Rows = 1
    XpRdm = False
    With rsPay
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxPay.Rows = 1
            TtlDP = 0
            Do While Not .EOF
                DP = IIf(IsNull(!DownPayment), 0, !DownPayment)
                'TtlDP = TtlDP + DP
                TtlDP = TtlDP + DP + !discount + IIf(!Installment > 0 And TtlDP < Val(flxAccts.TextMatrix(1, 14)), !Installment, 0)
                If TtlDP < Val(flxAccts.TextMatrix(1, 14)) And DP = 0 And !Installment > 0 Then
                    ORAmt = IIf(IsNull(!Installment), 0, !Installment)
                    ORType = "DP"
                    TtlDP = TtlDP + ORAmt
                ElseIf DP > 0 And !Installment > 0 Then
                    ORAmt = DP + !Installment
                    ORType = "DP"
                ElseIf DP > 0 Then
                    ORAmt = DP
                    ORType = "DP"
                ElseIf !Installment > 0 Then
                    ORAmt = IIf(IsNull(!Installment), 0, !Installment)
                    ORType = "INST"
                Else
                    ORAmt = 0
                    ORType = ""
                End If
                If ORAmt > 0 And ORType <> "" Then
                    flxPay.AddItem Format(!PAYDATE, "mm/dd/yyyy") & vbTab & IIf(!ORNO <> "", !ORNO, !CRNO) & vbTab & ORAmt & vbTab & IIf(IsNull(!discount), "", !discount) & vbTab & ORType
                End If
                .MoveNext
            Loop
        End If
        .Close
    End With
    Qry = "SELECT LEDGERADJ.REFDATE, LEDGERADJ.REFNO, LEDGERADJ.REFAMT, LEDGERADJ.TRANSTYPE From LEDGERADJ " & _
          "WHERE (LEDGERADJ.ACCOUNTNO='" & MyAcct & "' AND LEDGERADJ.TRANSTYPE<>'temp') ORDER BY LEDGERADJ.REFDATE ASC;"
    With rsAdj
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            RowLastDP = 0
            Do While Not .EOF
                For i = 1 To flxPay.Rows - 1
                    If !TRANSTYPE = "REDEEM" Then
                        flxPay.AddItem Format(!RefDate, "mm/dd/yyyy") & vbTab & !RefNo & vbTab & !RefAmt & vbTab & vbTab & !TRANSTYPE, IIf(RowLastDP > flxPay.Rows - 1, flxPay.Rows - 1, RowLastDP)
                        Exit For
                    Else
                        If DateValue(flxPay.TextMatrix(i, 0)) > !RefDate Then
                            flxPay.AddItem Format(!RefDate, "mm/dd/yyyy") & vbTab & !RefNo & vbTab & !RefAmt & vbTab & vbTab & !TRANSTYPE, i
                            Exit For
                        End If
                    End If
                Next i
                If !TRANSTYPE = "DEPOSIT" Then RowLastDP = i + 1  'flxPay.Rows - 1       'NEW ENTRY
                If i = flxPay.Rows Then flxPay.AddItem Format(!RefDate, "mm/dd/yyyy") & vbTab & !RefNo & vbTab & !RefAmt & vbTab & vbTab & !TRANSTYPE, i
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    If flxPay.Rows = 1 Then
        flxLedger.Rows = 0
        If flxAccts.TextMatrix(1, 6) = "D" Then
            If Val(flxAccts.TextMatrix(1, 14)) = 0 Then
                CDueDate = DateValue(flxAccts.TextMatrix(1, 4))
                flxLedger.AddItem "1" & vbTab & Format(CDueDate, "mm/dd/yyyy") & vbTab & Format(Val(flxAccts.TextMatrix(1, 3)), "#,##0.00") & vbTab & "MI"
                flxLedger.TextMatrix(flxLedger.Rows - 1, 9) = "0"
                flxLedger.TextMatrix(flxLedger.Rows - 1, 10) = Format(Val(flxAccts.TextMatrix(1, 13)), "#,##0.00")
                flxLedger.TextMatrix(flxLedger.Rows - 1, 11) = Val(flxAccts.TextMatrix(1, 3))
            Else
                CDueDate = DateValue(flxAccts.TextMatrix(1, 7))
                flxLedger.AddItem vbTab & Format(CDueDate, "mm/dd/yyyy") & vbTab & Format(Val(flxAccts.TextMatrix(1, 14)), "#,##0.00") & vbTab & "DP"
                flxLedger.TextMatrix(flxLedger.Rows - 1, 9) = "0"
                flxLedger.TextMatrix(flxLedger.Rows - 1, 10) = Format(Val(flxAccts.TextMatrix(1, 13)), "#,##0.00")
                flxLedger.TextMatrix(flxLedger.Rows - 1, 11) = Val(flxAccts.TextMatrix(1, 14))
            End If
        ElseIf flxAccts.TextMatrix(1, 6) = "C" Then
            CDueDate = DateValue(flxAccts.TextMatrix(1, 7))
            Balance = GetCODAmt(MyAcct)
            flxLedger.AddItem vbTab & Format(CDueDate, "mm/dd/yyyy") & vbTab & Format(Balance, "#,##0.00") & vbTab & "COD"
            flxLedger.TextMatrix(flxLedger.Rows - 1, 9) = "0"
            flxLedger.TextMatrix(flxLedger.Rows - 1, 10) = Format(Balance, "#,##0.00")
            flxLedger.TextMatrix(flxLedger.Rows - 1, 11) = Balance
        End If
        If CDueDate < DateSerial(Val(txtYr.Text), Val(txtMo.Text), 7) And flxAccts.TextMatrix(1, 6) <> "C" Then
            CDueDate = DateValue(flxAccts.TextMatrix(1, 4))  '1st due date
            MonDue = DateSerial(Val(txtYr.Text), Val(txtMo.Text), 6)  'IIf(Day(CDueDate) <= Day(Date), Day(CDueDate), Day(Date)))
            DateLastBill = DateAdd("M", Val(flxAccts.TextMatrix(1, 5)), flxAccts.TextMatrix(1, 4))
            If DateLastBill > DateValue(txtCutOff.Text) Then DateLastBill = DateValue(txtCutOff.Text)
            If DateLastBill < CDueDate Then DateLastBill = CDueDate
            cRow = 1: BillNo = 1: DP = Val(flxAccts.TextMatrix(1, 14)): cArrAdv = 0
            flxLedger.Rows = 0
            'flxLedger.ColWidth(11) = flxLedger.Width * 0.08
            TtlCredits = 0: cArrAdv = 0: Balance = Val(flxAccts.TextMatrix(1, 13))
            Do While CDueDate <= DateLastBill Or ExpAcct
                InsMos = Val(flxAccts.TextMatrix(1, 3))
                If Balance = 0 Then Exit Do
                If BillNo <= Val(flxAccts.TextMatrix(1, 5)) And CDueDate <= DateValue(txtCutOff.Text) Then      'New Code Added(4/21/2009): And CDueDate <= DateValue(txtCutOff.Text)   /* Para Dili mo bill kng wla pay cutoff
                    TtlAmtDue = InsMos + cArrAdv
                    cArrAdv = TtlAmtDue
                    flxLedger.AddItem BillNo & vbTab & Format(CDueDate, "mm/dd/yyyy") & vbTab & Format(InsMos, "#,##0.00") & vbTab & sAdvArr & vbTab & FormatNumber(TtlAmtDue, 2, vbTrue, vbTrue) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & Format(Balance, "#,##0.00") & vbTab & cArrAdv
                End If
    
                CDueDate = DateAdd("M", 1, CDueDate)
                BillNo = BillNo + 1
            Loop
        End If
        Exit Sub
    End If
    DateLastBill = DateAdd("D", Val(flxAccts.TextMatrix(1, 5)), flxAccts.TextMatrix(1, 4))
    FinalDateBill = DateLastBill
    Billy = True
    If Val(flxAccts.TextMatrix(1, 5)) > 1 Then   'Term
        CDueDate = DateValue(flxAccts.TextMatrix(1, 4))  '1st due date
        If DateLastBill > DateValue(txtCutOff.Text) Then DateLastBill = DateValue(txtCutOff.Text)
        If DateLastBill < CDueDate Then DateLastBill = CDueDate
        MonDue = DateAdd("M", 1, CDueDate)
        BillNo = DateDiff("D", CDueDate, DateSerial(Year(MonDue), Month(MonDue), 1))
        StartBN = BillNo
        cRow = 1:  DP = Val(flxAccts.TextMatrix(1, 14)): cArrAdv = 0
        flxLedger.Rows = 0
        TtlCredits = 0: cArrAdv = 0: Balance = Val(flxAccts.TextMatrix(1, 13))
        Do While Billy  'CDueDate <= DateLastBill Or ExpAcct
            If CDueDate = DateValue(flxAccts.TextMatrix(1, 4)) Then  'Current Due date = 1st due date:    DP and Adv 1st Monthly
                Do While DateValue(flxPay.TextMatrix(cRow, 0)) < CDueDate 'DateSerial(Year(CDueDate), Month(CDueDate), 6)
                    ORAmt = Val(flxPay.TextMatrix(cRow, 2))
                    Rebate = Val(flxPay.TextMatrix(cRow, 3))
                    If flxPay.TextMatrix(cRow, 4) = "DP" Then
                        InsMos = DP 'IIf(DP > 0 And flxLedger.Rows = 1, DP, DP - Val(Replace(flxLedger.TextMatrix(flxLedger.Rows - 1, 4), ",", "")) - (Val(Replace(flxLedger.TextMatrix(flxLedger.Rows - 1, 7), ",", "")) + Val(Replace(flxLedger.TextMatrix(flxLedger.Rows - 1, 8), ",", ""))))
                        TtlAmtDue = InsMos '+ IIf(InStr(flxLedger.TextMatrix(flxLedger.Rows - 1, 3), "DP") > 0, 0, CArrAdv)
                        If DP > 0 Then
                            sAdvArr = IIf(TtlAmtDue - (ORAmt + Rebate) > 0, "Prtl DP", "Full DP")
                        Else
                            sAdvArr = ""
                            TtlAmtDue = cArrAdv 'new 6/18/2011
                        End If
                        DP = IIf(TtlAmtDue - (ORAmt + Rebate) > 0, TtlAmtDue - (ORAmt + Rebate), 0)
                    Else
                        InsMos = 0
                        TtlAmtDue = cArrAdv 'IIf(flxLedger.Rows = 1, 0, Val(Replace(flxLedger.TextMatrix(flxLedger.Rows - 1, 4), ",", "")) - (Val(Replace(flxLedger.TextMatrix(flxLedger.Rows - 1, 7), ",", "")) + Val(Replace(flxLedger.TextMatrix(flxLedger.Rows - 1, 8), ",", ""))))
                        If flxPay.TextMatrix(cRow, 4) = "INST" Then
                            If DP > 0 Then
                                InsMos = DP 'IIf(DP > 0 And flxLedger.Rows = 1, DP, DP - Val(Replace(flxLedger.TextMatrix(flxLedger.Rows - 1, 4), ",", "")) - (Val(Replace(flxLedger.TextMatrix(flxLedger.Rows - 1, 7), ",", "")) + Val(Replace(flxLedger.TextMatrix(flxLedger.Rows - 1, 8), ",", ""))))
                                TtlAmtDue = InsMos + cArrAdv '+ IIf(InStr(flxLedger.TextMatrix(flxLedger.Rows - 1, 3), "DP") > 0, 0, CArrAdv)
                                sAdvArr = IIf(TtlAmtDue - (ORAmt + Rebate) > 0, "Prtl DP", "Full DP")
                                DP = IIf(DP > 0, DP - (ORAmt + Rebate), 0)
                            Else
                                sAdvArr = "" 'IIf(sAdvArr <> "M.I.", "M.I.", "")
                            End If
                        Else
                            sAdvArr = flxPay.TextMatrix(cRow, 4)
                            If sAdvArr = "DMCM" And DP > 0 Then TtlAmtDue = DP 'new 6/18/2011
                            DP = IIf(DP > 0, DP - (ORAmt + Rebate), 0)
                        End If
                    End If
                    TtlCredits = TtlCredits + ORAmt + Rebate
                    Balance = Balance - (ORAmt + Rebate)
                    cArrAdv = TtlAmtDue - (ORAmt + Rebate)
                    flxLedger.AddItem vbTab & IIf(flxPay.TextMatrix(cRow, 4) = "DP", flxPay.TextMatrix(cRow, 0), "") & vbTab & IIf(InsMos > 0, Format(InsMos, "#,##0.00"), "") & vbTab & sAdvArr & vbTab & IIf(TtlAmtDue > 0, IIf(sAdvArr = "DMCM", "", Format(TtlAmtDue, "#,##0.00")), "") & vbTab & flxPay.TextMatrix(cRow, 0) & vbTab & flxPay.TextMatrix(cRow, 1) & vbTab & Format(ORAmt, "#,##0.00") & vbTab & IIf(Rebate <> 0, Format(Rebate, "#,##0.00"), "") & vbTab & Format(TtlCredits, "#,##0.00") & vbTab & Format(Balance, "#,##0.00") & vbTab & cArrAdv
                    cRow = cRow + 1
                    If cRow > flxPay.Rows - 1 Then Exit Do
                Loop
            End If
            If cRow > flxPay.Rows - 1 Then
                DispBillOnly = True
                ExpAcct = False
            Else
                DispBillOnly = False
            End If
            If BillNo = StartBN Then  'IDENTIFY START OF BILLING
                InsMos = BillNo * Round(flxAccts.TextMatrix(1, 3), 2)
                If TtlCredits = 0 Then InsMos = InsMos + DP 'Add DP kng BillNo=1 and No Payments Made
            Else
                'MonDue = DateAdd("D", Val(flxAccts.TextMatrix(1, 5)), flxAccts.TextMatrix(1, 4))
                'myCut = IIf(DateAdd("D", Val(flxAccts.TextMatrix(1, 5)), flxAccts.TextMatrix(1, 4)) <= CDueDate, Day(CDueDate), 1)
                
                'nabilin nga adlaw sugod due date hantod Final Date of Billing < 30 pero terms kuhaan og Bill No > 30 pa kay naa Stop Billing w/ Redemption
                If DateDiff("d", CDueDate, FinalDateBill) <= 30 And (Val(flxAccts.TextMatrix(1, 5)) - PrevBN) > 30 And XpRdm Then
                    InsMos = DateDiff("D", DateSerial(Year(CDueDate), Month(CDueDate), 1), DateSerial(Year(MonDue), Month(MonDue), 1)) * Round(flxAccts.TextMatrix(1, 3), 2)
                
                'nabilin nga adlaw sugod due date hantod Final Date of Billing < 30 pero terms kuhaan og Bill No < 30 npod kay naa Stop Billing w/ Redemption
                ElseIf DateDiff("d", CDueDate, FinalDateBill) <= 30 And (Val(flxAccts.TextMatrix(1, 5)) - PrevBN) < 30 And XpRdm Then
                    InsMos = (Val(flxAccts.TextMatrix(1, 5)) - PrevBN) * Round(flxAccts.TextMatrix(1, 3), 2)  'nabilin nga billing days * daily due
                
                ElseIf DateDiff("d", CDueDate, DateAdd("D", Val(flxAccts.TextMatrix(1, 5)), flxAccts.TextMatrix(1, 4))) < 30 Then  'Compute Ins. Mos. for last billing
                    myCut = Day(DateAdd("D", Val(flxAccts.TextMatrix(1, 5)), flxAccts.TextMatrix(1, 4)))
                    InsMos = DateDiff("D", DateSerial(Year(CDueDate), Month(CDueDate), 1), DateSerial(Year(CDueDate), Month(CDueDate), myCut)) * Round(flxAccts.TextMatrix(1, 3), 2)
                Else
                    myCut = IIf(Val(flxAccts.TextMatrix(1, 5)) = BillNo, Day(DateAdd("D", Val(flxAccts.TextMatrix(1, 5)), flxAccts.TextMatrix(1, 4))), 1)
                    InsMos = DateDiff("D", DateSerial(Year(CDueDate), Month(CDueDate), 1), DateSerial(Year(MonDue), Month(MonDue), myCut)) * Round(flxAccts.TextMatrix(1, 3), 2)
                End If
                CDueDate = DateSerial(Year(CDueDate), Month(CDueDate), 1)   'Change Day to 1 after 1st due date
            End If
            If Balance = 0 And cRow >= flxPay.Rows - 1 Then Exit Do         'NEW ENTRY
            sAdvArr = IIf(cArrAdv <> 0, FormatNumber(cArrAdv, 2, vbTrue, vbTrue), "")
            If BillNo <= Val(flxAccts.TextMatrix(1, 5)) Then    'New Update 3/23/2012 for accounts w/ stop billing due to deposit pra mo post jud sa last bill
                TtlAmtDue = InsMos + cArrAdv
                cArrAdv = TtlAmtDue
                flxLedger.AddItem BillNo & vbTab & CDueDate & vbTab & Format(InsMos, "#,##0.00") & vbTab & sAdvArr & vbTab & FormatNumber(TtlAmtDue, 2, vbTrue, vbTrue) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & cArrAdv
                PrevBN = BillNo
            ElseIf BillNo > Val(flxAccts.TextMatrix(1, 5)) And PrevBN <= Val(flxAccts.TextMatrix(1, 5)) And XpRdm Then
                TtlAmtDue = InsMos + cArrAdv
                cArrAdv = TtlAmtDue
                flxLedger.AddItem flxAccts.TextMatrix(1, 5) & vbTab & CDueDate & vbTab & Format(InsMos, "#,##0.00") & vbTab & sAdvArr & vbTab & FormatNumber(TtlAmtDue, 2, vbTrue, vbTrue) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & cArrAdv
                PrevBN = BillNo
            End If
            flxLedger.TopRow = flxLedger.Rows - 1
            If Not DispBillOnly Then
                'If IsDate(flxAccts.TextMatrix(1, 4)) Then
                '    k = Day(DateValue(flxAccts.TextMatrix(1, 4)))
                '    If k > 6 Then k = 6
                'End If
                Do While DateValue(flxPay.TextMatrix(cRow, 0)) <= DateAdd("m", 1, DateSerial(Year(CDueDate), Month(CDueDate), 6)) Or BillNo > Val(flxAccts.TextMatrix(1, 5)) 'DateAdd("m", 1, DateSerial(Year(CDueDate), Month(CDueDate), 6))
                    ORAmt = Val(flxPay.TextMatrix(cRow, 2))
                    Rebate = Val(flxPay.TextMatrix(cRow, 3))
                    TtlCredits = TtlCredits + ORAmt + Rebate
                    Balance = Balance - (ORAmt + Rebate)
                    If flxPay.TextMatrix(cRow, 4) = "INST" Or flxPay.TextMatrix(cRow, 4) = "DP" Then
                        cArrAdv = TtlAmtDue - (ORAmt + Rebate)
                        If flxLedger.TextMatrix(flxLedger.Rows - 1, 5) = "" Then
                            flxLedger.TextMatrix(flxLedger.Rows - 1, 5) = flxPay.TextMatrix(cRow, 0)
                            flxLedger.TextMatrix(flxLedger.Rows - 1, 6) = flxPay.TextMatrix(cRow, 1)
                            flxLedger.TextMatrix(flxLedger.Rows - 1, 7) = Format(ORAmt, "#,##0.00")
                            flxLedger.TextMatrix(flxLedger.Rows - 1, 8) = IIf(Rebate <> 0, Format(Rebate, "#,##0.00"), "")
                            flxLedger.TextMatrix(flxLedger.Rows - 1, 9) = Format(TtlCredits, "#,##0.00")
                            flxLedger.TextMatrix(flxLedger.Rows - 1, 10) = Format(Balance, "#,##0.00")
                            flxLedger.TextMatrix(flxLedger.Rows - 1, 11) = cArrAdv
                        Else
                            flxLedger.AddItem vbTab & vbTab & vbTab & vbTab & vbTab & flxPay.TextMatrix(cRow, 0) & vbTab & flxPay.TextMatrix(cRow, 1) & vbTab & Format(ORAmt, "#,##0.00") & vbTab & IIf(Rebate <> 0, Format(Rebate, "#,##0.00"), "") & vbTab & Format(TtlCredits, "#,##0.00") & vbTab & Format(Balance, "#,##0.00") & vbTab & cArrAdv
                        End If
                        TtlAmtDue = cArrAdv
                    Else
                        If flxPay.TextMatrix(cRow, 4) = "REDEEM" Then   'NEW ENTRY
                            TtlCredits = flxPay.TextMatrix(cRow, 2)
                            Balance = Val(flxAccts.TextMatrix(1, 13)) - TtlCredits
                            flxLedger.AddItem vbTab & vbTab & vbTab & flxPay.TextMatrix(cRow, 4) & vbTab & vbTab & flxPay.TextMatrix(cRow, 0) & vbTab & flxPay.TextMatrix(cRow, 1) & vbTab & Format(ORAmt, "#,##0.00") & vbTab & IIf(Rebate <> 0, Format(Rebate, "#,##0.00"), "") & vbTab & Format(TtlCredits, "#,##0.00") & vbTab & Format(Balance, "#,##0.00") & vbTab & cArrAdv
                        Else
                            If flxPay.TextMatrix(cRow, 4) = "DEPOSIT" Then 'NEW ENTRY UPTO THIS LINE
                                If cRow + 1 <= flxPay.Rows - 1 Then
                                    If flxPay.TextMatrix(cRow + 1, 4) = "REDEEM" Then
                                        j = DateDiff("d", DateValue(flxPay.TextMatrix(cRow, 0)), DateValue(flxPay.TextMatrix(cRow + 1, 0)))
                                        XpRdm = True
                                        If j >= 15 Then
                                            CDueDate = DateSerial(Year(DateValue(flxPay.TextMatrix(cRow + 1, 0))), Month(DateValue(flxPay.TextMatrix(cRow + 1, 0))), Day(DateValue(flxAccts.TextMatrix(1, 4))))
                                        End If
                                    End If
                                End If
                            Else
                                cArrAdv = cArrAdv - (ORAmt + Rebate)
                                TtlAmtDue = cArrAdv
                            End If
                            flxLedger.AddItem vbTab & vbTab & vbTab & flxPay.TextMatrix(cRow, 4) & vbTab & vbTab & flxPay.TextMatrix(cRow, 0) & vbTab & flxPay.TextMatrix(cRow, 1) & vbTab & Format(ORAmt, "#,##0.00") & vbTab & IIf(Rebate <> 0, Format(Rebate, "#,##0.00"), "") & vbTab & Format(TtlCredits, "#,##0.00") & vbTab & Format(Balance, "#,##0.00") & vbTab & cArrAdv
                        End If
                    End If
                    cRow = cRow + 1
                    If cRow > flxPay.Rows - 1 Then Exit Do
                Loop
            End If
            CDueDate = DateAdd("M", 1, CDueDate)
            MonDue = DateAdd("M", 1, CDueDate)
            If DateDiff("d", CDueDate, FinalDateBill) <= 30 And Val(flxAccts.TextMatrix(1, 5)) - BillNo > 30 And XpRdm Then
                BillNo = BillNo + DateDiff("D", DateSerial(Year(CDueDate), Month(CDueDate), 1), DateSerial(Year(MonDue), Month(MonDue), 1))
            ElseIf Month(CDueDate) = Month(DateAdd("D", Val(flxAccts.TextMatrix(1, 5)), flxAccts.TextMatrix(1, 4))) And Year(CDueDate) = Year(DateAdd("D", Val(flxAccts.TextMatrix(1, 5)), flxAccts.TextMatrix(1, 4))) Then
                BillNo = BillNo + DateDiff("D", DateSerial(Year(CDueDate), Month(CDueDate), 1), DateSerial(Year(CDueDate), Month(CDueDate), Day(DateLastBill)))
            Else
                BillNo = BillNo + DateDiff("D", DateSerial(Year(CDueDate), Month(CDueDate), 1), DateSerial(Year(MonDue), Month(MonDue), 1))
            End If
            Billy = False
            If CDueDate = DateAdd("M", 1, flxAccts.TextMatrix(1, 4)) Then CDueDate = DateSerial(Year(CDueDate), Month(CDueDate), 1)  'After 1st Due Date dapat next duedate bill on Day 1 of the month
            If CDueDate <= DateLastBill Then Billy = True
            If Not Billy Then
                If cRow <= flxPay.Rows - 1 Then Billy = True
                'If DateDiff("m", DateValue(flxAccts.TextMatrix(1, 4)), DateValue(txtCutOff.Text)) > Val(flxAccts.TextMatrix(1, 5)) Then Billy = IIf(cRow > flxPay.Rows - 1, False, True)
            End If
            If Month(DateValue(txtCutOff.Text)) = Month(CDueDate) And Year(DateValue(txtCutOff.Text)) = Year(CDueDate) Then Billy = False
        Loop
    End If
End Sub

Private Function GetCODAmt(AcctNo As String) As Single
    Dim rsCOD As New ADODB.Recordset
    Dim Qry As String
    
    With rsCOD
        Qry = "SELECT SRP, LCP From unitordered WHERE (TERMTYPE ='C' AND ACCOUNTNO ='" & AcctNo & "');"
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            If !SRP > 0 Then
                GetCODAmt = !SRP
            Else
                GetCODAmt = !LCP
            End If
        Else
            GetCODAmt = 0
        End If
        .Close
    End With
End Function

Private Sub ReprocessAging()
    Dim Qry As String, Tmp As String, MyAcct As String, DateSClosed As Date, dFrom As Date, dTo As Date, tmpDate As Date, DP1date As Date, DP2date As Date, TmpDate2 As Date, TmpDate3 As Date
    Dim i As Integer, j As Integer, ans As Integer, TtlDays As Integer, PrvDP As Currency, AdvAmt As Currency, RebComm As Currency, SRP As Currency, COD As Currency
    Dim cArrAdv As Currency, Mi As Currency, BalDP As Currency, MyDP As Currency, Due1 As Currency, Due2 As Currency, Due3 As Currency, Due4 As Currency, ActArrAdv As Currency
    Dim Mi2 As Currency, Mi3 As Currency, Mi4 As Currency
    
    For i = 4 To 23
        flxAge1.TextMatrix(1, i) = ""
    Next i
    With flxAccts
        MyAcct = .TextMatrix(1, 1)
        dFrom = DateSerial(Val(txtYr.Text), Val(txtMo.Text), 7)
        dTo = DateAdd("m", 1, dFrom)
        'Outstanding Balance
        flxAge1.TextMatrix(1, 15) = Format(GetFromLedger(3), "#,##0.00")
        If Val(Replace(flxAge1.TextMatrix(1, 15), ",", "")) = 0 Then
            tmpDate = DateSerial(Val(txtYr.Text), Val(txtMo.Text), 1)
            TmpDate2 = DateAdd("m", 1, tmpDate)
            If DateValue(.TextMatrix(1, 7)) >= tmpDate And DateValue(.TextMatrix(1, 7)) <= DateAdd("d", -1, TmpDate2) Then
                flxAge1.TextMatrix(1, 22) = "Closed"   'Remarks
                flxAge1.TextMatrix(1, 23) = "New Sale"
            ElseIf IsDate(.TextMatrix(1, 4)) Then
                If DateValue(.TextMatrix(1, 4)) >= dTo Then
                    flxAge1.TextMatrix(1, 22) = "Closed"
                    flxAge1.TextMatrix(1, 23) = "Not Yet Due"
                Else
                    flxAge1.TextMatrix(1, 22) = "Closed"   'Remarks
                    flxAge1.TextMatrix(1, 23) = "Good"
                End If
            Else
                flxAge1.TextMatrix(1, 22) = "Closed"   'Remarks
                flxAge1.TextMatrix(1, 23) = "Good"
            End If
        End If
        If .TextMatrix(1, 6) = "C" Then
            flxAge1.TextMatrix(1, 4) = 0 'Rmng Terms
            flxAge1.TextMatrix(1, 16) = "" 'iif(isdate( Day(.TextMatrix(1, 4))  'Due Date
            dFrom = DateSerial(Val(txtYr.Text), Val(txtMo.Text), 7)
            dTo = DateAdd("m", 1, dFrom)
            'dTo = DateAdd("d", -1, dTo)
            Qry = "SELECT PAYDATE From PAYMENTS WHERE (PAYDATE>='" & Format(dFrom, "yyyy-mm-dd") & "' And PAYDATE<'" & Format(dTo, "yyyy-mm-dd") & "' AND ACCOUNTNO='" & MyAcct & "' AND CASH>0) ORDER BY PAYDATE DESC;"
            Tmp = GetFieldValue(Qry)
            If Tmp <> "" Then flxAge1.TextMatrix(1, 17) = Day(Tmp) 'Date Paid
            Qry = "SELECT Sum(PAYMENTS.CASH) AS SumOfCASH From PAYMENTS WHERE (PAYDATE>='" & Format(dFrom, "yyyy-mm-dd") & "' And PAYDATE<'" & Format(dTo, "yyyy-mm-dd") & "' AND ACCOUNTNO='" & MyAcct & "');"
            flxAge1.TextMatrix(1, 18) = Format(Val(GetFieldValue(Qry)), "#,##0.00")  'AmtPd
            flxAge1.TextMatrix(1, 19) = ""  'DP
            flxAge1.TextMatrix(1, 20) = ""  'Rebate
            flxAge1.TextMatrix(1, 21) = ""  'Option
            'Qry = "SELECT [SRP]-[ACTUALDP] AS CODBAL From UNITORDERED WHERE ((([SRP]-[ACTUALDP])>0) AND ((UNITORDERED.TERMTYPE)='C') AND ((UNITORDERED.ACCOUNTNO)='" & MyAcct & "'));"
            SRP = GetCODAmt(MyAcct)
            If SRP > 0 Then
                Qry = "SELECT SUM(payments.CASH + payments.DOWNPAYMENT + payments.INSTALLMENT) As TtlCOD From payments Where payments.ACCOUNTNO = '" & MyAcct & "';"
                COD = Val(GetFieldValue(Qry))
                Qry = "SELECT SUM(ledgeradj.REFAMT) As FIELD_1 From ledgeradj Where ledgeradj.ACCOUNTNO = '" & MyAcct & "' AND ledgeradj.TRANSTYPE <> 'temp'"
                COD = COD + Val(GetFieldValue(Qry))
                'Get DMCM Amt for account
                If SRP > COD Then
                    flxAge1.TextMatrix(1, 22) = "Open"
                    flxAge1.TextMatrix(1, 15) = SRP - COD
                    flxAge1.TextMatrix(1, 23) = IIf(DateValue(flxAccts.TextMatrix(1, 7)) < DateSerial(Val(txtYr.Text), Val(txtMo.Text), 1), "Over Due", "New Sale")
                Else
                    flxAge1.TextMatrix(1, 22) = "Closed" 'Remarks
                    flxAge1.TextMatrix(1, 15) = 0
                    flxAge1.TextMatrix(1, 23) = IIf(DateValue(flxAccts.TextMatrix(1, 7)) < DateSerial(Val(txtYr.Text), Val(txtMo.Text), 1), "Over Due", "New Sale")
                    'flxAge1.TextMatrix(1, 23) = "Good"
                End If
            Else
                flxAge1.TextMatrix(1, 15) = 0
                flxAge1.TextMatrix(1, 22) = "Closed"  'Remarks
                flxAge1.TextMatrix(1, 23) = "New Sale"
            End If
            'flxAge1.TextMatrix(1, 22) = IIf(Val(GetFieldValue(Qry)) > 0, "OPEN", "CLOSED")   'Remarks
            'flxAge1.TextMatrix(1, 23) = IIf(DateValue(.TextMatrix(1, 7)) < dFrom, "Overdue", "Good")
        ElseIf UCase$(flxAge1.TextMatrix(1, 22)) = "CLOSED" Then
            'Amount Pd
            Qry = "SELECT Sum(PAYMENTS.DOWNPAYMENT) AS SumOfDP From PAYMENTS WHERE (PAYDATE>='" & Format(dFrom, "yyyy-mm-dd") & "' And PAYDATE<'" & Format(dTo, "yyyy-mm-dd") & "' AND ACCOUNTNO='" & MyAcct & "');"
            flxAge1.TextMatrix(1, 19) = Format(Val(GetFieldValue(Qry)), "#,##0.00")
            
            Qry = "SELECT Sum(PAYMENTS.INSTALLMENT) AS SumOfINSTALLMENT From PAYMENTS WHERE (PAYDATE>='" & Format(dFrom, "yyyy-mm-dd") & "' And PAYDATE<'" & Format(dTo, "yyyy-mm-dd") & "' AND ACCOUNTNO='" & MyAcct & "');"
            flxAge1.TextMatrix(1, 18) = Format(Val(GetFieldValue(Qry)), "#,##0.00")
            If Val(Replace(flxAge1.TextMatrix(1, 18), ",", "")) > 0 Then
                'Qry = "SELECT PAYMENTS.PAYDATE From PAYMENTS WHERE (((PAYMENTS.PAYDATE)>=#" & dFrom & "# And (PAYMENTS.PAYDATE)<#" & dTo & "#) AND ((PAYMENTS.INSTALLMENT)>0) AND ((PAYMENTS.ACCOUNTNO)='" & MyAcct & "')) ORDER BY PAYMENTS.PAYDATE DESC;"
                Qry = "SELECT PAYMENTS.PAYDATE From PAYMENTS WHERE (PAYDATE>='" & Format(dFrom, "yyyy-mm-dd") & "' And PAYDATE<'" & Format(dTo, "yyyy-mm-dd") & "' AND INSTALLMENT>0 AND ACCOUNTNO='" & MyAcct & "') ORDER BY PAYMENTS.PAYDATE DESC;"
                Tmp = GetFieldValue(Qry) 'Date Paid
                flxAge1.TextMatrix(1, 17) = IIf(IsDate(Tmp), Day(DateValue(Tmp)), "")
            End If
            'Rebate
            'Qry = "SELECT Sum(PAYMENTS.DISCOUNT) AS SumOfDISCOUNT From PAYMENTS WHERE (((PAYMENTS.PAYDATE)>=#" & dFrom & "# And (PAYMENTS.PAYDATE)<#" & dTo & "#) AND ((PAYMENTS.ACCOUNTNO)='" & MyAcct & "'));"
            Qry = "SELECT Sum(PAYMENTS.DISCOUNT) AS SumOfDISCOUNT From PAYMENTS WHERE (PAYDATE>='" & Format(dFrom, "yyyy-mm-dd") & "' And PAYDATE<'" & Format(dTo, "yyyy-mm-dd") & "' AND ACCOUNTNO='" & MyAcct & "');"
            flxAge1.TextMatrix(1, 20) = Format(Val(GetFieldValue(Qry)), "#,##0.00")
            flxAge1.TextMatrix(1, 4) = "0 " & .TextMatrix(1, 6)
            Qry = "SELECT ledgeradj.REFAMT From ledgeradj " & _
                 "Where ledgeradj.ACCOUNTNO='" & MyAcct & "' AND ledgeradj.REFDATE >= '" & Format(dFrom, "YYYY-MM-DD") & "' AND ledgeradj.REFDATE < '" & Format(dTo, "YYYY-MM-DD") & "' AND ledgeradj.TRANSTYPE = 'DEPOSIT'"
            flxAge1.TextMatrix(1, 21) = GetFieldValue(Qry)
            If Val(flxAge1.TextMatrix(1, 21)) > 0 Then
                flxAge1.TextMatrix(1, 22) = "Deposited"
                flxAge1.TextMatrix(1, 23) = "Good"
            Else
                'Qry = "SELECT Sum(LEDGERADJ.REFAMT) AS SumOfREFAMT From LEDGERADJ WHERE (((LEDGERADJ.TRANSTYPE)='DMCM') AND ((LEDGERADJ.ACCOUNTNO)='" & MyAcct & "'));"
                Qry = "SELECT Sum(REFAMT) as RefAmt From ledgeradj WHERE (TRANSTYPE ='DMCM' AND REFDATE >='" & Format(dFrom, "yyyy-mm-dd") & "' AND REFDATE <'" & Format(dTo, "yyyy-mm-dd") & "' AND ACCOUNTNO ='" & MyAcct & "');"
                RebComm = Val(GetFieldValue(Qry))
                If RebComm > 0 Then flxAge1.TextMatrix(1, 21) = Format(RebComm, "#,##0.00")
            End If
        Else
            'Call GetOptionAmt(MyAcct, 1)
            If flxAge1.TextMatrix(1, 22) = "" And flxLedger.Rows > 0 Then
                'Call DisplayLedger(1)
                cArrAdv = GetFromLedger(2)
                Mi = Val(.TextMatrix(1, 3))
                flxAge1.TextMatrix(1, 13) = IIf(cArrAdv < 0, Format(Abs(cArrAdv), "#,##0.00"), "")
                If .TextMatrix(1, 6) = "M" Then
                    flxAge1.TextMatrix(1, 4) = Val(.TextMatrix(1, 5)) - GetFromLedger(1)   'Rmng Terms
                    If Val(flxAge1.TextMatrix(1, 4)) = 0 Then
                        If DateDiff("m", DateValue(.TextMatrix(1, 4)), DateValue(txtCutOff.Text)) >= Val(.TextMatrix(1, 5)) Then flxAge1.TextMatrix(1, 22) = "Expired"
                    Else
                        flxAge1.TextMatrix(1, 6) = Format(Val(flxAge1.TextMatrix(1, 4)) * Val(.TextMatrix(1, 3)), "#,##0.00")
                    End If
                    DateSClosed = DateAdd("m", Val(.TextMatrix(1, 5)) - 1, DateValue(.TextMatrix(1, 4)))
                Else
                    If DateValue(.TextMatrix(1, 4)) > dTo Then 'new sale
                        flxAge1.TextMatrix(1, 4) = Val(.TextMatrix(1, 5))
                    ElseIf (DateValue(.TextMatrix(1, 7)) >= dFrom And DateValue(.TextMatrix(1, 7)) < dTo) And (DateValue(.TextMatrix(1, 4)) >= dFrom And DateValue(.TextMatrix(1, 4)) < dTo) Then
                        TtlDays = DateDiff("D", DateValue(.TextMatrix(1, 4)), DateSerial(Year(dTo), Month(dTo), 1))
                        flxAge1.TextMatrix(1, 4) = IIf(TtlDays > 0, Val(.TextMatrix(1, 5)) - TtlDays, .TextMatrix(1, 5))
                    Else
                        If DateAdd("d", Val(.TextMatrix(1, 5)), .TextMatrix(1, 4)) < DateSerial(Year(dFrom), Month(dFrom), 1) Then
                            Tmp = GetLastBillDate 'DateAdd("d", Val(.TextMatrix(1, 5)), .TextMatrix(1, 4))
                        Else
                            Tmp = IIf(DateValue(.TextMatrix(1, 4)) > DateSerial(Year(dFrom), Month(dFrom), 1), DateValue(.TextMatrix(1, 4)), DateSerial(Year(dFrom), Month(dFrom), 1))
                        End If
                        'flxAge1.TextMatrix(1, 4) = Val(.TextMatrix(1, 5)) - GetBillNoDaily(DateValue(Tmp)) 'DateDiff("d", DateValue(.TextMatrix(1, 4)), TmpDate)
                        If DateValue(.TextMatrix(1, 4)) >= DateSerial(Year(dTo), Month(dTo), 1) Then
                            flxAge1.TextMatrix(1, 4) = Val(.TextMatrix(1, 5))
                        Else
                            flxAge1.TextMatrix(1, 4) = Val(.TextMatrix(1, 5)) - GetBillNoDaily(DateValue(Tmp)) 'DateDiff("d", DateValue(.TextMatrix(1, 4)), TmpDate)
                        End If
                    End If
                    If Val(flxAge1.TextMatrix(1, 4)) = 0 Then
                        If DateDiff("d", DateValue(.TextMatrix(1, 4)), DateValue(txtCutOff.Text)) > Val(.TextMatrix(1, 5)) Then flxAge1.TextMatrix(1, 22) = "Expired"
                    Else
                        flxAge1.TextMatrix(1, 6) = Format(Val(flxAge1.TextMatrix(1, 4)) * Val(.TextMatrix(1, 3)), "#,##0.00")
                    End If
                    DateSClosed = DateAdd("d", Val(.TextMatrix(1, 5)), DateValue(.TextMatrix(1, 4)))
                End If  'Computation of DUE NEXT MONTH
                If DateSerial(txtYr.Text, txtMo.Text, Day(DateValue(.TextMatrix(1, 4)))) < DateSClosed Then 'Wla pa nag lapse
                    If .TextMatrix(1, 6) = "M" Then 'Monthly
                        BalDP = GetFromLedger(4)    'Total Credit
                        BalDP = IIf(BalDP < Val(.TextMatrix(1, 14)), Val(.TextMatrix(1, 14)) - BalDP, 0) 'if total credit<DP then
                        If DateValue(.TextMatrix(1, 4)) >= DateAdd("m", 1, dTo) Then    'first due date is greater than next month cutoff
                            flxAge1.TextMatrix(1, 5) = Format(BalDP, "#,##0.00")
                        ElseIf BalDP > 0 And DateValue(.TextMatrix(1, 4)) >= dTo Then
                            If DateValue(.TextMatrix(1, 7)) < DateSerial(Val(txtYr.Text), Val(txtMo.Text), 1) Then 'Previous month na purchase reflect bal in DP
                                flxAge1.TextMatrix(1, 7) = Format(BalDP, "#,##0.00")
                                flxAge1.TextMatrix(1, 5) = Format(Mi, "#,##0.00")
                                flxAge1.TextMatrix(1, 22) = "Bal. in DP"
                                cArrAdv = cArrAdv - BalDP
                            Else
                                flxAge1.TextMatrix(1, 5) = Format(Mi + BalDP, "#,##0.00")
                                cArrAdv = cArrAdv - BalDP
                            End If
                        Else
                            flxAge1.TextMatrix(1, 5) = IIf(Val(Replace(flxAge1.TextMatrix(1, 13), ",", "")) >= Mi, "", Format(Mi, "#,##0.00"))
                        End If
                    Else    'Daily
                        'if term=term then getlastbilldate
                        BalDP = GetFromLedger(4)    'Total Credit
                        BalDP = IIf(BalDP < Val(.TextMatrix(1, 14)), Val(.TextMatrix(1, 14)) - BalDP, 0) 'if total credit<DP then
                        If DateValue(.TextMatrix(1, 4)) >= DateAdd("m", 1, dTo) Then    'first due date is greater than next month cutoff
                            flxAge1.TextMatrix(1, 5) = Format(BalDP, "#,##0.00")
                        ElseIf BalDP > 0 And DateValue(.TextMatrix(1, 4)) >= dTo Then
                            If DateValue(.TextMatrix(1, 7)) < DateSerial(Val(txtYr.Text), Val(txtMo.Text), 1) Then 'Previous month na purchase reflect bal in DP
                                flxAge1.TextMatrix(1, 7) = Format(BalDP, "#,##0.00")
                                j = DateDiff("d", DateSerial(Year(dTo), Month(dTo), 1), DateAdd("m", 1, DateSerial(Year(dTo), Month(dTo), 1)))
                                flxAge1.TextMatrix(1, 5) = Format(j * Mi, "#,##0.00")
                                flxAge1.TextMatrix(1, 22) = "Bal. in DP"
                                cArrAdv = cArrAdv - BalDP
                            Else
                                j = DateDiff("d", DateSerial(Year(dTo), Month(dTo), 1), DateAdd("m", 1, DateSerial(Year(dTo), Month(dTo), 1)))
                                flxAge1.TextMatrix(1, 5) = Format((Mi * j) + BalDP, "#,##0.00")
                                cArrAdv = cArrAdv - BalDP
                            End If
                        Else
                            j = DateDiff("d", DateSerial(Year(dTo), Month(dTo), 1), DateAdd("m", 1, DateSerial(Year(dTo), Month(dTo), 1)))
                            flxAge1.TextMatrix(1, 5) = IIf(Val(Replace(flxAge1.TextMatrix(1, 13), ",", "")) >= Mi * j, "", Format(Mi * j, "#,##0.00"))
                        End If
                    End If
                End If
                If DateSClosed < dFrom Then ' Separate condition in getting Status and Remarks for Expired Accounts
                    Tmp = GetLastBillDate
                    If IsDate(Tmp) Then
                        If .TextMatrix(1, 6) = "M" Then
                            tmpDate = DateAdd("m", -1, dTo)
                            If DateValue(Tmp) >= tmpDate Then
                                j = 30
                            Else
                                tmpDate = DateAdd("m", -2, dTo)
                                If DateValue(Tmp) >= tmpDate Then
                                    j = 60
                                Else
                                    tmpDate = DateAdd("m", -3, dTo)
                                    If DateValue(Tmp) >= tmpDate Then
                                        j = 90
                                    Else
                                        j = 100
                                    End If
                                End If
                            End If
                        ElseIf .TextMatrix(1, 6) = "D" Then
                            j = DateDiff("d", DateValue(Tmp), DateSerial(Year(dTo), Month(dTo), Day(DateValue(Tmp))))
                        End If
                        If j <= 30 Then
                            If .TextMatrix(1, 6) = "M" Then 'Monthly
                                If cArrAdv > 0 And cArrAdv < Mi Then flxAge1.TextMatrix(1, 7) = Format(cArrAdv, "#,##0.00"): flxAge1.TextMatrix(1, 23) = "Good" '1 to 30 Days Overdue
                                If cArrAdv >= Mi Then flxAge1.TextMatrix(1, 7) = Format(Mi, "#,##0.00"): flxAge1.TextMatrix(1, 23) = "Bad" '31 to 60 Days Overdue
                                If cArrAdv > Mi And cArrAdv < Mi * 2 Then flxAge1.TextMatrix(1, 8) = Format(cArrAdv - Mi, "#,##0.00")
                                If cArrAdv >= Mi * 2 Then flxAge1.TextMatrix(1, 8) = Format(Mi, "#,##0.00") '61 to 90 Days Overdue
                                If cArrAdv > Mi * 2 And cArrAdv < Mi * 3 Then flxAge1.TextMatrix(1, 9) = Format(cArrAdv - (Mi * 2), "#,##0.00")
                                If cArrAdv >= Mi * 3 Then flxAge1.TextMatrix(1, 9) = Format(Mi, "#,##0.00") 'over 90 Days Overdue
                                If (cArrAdv >= Mi * 3 And cArrAdv < Mi * 4) Or cArrAdv >= Mi * 4 Then flxAge1.TextMatrix(1, 10) = Format(cArrAdv - (Mi * 3), "#,##0.00")
                            Else
                                tmpDate = IIf(DateValue(.TextMatrix(1, 4)) > DateSerial(Year(dFrom), Month(dFrom), 1), DateValue(.TextMatrix(1, 4)), DateSerial(Year(dFrom), Month(dFrom), 1))
                                Due1 = DateDiff("d", tmpDate, DateSerial(Year(dTo), Month(dTo), 1)) * Mi
                                If cArrAdv > 0 And cArrAdv < Due1 Then flxAge1.TextMatrix(1, 7) = Format(cArrAdv, "#,##0.00")
                                If cArrAdv >= Due1 Then flxAge1.TextMatrix(1, 7) = Format(Due1, "#,##0.00")
                                '31 to 60 Days Overdue
                                j = IIf(DateValue(.TextMatrix(1, 4)) > DateSerial(Year(DateAdd("m", -1, dFrom)), Month(DateAdd("m", -1, dFrom)), 1), Day(DateValue(.TextMatrix(1, 4))), 1)
                                TtlDays = DateDiff("d", DateSerial(Year(DateAdd("m", -1, dFrom)), Month(DateAdd("m", -1, dFrom)), j), DateSerial(Year(dFrom), Month(dFrom), 1))
                                Due2 = Mi * TtlDays
                                If cArrAdv > Due1 And cArrAdv < Due1 + Due2 Then flxAge1.TextMatrix(1, 8) = Format(cArrAdv - Due1, "#,##0.00")
                                If cArrAdv >= Due1 + Due2 Then flxAge1.TextMatrix(1, 8) = Format(Due2, "#,##0.00")
                                '61 to 90 Days Overdue
                                j = IIf(DateValue(.TextMatrix(1, 4)) > DateSerial(Year(DateAdd("m", -2, dFrom)), Month(DateAdd("m", -2, dFrom)), 1), Day(DateValue(.TextMatrix(1, 4))), 1)
                                TtlDays = DateDiff("d", DateSerial(Year(DateAdd("m", -2, dFrom)), Month(DateAdd("m", -2, dFrom)), j), DateSerial(Year(DateAdd("m", -1, dFrom)), Month(DateAdd("m", -1, dFrom)), 1))
                                Due3 = Mi * TtlDays
                                If cArrAdv > Due1 + Due2 And cArrAdv < Due1 + Due2 + Due3 Then flxAge1.TextMatrix(1, 9) = Format(cArrAdv - (Due1 + Due2), "#,##0.00")
                                If cArrAdv >= Due1 + Due2 + Due3 Then flxAge1.TextMatrix(1, 9) = Format(Due3, "#,##0.00")
                                'over 90 Days Overdue
                                j = IIf(DateValue(.TextMatrix(1, 4)) > DateSerial(Year(DateAdd("m", -3, dFrom)), Month(DateAdd("m", -3, dFrom)), 1), Day(DateValue(.TextMatrix(1, 4))), 1)
                                TtlDays = DateDiff("d", DateSerial(Year(DateAdd("m", -3, dFrom)), Month(DateAdd("m", -3, dFrom)), j), DateSerial(Year(DateAdd("m", -2, dFrom)), Month(DateAdd("m", -2, dFrom)), 1))
                                Due4 = Mi * TtlDays
                                If (cArrAdv >= Due1 + Due2 + Due3 And cArrAdv < Due1 + Due2 + Due3 + Due4) Or cArrAdv > Due1 + Due2 + Due3 Then flxAge1.TextMatrix(1, 10) = Format(cArrAdv - (Due1 + Due2 + Due3), "#,##0.00")
                            End If
                        ElseIf j > 30 And j <= 60 Then
                            If .TextMatrix(1, 6) = "M" Then 'Monthly
                                If cArrAdv > 0 And cArrAdv < Mi Then flxAge1.TextMatrix(1, 8) = Format(cArrAdv, "#,##0.00") '1 to 30 Days Overdue
                                If cArrAdv >= Mi Then flxAge1.TextMatrix(1, 8) = Format(Mi, "#,##0.00") '31 to 60 Days Overdue
                                If cArrAdv > Mi And cArrAdv < Mi * 2 Then flxAge1.TextMatrix(1, 9) = Format(cArrAdv - Mi, "#,##0.00")
                                If cArrAdv >= Mi * 2 Then flxAge1.TextMatrix(1, 9) = Format(Mi, "#,##0.00") '61 to 90 Days Overdue
                                If cArrAdv > Mi * 2 And cArrAdv < Mi * 3 Then flxAge1.TextMatrix(1, 10) = Format(cArrAdv - (Mi * 2), "#,##0.00")
                                If cArrAdv >= Mi * 3 Then flxAge1.TextMatrix(1, 10) = Format(cArrAdv - (Mi * 2), "#,##0.00")
                                flxAge1.TextMatrix(1, 23) = "Bad"
                            Else
                                tmpDate = IIf(DateValue(.TextMatrix(1, 4)) > DateSerial(Year(dFrom), Month(dFrom), 1), DateValue(.TextMatrix(1, 4)), DateSerial(Year(dFrom), Month(dFrom), 1))
                                Due1 = DateDiff("d", tmpDate, DateSerial(Year(dTo), Month(dTo), 1)) * Mi
                                If cArrAdv > 0 And cArrAdv < Due1 Then flxAge1.TextMatrix(1, 8) = Format(cArrAdv, "#,##0.00")
                                If cArrAdv >= Due1 Then flxAge1.TextMatrix(1, 8) = Format(Due1, "#,##0.00")
                                '31 to 60 Days Overdue
                                j = IIf(DateValue(.TextMatrix(1, 4)) > DateSerial(Year(DateAdd("m", -1, dFrom)), Month(DateAdd("m", -1, dFrom)), 1), Day(DateValue(.TextMatrix(1, 4))), 1)
                                TtlDays = DateDiff("d", DateSerial(Year(DateAdd("m", -1, dFrom)), Month(DateAdd("m", -1, dFrom)), j), DateSerial(Year(dFrom), Month(dFrom), 1))
                                Due2 = Mi * TtlDays
                                If cArrAdv > Due1 And cArrAdv < Due1 + Due2 Then flxAge1.TextMatrix(1, 9) = Format(cArrAdv - Due1, "#,##0.00")
                                If cArrAdv >= Due1 + Due2 Then flxAge1.TextMatrix(1, 9) = Format(Due2, "#,##0.00")
                                '61 to 90 Days Overdue
                                j = IIf(DateValue(.TextMatrix(1, 4)) > DateSerial(Year(DateAdd("m", -2, dFrom)), Month(DateAdd("m", -2, dFrom)), 1), Day(DateValue(.TextMatrix(1, 4))), 1)
                                TtlDays = DateDiff("d", DateSerial(Year(DateAdd("m", -2, dFrom)), Month(DateAdd("m", -2, dFrom)), j), DateSerial(Year(DateAdd("m", -1, dFrom)), Month(DateAdd("m", -1, dFrom)), 1))
                                Due3 = Mi * TtlDays
                                If cArrAdv > Due1 + Due2 And cArrAdv < Due1 + Due2 + Due3 Then flxAge1.TextMatrix(1, 10) = Format(cArrAdv - (Due1 + Due2), "#,##0.00")
                                If cArrAdv >= Due1 + Due2 + Due3 Then flxAge1.TextMatrix(1, 10) = Format(cArrAdv - (Due1 + Due2), "#,##0.00")
                                flxAge1.TextMatrix(1, 23) = "Bad"
                            End If
                        ElseIf j > 60 And j <= 90 Then
                            If .TextMatrix(1, 6) = "M" Then 'Monthly
                                If cArrAdv > 0 And cArrAdv < Mi Then flxAge1.TextMatrix(1, 9) = Format(cArrAdv, "#,##0.00") '1 to 30 Days Overdue
                                If cArrAdv >= Mi Then
                                    flxAge1.TextMatrix(1, 9) = Format(Mi, "#,##0.00") '31 to 60 Days Overdue
                                    flxAge1.TextMatrix(1, 10) = Format(cArrAdv - Mi, "#,##0.00")
                                End If
                                flxAge1.TextMatrix(1, 23) = "Bad"
                            Else
                                tmpDate = IIf(DateValue(.TextMatrix(1, 4)) > DateSerial(Year(dFrom), Month(dFrom), 1), DateValue(.TextMatrix(1, 4)), DateSerial(Year(dFrom), Month(dFrom), 1))
                                Due1 = DateDiff("d", tmpDate, DateSerial(Year(dTo), Month(dTo), 1)) * Mi
                                If cArrAdv > Due1 Then
                                    flxAge1.TextMatrix(1, 9) = Format(Due1, "#,##0.00")
                                    flxAge1.TextMatrix(1, 10) = Format(cArrAdv - Due1, "#,##0.00")
                                End If
                                flxAge1.TextMatrix(1, 23) = "Bad"
                            End If
                        Else    'to be continue
                            flxAge1.TextMatrix(1, 10) = Format(cArrAdv, "#,##0.00")
                            flxAge1.TextMatrix(1, 23) = "Bad"
                        End If
                        flxAge1.TextMatrix(1, 11) = Format(cArrAdv, "#,##0.00")
                    End If
                Else    'Normal Accounts - Compuation of OVER DUE
                    If .TextMatrix(1, 6) = "M" Then
                        If BalDP > 0 And flxAge1.TextMatrix(1, 22) <> "Bal. in DP" Then 'w/ balance DP
                            j = DateDiff("d", DateValue(.TextMatrix(1, 4)), dTo)
                            ActArrAdv = cArrAdv - BalDP
                            If j > 0 And j <= 30 Then '1stDueDate < cutoff  and not NEW SALES
                                flxAge1.TextMatrix(1, 7) = BalDP
                            ElseIf j > 30 And j <= 60 Then '1stDueDate 31 to 60
                                flxAge1.TextMatrix(1, 8) = BalDP
                            ElseIf j > 60 And j <= 90 Then '1stDueDate 61 to 90
                                flxAge1.TextMatrix(1, 9) = BalDP
                            ElseIf j > 90 Then '1stDueDate > 90
                                flxAge1.TextMatrix(1, 10) = BalDP
                            End If
                            If ActArrAdv > 0 And ActArrAdv < Mi Then flxAge1.TextMatrix(1, 7) = Format(Val(flxAge1.TextMatrix(1, 7)) + ActArrAdv, "#,##0.00")
                            If ActArrAdv >= Mi Then flxAge1.TextMatrix(1, 7) = Format(Val(flxAge1.TextMatrix(1, 7)) + Mi, "#,##0.00")
                            '31 to 60 Days Overdue
                            If ActArrAdv > Mi And ActArrAdv < Mi * 2 Then flxAge1.TextMatrix(1, 8) = Format(Val(flxAge1.TextMatrix(1, 8)) + (ActArrAdv - Mi), "#,##0.00")
                            If ActArrAdv >= Mi * 2 Then flxAge1.TextMatrix(1, 8) = Format(Val(flxAge1.TextMatrix(1, 8)) + Mi, "#,##0.00")
                            '61 to 90 Days Overdue
                            If ActArrAdv > Mi * 2 And ActArrAdv < Mi * 3 Then flxAge1.TextMatrix(1, 9) = Format(Val(flxAge1.TextMatrix(1, 9)) + (ActArrAdv - (Mi * 2)), "#,##0.00")
                            If ActArrAdv >= Mi * 3 Then flxAge1.TextMatrix(1, 9) = Format(Val(flxAge1.TextMatrix(1, 9)) + Mi, "#,##0.00")
                            'over 90 Days Overdue
                            If (ActArrAdv >= Mi * 3 And ActArrAdv < Mi * 4) Or ActArrAdv >= Mi * 4 Then flxAge1.TextMatrix(1, 10) = Format(Val(flxAge1.TextMatrix(1, 10)) + (ActArrAdv - (Mi * 3)), "#,##0.00")
                        ElseIf BalDP = 0 Then
                            '1 to 30 Days Overdue
                            If cArrAdv > 0 And cArrAdv < Mi Then flxAge1.TextMatrix(1, 7) = Format(cArrAdv, "#,##0.00")
                            If cArrAdv >= Mi Then flxAge1.TextMatrix(1, 7) = Format(Mi, "#,##0.00")
                            '31 to 60 Days Overdue
                            If cArrAdv > Mi And cArrAdv < Mi * 2 Then flxAge1.TextMatrix(1, 8) = Format(cArrAdv - Mi, "#,##0.00")
                            If cArrAdv >= Mi * 2 Then flxAge1.TextMatrix(1, 8) = Format(Mi, "#,##0.00")
                            '61 to 90 Days Overdue
                            If cArrAdv > Mi * 2 And cArrAdv < Mi * 3 Then flxAge1.TextMatrix(1, 9) = Format(cArrAdv - (Mi * 2), "#,##0.00")
                            If cArrAdv >= Mi * 3 Then flxAge1.TextMatrix(1, 9) = Format(Mi, "#,##0.00")
                            'over 90 Days Overdue
                            If (cArrAdv >= Mi * 3 And cArrAdv < Mi * 4) Or cArrAdv >= Mi * 3 Then flxAge1.TextMatrix(1, 10) = Format(cArrAdv - (Mi * 3), "#,##0.00")
                            j = DateDiff("D", .TextMatrix(1, 7), dTo)   'Adjustment for lapsed overdue
                            If j <= 60 And Val(Replace(flxAge1.TextMatrix(1, 9), ",", "")) > 0 Then
                                flxAge1.TextMatrix(1, 8) = Val(Replace(flxAge1.TextMatrix(1, 8), ",", "")) + Val(Replace(flxAge1.TextMatrix(1, 9), ",", ""))
                                flxAge1.TextMatrix(1, 9) = ""
                            End If
                            If j <= 90 And Val(flxAge1.TextMatrix(1, 10)) > 0 Then
                                flxAge1.TextMatrix(1, 9) = Val(Replace(flxAge1.TextMatrix(1, 9), ",", "")) + Val(Replace(flxAge1.TextMatrix(1, 10), ",", ""))
                                flxAge1.TextMatrix(1, 10) = ""
                            End If
                        End If
                    Else    'Daily
                        If BalDP > 0 And flxAge1.TextMatrix(1, 22) <> "Bal. in DP" Then 'w/ balance DP
                            j = DateDiff("d", DateValue(.TextMatrix(1, 4)), dTo)
                            ActArrAdv = cArrAdv - BalDP
                            If j > 0 And j <= 30 Then '1stDueDate < cutoff  and not NEW SALES
                                 flxAge1.TextMatrix(1, 7) = BalDP
                            ElseIf j > 30 And j <= 60 Then '1stDueDate 31 to 60
                                 flxAge1.TextMatrix(1, 8) = BalDP
                            ElseIf j > 60 And j <= 90 Then '1stDueDate 61 to 90
                                 flxAge1.TextMatrix(1, 9) = BalDP
                            ElseIf j > 90 Then '1stDueDate > 90
                                flxAge1.TextMatrix(1, 10) = BalDP
                            End If
                            If DateValue(flxAccts.TextMatrix(1, 4)) >= DateSerial(Year(dFrom), Month(dFrom), 1) And DateValue(flxAccts.TextMatrix(1, 4)) < DateSerial(Year(dTo), Month(dTo), 1) Then
                                i = DateDiff("D", DateValue(flxAccts.TextMatrix(1, 4)), DateSerial(Year(dTo), Month(dTo), 1))
                                Mi1 = i * Val(flxAccts.TextMatrix(1, 3))
                            Else
                                i = DateDiff("D", DateSerial(Year(dFrom), Month(dFrom), 1), DateSerial(Year(dTo), Month(dTo), 1))
                                Mi1 = i * Val(flxAccts.TextMatrix(1, 3))
                            End If
                            If ActArrAdv > 0 And ActArrAdv < Mi1 Then flxAge1.TextMatrix(1, 7) = Format(Val(flxAge1.TextMatrix(1, 7)) + ActArrAdv, "#,##0.00")
                            If ActArrAdv >= Mi1 Then flxAge1.TextMatrix(1, 7) = Format(Val(flxAge1.TextMatrix(1, 7)) + Mi1, "#,##0.00")
                            '31 to 60 Days Overdue
                            tmpDate = DateAdd("M", 1, dTo)
                            i = DateDiff("D", DateSerial(Year(dTo), Month(dTo), 1), DateSerial(Year(tmpDate), Month(tmpDate), 1))
                            Mi2 = i * Val(flxAccts.TextMatrix(1, 3))
                            If ActArrAdv > Mi1 And ActArrAdv < (Mi1 + Mi2) Then flxAge1.TextMatrix(1, 8) = Format(Val(flxAge1.TextMatrix(1, 8)) + (ActArrAdv - Mi1), "#,##0.00")
                            If ActArrAdv >= (Mi1 + Mi2) Then flxAge1.TextMatrix(1, 8) = Format(Val(flxAge1.TextMatrix(1, 8)) + Mi2, "#,##0.00")
                            '61 to 90 Days Overdue
                            TmpDate2 = DateAdd("M", 1, tmpDate)
                            i = DateDiff("D", DateSerial(Year(tmpDate), Month(tmpDate), 1), DateSerial(Year(TmpDate2), Month(TmpDate2), 1))
                            Mi3 = i * Val(flxAccts.TextMatrix(1, 3))
                            If ActArrAdv > (Mi1 + Mi2) And ActArrAdv < (Mi1 + Mi2 + Mi3) Then flxAge1.TextMatrix(1, 9) = Format(Val(flxAge1.TextMatrix(1, 9)) + (ActArrAdv - (Mi1 + Mi2)), "#,##0.00")
                            If ActArrAdv >= (Mi1 + Mi12 + Mi3) Then flxAge1.TextMatrix(1, 9) = Format(Val(flxAge1.TextMatrix(1, 9)) + (Mi1 + Mi2 + Mi3), "#,##0.00")
                            'over 90 Days Overdue
                            TmpDate3 = DateAdd("M", 1, TmpDate2)
                            i = DateDiff("D", DateSerial(Year(TmpDate2), Month(TmpDate2), 1), DateSerial(Year(TmpDate3), Month(TmpDate3), 1))
                            Mi4 = i * Val(flxAccts.TextMatrix(1, 3))
                            If (ActArrAdv >= (Mi1 + Mi2 + Mi3) And ActArrAdv < (Mi1 + Mi2 + Mi3 + Mi4)) Or ActArrAdv >= (Mi1 + Mi2 + Mi3 + Mi4) Then flxAge1.TextMatrix(1, 10) = Format(Val(flxAge1.TextMatrix(1, 10)) + (ActArrAdv - (Mi1 + Mi2 + Mi3)), "#,##0.00")
                            
                        ElseIf BalDP = 0 Then
                            '1 to 30 Days Overdue
                            Due1 = DateDiff("d", DateSerial(Year(dFrom), Month(dFrom), 1), DateSerial(Year(dTo), Month(dTo), 1)) * Mi
                            If cArrAdv > 0 And cArrAdv < Due1 Then flxAge1.TextMatrix(1, 7) = Format(cArrAdv, "#,##0.00")
                            If cArrAdv >= Due1 Then flxAge1.TextMatrix(1, 7) = Format(Due1, "#,##0.00")
                            '31 to 60 Days Overdue
                            j = IIf(DateValue(.TextMatrix(1, 4)) > DateSerial(Year(DateAdd("m", -1, dFrom)), Month(DateAdd("m", -1, dFrom)), 1), Day(DateValue(.TextMatrix(1, 4))), 1)
                            TtlDays = DateDiff("d", DateSerial(Year(DateAdd("m", -1, dFrom)), Month(DateAdd("m", -1, dFrom)), j), DateSerial(Year(dFrom), Month(dFrom), 1))
                            Due2 = Mi * TtlDays
                            If cArrAdv > Due1 And cArrAdv < Due1 + Due2 Then flxAge1.TextMatrix(1, 8) = Format(cArrAdv - Due1, "#,##0.00")
                            If cArrAdv >= Due1 + Due2 Then flxAge1.TextMatrix(1, 8) = Format(Due2, "#,##0.00")
                            '61 to 90 Days Overdue
                            j = IIf(DateValue(.TextMatrix(1, 4)) > DateSerial(Year(DateAdd("m", -2, dFrom)), Month(DateAdd("m", -2, dFrom)), 1), Day(DateValue(.TextMatrix(1, 4))), 1)
                            TtlDays = DateDiff("d", DateSerial(Year(DateAdd("m", -2, dFrom)), Month(DateAdd("m", -2, dFrom)), j), DateSerial(Year(DateAdd("m", -1, dFrom)), Month(DateAdd("m", -1, dFrom)), 1))
                            Due3 = Mi * TtlDays
                            If cArrAdv > Due1 + Due2 And cArrAdv < Due1 + Due2 + Due3 Then flxAge1.TextMatrix(1, 9) = Format(cArrAdv - (Due1 + Due2), "#,##0.00")
                            If cArrAdv >= Due1 + Due2 + Due3 Then flxAge1.TextMatrix(1, 9) = Format(Due3, "#,##0.00")
                            'over 90 Days Overdue
                            j = IIf(DateValue(.TextMatrix(1, 4)) > DateSerial(Year(DateAdd("m", -3, dFrom)), Month(DateAdd("m", -3, dFrom)), 1), Day(DateValue(.TextMatrix(1, 4))), 1)
                            TtlDays = DateDiff("d", DateSerial(Year(DateAdd("m", -3, dFrom)), Month(DateAdd("m", -3, dFrom)), j), DateSerial(Year(DateAdd("m", -2, dFrom)), Month(DateAdd("m", -2, dFrom)), 1))
                            Due4 = Mi * TtlDays
                            If (cArrAdv >= Due1 + Due2 + Due3 And cArrAdv < Due1 + Due2 + Due3 + Due4) Or cArrAdv > Due1 + Due2 + Due3 Then flxAge1.TextMatrix(1, 10) = Format(cArrAdv - (Due1 + Due2 + Due3), "#,##0.00")
                        End If
                    End If
                    'Total Over Due
                    If DateValue(.TextMatrix(1, 4)) >= DateAdd("m", 1, dTo) Then
                        flxAge1.TextMatrix(1, 7) = ""
                        flxAge1.TextMatrix(1, 8) = ""
                        flxAge1.TextMatrix(1, 9) = ""
                    Else
                        flxAge1.TextMatrix(1, 11) = Format(Val(Replace(flxAge1.TextMatrix(1, 7), ",", "")) + Val(Replace(flxAge1.TextMatrix(1, 8), ",", "")) + Val(Replace(flxAge1.TextMatrix(1, 9), ",", "")) + Val(Replace(flxAge1.TextMatrix(1, 10), ",", "")), "#,##0.00")
                    End If
                End If
                'Get Remarks & Status for New Sales and Not yet due
                DP1date = DateSerial(txtYr.Text, txtMo.Text, 1)
                DP2date = DateAdd("m", 1, DP1date)
                DP2date = DateAdd("d", -1, DP2date)
                If DateValue(.TextMatrix(1, 7)) >= DP1date And DateValue(.TextMatrix(1, 7)) <= DP2date Then
                    flxAge1.TextMatrix(1, 22) = "New Sale"
                    flxAge1.TextMatrix(1, 23) = "New Sale"
                    If BalDP > 0 Then
                        If .TextMatrix(1, 6) = "M" Then 'Monthly
                            flxAge1.TextMatrix(1, 6) = Format(Val(Replace(flxAge1.TextMatrix(1, 6), ",", "")) + BalDP, "#,##0.00")
                            flxAge1.TextMatrix(1, 22) = "Bal. in DP"
                        Else
                            flxAge1.TextMatrix(1, 7) = Format(BalDP, "#,##0.00")
                            flxAge1.TextMatrix(1, 8) = ""
                            flxAge1.TextMatrix(1, 9) = ""
                            flxAge1.TextMatrix(1, 11) = Format(BalDP, "#,##0.00")
                        End If
                    ElseIf cArrAdv <= 0 And .TextMatrix(1, 6) = "D" Then            'new added code 2013-05-10 START
                        If DateValue(.TextMatrix(1, 4)) >= DP1date And DateValue(.TextMatrix(1, 4)) <= DP2date Then     'Date Purchase is within Aging Period
                            flxAge1.TextMatrix(1, 22) = "First Monthly"
                            flxAge1.TextMatrix(1, 23) = "Good"
                            flxAge1.TextMatrix(1, 13) = Format(Abs(cArrAdv), "#,##0.00")
                         End If
                    ElseIf cArrAdv > 0 And .TextMatrix(1, 6) = "D" Then
                        If DateValue(.TextMatrix(1, 4)) >= DP1date And DateValue(.TextMatrix(1, 4)) <= DP2date Then     'Date Purchase is within Aging Period
                            flxAge1.TextMatrix(1, 22) = "First Monthly"
                            flxAge1.TextMatrix(1, 23) = "Bad"
                            flxAge1.TextMatrix(1, 7) = Format(cArrAdv, "#,##0.00")
                            flxAge1.TextMatrix(1, 8) = ""
                            flxAge1.TextMatrix(1, 9) = ""
                            flxAge1.TextMatrix(1, 11) = Format(cArrAdv, "#,##0.00")
                            flxAge1.TextMatrix(1, 13) = 0                          'new added code 2013-05-10 END
                        End If
                    ElseIf cArrAdv < 0 Then
                        flxAge1.TextMatrix(1, 13) = Format(Abs(cArrAdv), "#,##0.00")
                    Else
                        Qry = "SELECT SUM(payments.DOWNPAYMENT) + SUM(payments.INSTALLMENT)+ SUM(payments.DISCOUNT) AS TotalDP From payments " & _
                              "Where (payments.DOWNPAYMENT>0 OR payments.INSTALLMENT>0) AND payments.ACCOUNTNO='" & MyAcct & "' AND payments.PAYDATE >= '" & Format(dFrom, "yyyy-mm-dd") & "' AND payments.PAYDATE < '" & Format(dTo, "yyyy-mm-dd") & "';"
                        MyDP = Val(GetFieldValue(Qry))
                        Qry = "SELECT ledgeradj.REFAMT From ledgeradj Where ledgeradj.ACCOUNTNO = '" & MyAcct & "' AND ledgeradj.REASONS LIKE '%REBATE%' AND ledgeradj.REFTYPE = 'APPLIED ADJUSTMENTS';"
                        MyDP = MyDP + Val(GetFieldValue(Qry))
                        If MyDP > Val(.TextMatrix(1, 14)) Then  'Payment>DP dapat Adv. uf walay overdue
                            flxAge1.TextMatrix(1, 13) = MyDP - Val(.TextMatrix(1, 14))
                            flxAge1.TextMatrix(1, 7) = ""
                            flxAge1.TextMatrix(1, 11) = ""
                        End If
                    End If
                    Qry = "SELECT unitordered.MGRREMARKS From unitordered Where unitordered.ACCOUNTNO = '" & MyAcct & "';"
                    Tmp = GetFieldValue(Qry)
                    If InStr(Tmp, "RECONSTRUCTED") > 0 Then
                        flxAge1.TextMatrix(1, 22) = "Recon"
                        flxAge1.TextMatrix(1, 23) = "Not Yet Due"
                        If DateValue(.TextMatrix(1, 4)) >= dFrom And DateValue(.TextMatrix(1, 4)) < dTo Then
                            flxAge1.TextMatrix(1, 22) = "First Monthly"
                            flxAge1.TextMatrix(1, 23) = IIf(cArrAdv > 0, "Bad", "Good")
                        End If
                    End If
                ElseIf DateValue(.TextMatrix(1, 7)) < dFrom And DateValue(.TextMatrix(1, 4)) >= dTo Then
                    If flxAge1.TextMatrix(1, 22) = "" Then flxAge1.TextMatrix(1, 22) = "Not Yet Due"
                    flxAge1.TextMatrix(1, 23) = "Not Yet Due"
                ElseIf DateValue(.TextMatrix(1, 4)) >= dFrom And DateValue(.TextMatrix(1, 4)) < dTo Then
                    If flxAge1.TextMatrix(1, 22) <> "New Sale" Then
                        flxAge1.TextMatrix(1, 22) = "First Monthly"
                        flxAge1.TextMatrix(1, 23) = IIf(cArrAdv > 0, "Bad", "Good")
                        If BalDP > 0 Then
                        
                        End If
                    End If
                ElseIf Val(flxAge1.TextMatrix(1, 8)) > 0 Or Val(flxAge1.TextMatrix(1, 9)) > 0 Or Val(flxAge1.TextMatrix(1, 10)) > 0 Then    '60 days overdue>0
                    If flxAge1.TextMatrix(1, 22) <> "Expired" Then flxAge1.TextMatrix(1, 22) = "Over Due"
                    flxAge1.TextMatrix(1, 23) = "Bad"
                ElseIf Val(flxAge1.TextMatrix(1, 7)) > 0 Then
                    If flxAge1.TextMatrix(1, 22) <> "Expired" Then flxAge1.TextMatrix(1, 22) = "Over Due"
                    flxAge1.TextMatrix(1, 23) = "Good"
                ElseIf cArrAdv <= 0 Then
                    flxAge1.TextMatrix(1, 22) = "Updated"
                    flxAge1.TextMatrix(1, 23) = "Good"
                End If
            End If
            'Collectible Next Month
            If Val(Replace(flxAge1.TextMatrix(1, 5), ",", "")) + Val(Replace(flxAge1.TextMatrix(1, 11), ",", "")) >= Val(Replace(flxAge1.TextMatrix(1, 15), ",", "")) Then
                flxAge1.TextMatrix(1, 12) = flxAge1.TextMatrix(1, 15)
            Else
                flxAge1.TextMatrix(1, 12) = Format(Val(Replace(flxAge1.TextMatrix(1, 5), ",", "")) + Val(Replace(flxAge1.TextMatrix(1, 11), ",", "")), "#,##0.00")
            End If
            'Installment Amount
            If Val(flxAge1.TextMatrix(1, 4)) > 0 Then
                If .TextMatrix(1, 6) = "M" Then
                    flxAge1.TextMatrix(1, 14) = Format(Mi, "#,##0.00")
                Else
                    If Val(flxAge1.TextMatrix(1, 4)) <> Val(.TextMatrix(1, 5)) Then
                        tmpDate = IIf(DateValue(.TextMatrix(1, 4)) > DateSerial(Year(dFrom), Month(dFrom), 1), DateValue(.TextMatrix(1, 4)), DateSerial(Year(dFrom), Month(dFrom), 1))
                        If tmpDate < DateSerial(Year(dTo), Month(dTo), 1) Then
                            Due1 = DateDiff("d", tmpDate, DateSerial(Year(dTo), Month(dTo), 1)) * Mi
                            flxAge1.TextMatrix(1, 14) = Format(Due1, "#,##0.00")
                        End If
                    End If
                End If
            End If
            'Due Date
            flxAge1.TextMatrix(1, 16) = Day(.TextMatrix(1, 4))
            'Amount Pd
            'Qry = "SELECT Sum(PAYMENTS.INSTALLMENT) AS SumOfINSTALLMENT From PAYMENTS WHERE (((PAYMENTS.PAYDATE)>=#" & dFrom & "# And (PAYMENTS.PAYDATE)<#" & dTo & "#) AND ((PAYMENTS.ACCOUNTNO)='" & MyAcct & "'));"
            Qry = "SELECT Sum(PAYMENTS.INSTALLMENT) AS SumOfINSTALLMENT From PAYMENTS WHERE (PAYDATE>='" & Format(dFrom, "yyyy-mm-dd") & "' And PAYDATE<'" & Format(dTo, "yyyy-mm-dd") & "' AND ACCOUNTNO='" & MyAcct & "');"
            flxAge1.TextMatrix(1, 18) = Format(Val(GetFieldValue(Qry)), "#,##0.00")
            'DP
            DP1date = DateSerial(txtYr.Text, txtMo.Text, 6)
            DP2date = DateAdd("m", 1, DP1date)
            DP2date = DateAdd("d", -1, DP2date)
            'Qry = "SELECT SUM(payments.DOWNPAYMENT) + SUM(payments.INSTALLMENT) AS TotalDP From payments " & _
                  "Where payments.ACCOUNTNO='" & MyAcct & "' AND payments.PAYDATE >= '" & Format(DP1date, "yyyy-mm-dd") & "' AND payments.PAYDATE <= '" & Format(DP2date, "yyyy-mm-dd") & "' AND payments.DOWNPAYMENT > 0"
            Qry = "SELECT SUM(payments.DOWNPAYMENT) + SUM(payments.INSTALLMENT) AS TotalDP From payments " & _
                  "Where payments.ACCOUNTNO='" & MyAcct & "' AND payments.PAYDATE >= '" & Format(dFrom, "yyyy-mm-dd") & "' AND payments.PAYDATE < '" & Format(dTo, "yyyy-mm-dd") & "';"
            'Qry = "SELECT Sum(PAYMENTS.DOWNPAYMENT) AS SumOfDOWNPAYMENT From PAYMENTS WHERE (PAYDATE>='" & Format(dFrom, "yyyy-mm-dd") & "' And PAYDATE<'" & Format(dTo, "yyyy-mm-dd") & "' AND ACCOUNTNO='" & MyAcct & "');"
            MyDP = Val(GetFieldValue(Qry))
            Qry = "SELECT SUM(payments.DOWNPAYMENT) + SUM(payments.INSTALLMENT) + SUM(payments.DISCOUNT) AS TotalDP From PAYMENTS WHERE (PAYDATE<'" & Format(dFrom, "yyyy-mm-dd") & "' AND ACCOUNTNO='" & MyAcct & "');"
            PrvDP = Val(GetFieldValue(Qry))
            Qry = "SELECT Sum(LEDGERADJ.REFAMT) AS SumOfREFAMT From LEDGERADJ WHERE (REFDATE<'" & Format(dFrom, "yyyy-mm-dd") & "' And TRANSTYPE<>'temp' AND ACCOUNTNO='" & MyAcct & "');"
            PrvDP = PrvDP + Val(GetFieldValue(Qry))
            Qry = "SELECT SUM(ledgeradj.REFAMT) As DMCMAMT From unitordered INNER JOIN ledgeradj ON (unitordered.ORDERID = ledgeradj.ORDERID) " & _
                  "Where unitordered.ACCOUNTNO = '" & MyAcct & "' AND unitordered.MGRREMARKS LIKE '%TRANSFERRED%' AND ledgeradj.REFDATE >= '" & Format(dFrom, "yyyy-mm-dd") & "' AND ledgeradj.REFDATE < '" & Format(dTo, "yyyy-mm-dd") & "' AND ledgeradj.TRANSTYPE <> 'temp';"
            PrvDP = PrvDP + Val(GetFieldValue(Qry))
            
            Qry = "SELECT Sum(LEDGERADJ.REFAMT) AS SumOfREFAMT From LEDGERADJ WHERE (REFDATE>='" & Format(dFrom, "yyyy-mm-dd") & "' And REFDATE<'" & Format(dTo, "yyyy-mm-dd") & "' AND TRANSTYPE<>'temp' AND ACCOUNTNO='" & MyAcct & "');"
            RebComm = Val(GetFieldValue(Qry))
            If RebComm > 0 Then
                flxAge1.TextMatrix(1, 21) = Format(RebComm, "#,##0.00")
            Else
                Qry = "SELECT unitordered.REBATECOMM From unitordered Where unitordered.ACCOUNTNO = '" & MyAcct & "';"
                RebComm = Val(GetFieldValue(Qry))
                If RebComm > 0 Then
                    Qry = "SELECT SUM(payments.DISCOUNT) AS REBCOM From PAYMENTS WHERE ACCOUNTNO='" & MyAcct & "' AND payments.PAYDATE >= '" & Format(dFrom, "yyyy-mm-dd") & "' AND payments.PAYDATE < '" & Format(dTo, "yyyy-mm-dd") & "';"
                    If Val(GetFieldValue(Qry)) = 0 Then RebComm = 0         'NI REFLECT NA ANG REBATE COMM SA PREVIOUS AGING
                End If
            End If
            If PrvDP > 0 And PrvDP < Val(.TextMatrix(1, 14)) And MyDP > 0 Then 'Ang previous nabayad wala pa ka full Down
                If PrvDP < Val(.TextMatrix(1, 14)) And flxAge1.TextMatrix(1, 22) = "First Monthly" Then
                    BalDP = Val(.TextMatrix(1, 14)) - (PrvDP + RebComm)
                    If MyDP >= BalDP Then
                        flxAge1.TextMatrix(1, 19) = Format(BalDP, "#,##0.00")
                        flxAge1.TextMatrix(1, 18) = Format(MyDP - BalDP, "#,##0.00")
                    Else
                        flxAge1.TextMatrix(1, 19) = Format(MyDP, "#,##0.00")
                        flxAge1.TextMatrix(1, 18) = ""
                    End If
                ElseIf MyDP >= Val(.TextMatrix(1, 14)) - (PrvDP + RebComm) Then
                    flxAge1.TextMatrix(1, 19) = Format(Val(.TextMatrix(1, 14)) - (PrvDP + RebComm), "#,##0.00")
                    flxAge1.TextMatrix(1, 18) = Format(MyDP - Val(Replace(flxAge1.TextMatrix(1, 19), ",", "")), "#,##0.00")
                Else
                    flxAge1.TextMatrix(1, 19) = Format(MyDP, "#,##0.00")
                End If
            ElseIf PrvDP > 0 And PrvDP = Val(.TextMatrix(1, 14)) Then 'Ang previous nabayad Nka full Down
                flxAge1.TextMatrix(1, 19) = 0
                flxAge1.TextMatrix(1, 18) = Format(MyDP, "#,##0.00")
            ElseIf PrvDP > 0 And PrvDP > Val(.TextMatrix(1, 14)) Then
                flxAge1.TextMatrix(1, 18) = Format(MyDP, "#,##0.00")
            ElseIf MyDP > Val(.TextMatrix(1, 14)) Then  'Sobra sa DP sa Payment ibutang
                Qry = "SELECT unitordered.REBATECOMM From unitordered Where unitordered.ACCOUNTNO = '" & MyAcct & "';"
                flxAge1.TextMatrix(1, 19) = Format(Val(.TextMatrix(1, 14)) - Val(GetFieldValue(Qry)), "#,##0.00")
                flxAge1.TextMatrix(1, 18) = Format(MyDP - Val(Replace(flxAge1.TextMatrix(1, 19), ",", "")), "#,##0.00")
            ElseIf MyDP < Val(.TextMatrix(1, 14)) And flxAge1.TextMatrix(1, 23) = "New Sale" Then
                Qry = "SELECT unitordered.REBATECOMM From unitordered Where unitordered.ACCOUNTNO = '" & MyAcct & "';"
                BalDP = Val(.TextMatrix(1, 14)) - Val(GetFieldValue(Qry))
                If MyDP > BalDP Then
                    flxAge1.TextMatrix(1, 19) = Format(BalDP, "#,##0.00")
                    flxAge1.TextMatrix(1, 18) = Format(MyDP - BalDP, "#,##0.00")
                Else
                    flxAge1.TextMatrix(1, 19) = Format(MyDP, "#,##0.00")
                    flxAge1.TextMatrix(1, 18) = ""
                End If
            Else
                flxAge1.TextMatrix(1, 19) = Format(MyDP, "#,##0.00")
                
            End If
            If Val(flxAge1.TextMatrix(1, 19)) > 0 Then
                'Qry = "SELECT PAYMENTS.PAYDATE From PAYMENTS WHERE (((PAYMENTS.PAYDATE)>=#" & dFrom & "# And (PAYMENTS.PAYDATE)<#" & dTo & "#) AND ((PAYMENTS.DOWNPAYMENT)>0) AND ((PAYMENTS.ACCOUNTNO)='" & MyAcct & "')) ORDER BY PAYMENTS.PAYDATE DESC;"
                Qry = "SELECT PAYDATE From PAYMENTS WHERE (PAYDATE>='" & Format(dFrom, "yyyy-mm-dd") & "' And PAYDATE<'" & Format(dTo, "yyyy-mm-dd") & "' AND DOWNPAYMENT>0 AND ACCOUNTNO='" & MyAcct & "') ORDER BY PAYMENTS.PAYDATE DESC;"
                Tmp = GetFieldValue(Qry) 'Date Paid
                If IsDate(Tmp) Then flxAge1.TextMatrix(1, 17) = Day(DateValue(Tmp))
            ElseIf Val(flxAge1.TextMatrix(1, 18)) > 0 Then
                'Qry = "SELECT PAYMENTS.PAYDATE From PAYMENTS WHERE (((PAYMENTS.PAYDATE)>=#" & dFrom & "# And (PAYMENTS.PAYDATE)<#" & dTo & "#) AND ((PAYMENTS.INSTALLMENT)>0) AND ((PAYMENTS.ACCOUNTNO)='" & MyAcct & "')) ORDER BY PAYMENTS.PAYDATE DESC;"
                Qry = "SELECT PAYDATE From PAYMENTS WHERE (PAYDATE>='" & Format(dFrom, "yyyy-mm-dd") & "' And PAYDATE<'" & Format(dTo, "yyyy-mm-dd") & "' AND INSTALLMENT>0 AND ACCOUNTNO='" & MyAcct & "') ORDER BY PAYMENTS.PAYDATE DESC;"
                Tmp = GetFieldValue(Qry) 'Date Paid
                If IsDate(Tmp) Then flxAge1.TextMatrix(1, 17) = Day(DateValue(Tmp))
            End If
            'Rebate
            'Qry = "SELECT Sum(PAYMENTS.DISCOUNT) AS SumOfDISCOUNT From PAYMENTS WHERE (((PAYMENTS.PAYDATE)>=#" & dFrom & "# And (PAYMENTS.PAYDATE)<#" & dTo & "#) AND ((PAYMENTS.ACCOUNTNO)='" & MyAcct & "'));"
            Qry = "SELECT Sum(PAYMENTS.DISCOUNT) AS SumOfDISCOUNT From PAYMENTS WHERE (PAYDATE>='" & Format(dFrom, "yyyy-mm-dd") & "' And PAYDATE<'" & Format(dTo, "yyyy-mm-dd") & "' AND ACCOUNTNO='" & MyAcct & "');"
            flxAge1.TextMatrix(1, 20) = Format(Val(GetFieldValue(Qry)), "#,##0.00")
            flxAge1.TextMatrix(1, 4) = flxAge1.TextMatrix(1, 4) & " " & .TextMatrix(1, 6)
        End If
    End With
    'Prompt to Save Aging if current
    LYRcon.Execute "UPDATE agingcurrent SET RMNGTRMS = '" & flxAge1.TextMatrix(1, 4) & "', DUENXTM = " & Val(Replace(flxAge1.TextMatrix(1, 5), ",", "")) & ", CURRENT = " & Val(Replace(flxAge1.TextMatrix(1, 6), ",", "")) & ", DUE30 = " & Val(Replace(flxAge1.TextMatrix(1, 7), ",", "")) & ", DUE60 = " & Val(Replace(flxAge1.TextMatrix(1, 8), ",", "")) & ", DUE90 = " & Val(Replace(flxAge1.TextMatrix(1, 9), ",", "")) & ", DUEOVER90 = " & Val(Replace(flxAge1.TextMatrix(1, 10), ",", "")) & ", TTLOVERDUE = " & Val(Replace(flxAge1.TextMatrix(1, 11), ",", "")) & ", " & _
                   "COLLNXTM = " & Val(Replace(flxAge1.TextMatrix(1, 12), ",", "")) & ", ADVPAY = " & Val(Replace(flxAge1.TextMatrix(1, 13), ",", "")) & ", INSTAMT = " & Val(Replace(flxAge1.TextMatrix(1, 14), ",", "")) & ", STNDNGBAL = " & Val(Replace(flxAge1.TextMatrix(1, 15), ",", "")) & ", DUEDATE = '" & flxAge1.TextMatrix(1, 16) & "', DATEPAID = '" & flxAge1.TextMatrix(1, 17) & "', AMTPAID = " & Val(Replace(flxAge1.TextMatrix(1, 18), ",", "")) & ", DP = " & Val(Replace(flxAge1.TextMatrix(1, 19), ",", "")) & ", REBATE = " & Val(Replace(flxAge1.TextMatrix(1, 20), ",", "")) & ", `OPTION` = " & Val(Replace(flxAge1.TextMatrix(1, 21), ",", "")) & ", REMARKS = '" & flxAge1.TextMatrix(1, 22) & "', `STATUS` = '" & flxAge1.TextMatrix(1, 23) & "' " & _
                   "Where agingcurrent.ORDERID = " & Val(flxAccts.TextMatrix(1, 2)) & " AND agingcurrent.AGEMONTH = " & txtMo.Text & " AND agingcurrent.AGEYEAR = " & txtYr.Text & ";"
End Sub


Function GetFromLedger(iCode As Integer) As Currency
    Dim i As Integer, dFrom As Date, Tmp As String, Qry As String
    
    'iCode: 1-MaxBillNo, 2-CrntArrAdv(Monthly), 3=OutStanding Balance, 4=Total Credits,
    '       5-Billing Date, 6-CrntArrAdv(Daily), 7-Ins.Mos.(Daily)
    
    With flxLedger
        If flxLedger.Rows = 0 Then
            GetFromLedger = 0
            Exit Function
        End If
        If iCode = 2 Then
            For i = .Rows - 1 To 0 Step -1
                If IsDate(.TextMatrix(i, 5)) Then
                    If DateValue(.TextMatrix(i, 5)) <= DateValue(txtCutOff.Text) Then
                        GetFromLedger = Val(.TextMatrix(i, 11))
                        Exit Function
                    End If
                ElseIf IsDate(.TextMatrix(i, 1)) Then
                    If DateValue(.TextMatrix(i, 1)) <= DateValue(txtCutOff.Text) Then
                        
                        GetFromLedger = Val(.TextMatrix(i, 11))
                        Exit Function
                    End If
                End If
            Next i
        Else
            For i = .Rows - 1 To 0 Step -1
                If iCode = 1 Then
                    If Val(.TextMatrix(i, 0)) > 0 Then
                        If DateValue(.TextMatrix(i, 1)) <= DateValue(txtCutOff.Text) Then
                            GetFromLedger = Val(.TextMatrix(i, 0))
                            Exit Function
                        End If
                    End If
                ElseIf iCode = 3 Then
                    If IsDate(.TextMatrix(i, 5)) Then
                        If DateValue(.TextMatrix(i, 5)) <= DateValue(txtCutOff.Text) And .TextMatrix(i, 6) <> "" Then
                            GetFromLedger = Val(Replace(.TextMatrix(i, 10), ",", ""))
                            Exit Function
                        End If
                    ElseIf .Rows = 1 And Val(Replace(.TextMatrix(i, 10), ",", "")) > 0 Then
                        GetFromLedger = Val(Replace(.TextMatrix(i, 10), ",", ""))
                        Exit Function
                    ElseIf IsDate(.TextMatrix(i, 1)) And Val(Replace(.TextMatrix(i, 10), ",", "")) > 0 Then
                        GetFromLedger = Val(Replace(.TextMatrix(i, 10), ",", ""))
                        Exit Function
                    End If
                ElseIf iCode = 4 Then
                    If IsDate(.TextMatrix(i, 5)) Then
                        If DateValue(.TextMatrix(i, 5)) <= DateValue(txtCutOff.Text) And Val(Replace(.TextMatrix(i, 9), ",", "")) > 0 Then
                            GetFromLedger = Val(Replace(.TextMatrix(i, 9), ",", ""))
                            Exit Function
                        End If
                    ElseIf .Rows = 1 And Val(Replace(.TextMatrix(i, 9), ",", "")) > 0 Then
                        GetFromLedger = Val(Replace(.TextMatrix(i, 9), ",", "")) > 0
                        Exit Function
                    End If
                ElseIf iCode = 5 Then
                    If IsDate(.TextMatrix(i, 1)) Then
                        GetFromLedger = .TextMatrix(i, 1)
                        Exit Function
                    End If
                ElseIf iCode = 6 Then
                    If IsDate(.TextMatrix(i, 5)) Then
                        If DateValue(.TextMatrix(i, 5)) <= DateValue(txtCutOff.Text) And .TextMatrix(i, 6) <> "" Then
                            GetFromLedger = Val(Replace(.TextMatrix(i, 11), ",", ""))
                            Exit Function
                        End If
                    End If
                    
                    
                ElseIf iCode = 7 Then
                    If IsDate(.TextMatrix(i, 1)) Then
                        GetFromLedger = Val(Replace(.TextMatrix(i, 2), ",", ""))
                        Exit Function
                    End If
                End If
            Next i
            GetFromLedger = 0
        End If
    End With
End Function


Function GetLastBillDate() As String
    Dim i As Integer, dFrom As Date
    
    'iCode: 1=Billing Date
    
    With flxLedger
        If flxLedger.Rows = 0 Then
            GetLastBillDate = ""
            Exit Function
        End If
        For i = .Rows - 1 To 0 Step -1
            If IsDate(.TextMatrix(i, 1)) Then
                GetLastBillDate = .TextMatrix(i, 1)
                Exit Function
            End If
        Next i
        GetLastBillDate = ""
    End With
End Function

Function GetBillNoDaily(AgePrd As Date) As Currency
    Dim i As Integer, dFrom As Date
    
    With flxLedger
        If flxLedger.Rows = 0 Then
            GetBillNoDaily = 0
            Exit Function
        End If
        For i = .Rows - 1 To 0 Step -1
            If IsDate(.TextMatrix(i, 1)) Then
                If DateValue(.TextMatrix(i, 1)) = AgePrd Then
                    GetBillNoDaily = Val(.TextMatrix(i, 0))
                    Exit Function
                End If
            End If
        Next i
        GetBillNoDaily = 0
    End With
End Function



