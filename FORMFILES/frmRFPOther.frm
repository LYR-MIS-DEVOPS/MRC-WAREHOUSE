VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRFPOther 
   Caption         =   "Request for Payment in Other Payables"
   ClientHeight    =   10260
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   14925
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10260
   ScaleWidth      =   14925
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab ssTabRBhist 
      Height          =   6015
      Left            =   2400
      TabIndex        =   76
      Top             =   2040
      Visible         =   0   'False
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      ForeColor       =   255
      TabCaption(0)   =   "TRANSACTION HISTORY"
      TabPicture(0)   =   "frmRFPOther.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblClose"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame4 
         Height          =   5535
         Left            =   120
         TabIndex        =   77
         Top             =   360
         Width           =   10335
         Begin MSFlexGridLib.MSFlexGrid flxRBTransHist 
            Height          =   5175
            Left            =   120
            TabIndex        =   78
            Top             =   240
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   9128
            _Version        =   393216
            Cols            =   5
            FormatString    =   "|<DATE |<TRANSACTION TYPE |<STATUS |<RESPONSIBLE"
         End
      End
      Begin VB.Label lblClose 
         Alignment       =   2  'Center
         Caption         =   " ( X )"
         Height          =   255
         Left            =   10080
         MousePointer    =   99  'Custom
         TabIndex        =   79
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.TextBox txtFlxValueDeduct 
      Height          =   285
      Left            =   600
      TabIndex        =   72
      Top             =   5040
      Visible         =   0   'False
      Width           =   3492
   End
   Begin VB.ComboBox cboFlxType 
      Height          =   288
      ItemData        =   "frmRFPOther.frx":001C
      Left            =   4320
      List            =   "frmRFPOther.frx":0029
      Style           =   2  'Dropdown List
      TabIndex        =   71
      Top             =   5040
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.TextBox txtFlxDeductReb 
      Height          =   285
      Left            =   10680
      TabIndex        =   70
      Top             =   5040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox cboFlxParticular 
      Height          =   288
      ItemData        =   "frmRFPOther.frx":003A
      Left            =   5880
      List            =   "frmRFPOther.frx":003C
      TabIndex        =   69
      Top             =   5040
      Visible         =   0   'False
      Width           =   4452
   End
   Begin VB.TextBox txtRemarks 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   11
      Tag             =   "XYZ"
      Top             =   8400
      Width           =   7452
   End
   Begin VB.ComboBox cboFlxAcctCode 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1440
      TabIndex        =   55
      Tag             =   "XYZ"
      Top             =   6960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtFlxAmnt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   11640
      TabIndex        =   54
      Tag             =   "XYZ"
      Text            =   "txtFlxAmnt"
      Top             =   6960
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ComboBox cboFlxAcctTitles 
      BackColor       =   &H00FFFFFF&
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
      Left            =   3360
      TabIndex        =   53
      Tag             =   "XYZ"
      Top             =   6960
      Visible         =   0   'False
      Width           =   5172
   End
   Begin VB.ComboBox cboFlxAcctType 
      Height          =   315
      ItemData        =   "frmRFPOther.frx":003E
      Left            =   9720
      List            =   "frmRFPOther.frx":0048
      Style           =   2  'Dropdown List
      TabIndex        =   52
      Top             =   6960
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Frame framHoldButtton 
      Caption         =   "DEDUCTION / ADD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1572
      Left            =   120
      TabIndex        =   61
      Tag             =   "xyz"
      Top             =   4440
      Width           =   14655
      Begin MSFlexGridLib.MSFlexGrid flxDeduct 
         Height          =   972
         Left            =   120
         TabIndex        =   8
         Tag             =   "XYZ"
         Top             =   240
         Width           =   14412
         _ExtentX        =   25426
         _ExtentY        =   1720
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         RowHeightMin    =   315
         BackColorBkg    =   -2147483633
         HighLight       =   2
         GridLines       =   2
         AllowUserResizing=   3
         FormatString    =   $"frmRFPOther.frx":005B
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
      Begin VB.Label lblNetPay 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   12240
         TabIndex        =   68
         Tag             =   "XYZ"
         Top             =   1200
         Width           =   2292
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "NET PAYABLE:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   10680
         TabIndex        =   67
         Tag             =   "XYZ"
         Top             =   1200
         Width           =   1572
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL ADD:"
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
         Left            =   4560
         TabIndex        =   65
         Tag             =   "XYZ"
         Top             =   1230
         Width           =   1215
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL LESS:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   64
         Tag             =   "XYZ"
         Top             =   1230
         Width           =   1212
      End
      Begin VB.Label lblTotLess 
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1440
         TabIndex        =   63
         Tag             =   "XYZ"
         Top             =   1230
         Width           =   2412
      End
      Begin VB.Label lblTotAdd 
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   5880
         TabIndex        =   62
         Tag             =   "XYZ"
         Top             =   1230
         Width           =   2412
      End
   End
   Begin VB.TextBox txtFlxRefNo 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3360
      MaxLength       =   30
      TabIndex        =   60
      Top             =   2400
      Visible         =   0   'False
      Width           =   2292
   End
   Begin VB.TextBox txtFlxParticular 
      Height          =   285
      Left            =   6240
      TabIndex        =   59
      Top             =   2400
      Visible         =   0   'False
      Width           =   4692
   End
   Begin VB.TextBox txtFlxBrchDpt 
      Height          =   285
      Left            =   600
      MaxLength       =   60
      TabIndex        =   58
      Top             =   2400
      Visible         =   0   'False
      Width           =   2292
   End
   Begin VB.TextBox txtFlxRFPAmnt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   11640
      TabIndex        =   57
      Top             =   2400
      Visible         =   0   'False
      Width           =   1692
   End
   Begin lvButton.lvButtons_H cmdHistory 
      Height          =   372
      Left            =   13200
      TabIndex        =   51
      Tag             =   "XYF"
      Top             =   8520
      Width           =   1572
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "SHOW HISTORY"
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
      Image           =   "frmRFPOther.frx":016C
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtCheckNo 
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      MaxLength       =   50
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Tag             =   "XYZ"
      Top             =   8760
      Width           =   1935
   End
   Begin VB.Timer tmrSync 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7440
      Top             =   0
   End
   Begin VB.Timer tmrGetRecord 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8400
      Top             =   0
   End
   Begin VB.Frame framHoldMid 
      Height          =   2652
      Left            =   120
      TabIndex        =   46
      Tag             =   "xyz"
      Top             =   1800
      Width           =   14655
      Begin VB.ComboBox cboFlxMainBranch 
         BackColor       =   &H00FFFFFF&
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
         Left            =   480
         TabIndex        =   85
         Tag             =   "XYZ"
         Top             =   600
         Visible         =   0   'False
         Width           =   2535
      End
      Begin MSFlexGridLib.MSFlexGrid flxRFP 
         Height          =   2052
         Left            =   120
         TabIndex        =   7
         Tag             =   "XYZ"
         Top             =   240
         Width           =   14412
         _ExtentX        =   25426
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         RowHeightMin    =   315
         BackColorBkg    =   -2147483633
         HighLight       =   2
         GridLines       =   2
         AllowUserResizing=   3
         FormatString    =   $"frmRFPOther.frx":595E
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
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   10560
         TabIndex        =   48
         Tag             =   "XYZ"
         Top             =   2352
         Width           =   1932
      End
      Begin VB.Label lblTotPOAmnt 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   12600
         TabIndex        =   47
         Tag             =   "XYZ"
         Top             =   2352
         Width           =   1812
      End
   End
   Begin VB.Frame fraHoldAcctTitles 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2292
      Left            =   120
      TabIndex        =   40
      Tag             =   "xyz"
      Top             =   6000
      Width           =   14655
      Begin VB.ComboBox cboFlxBranch 
         BackColor       =   &H00FFFFFF&
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
         Left            =   360
         TabIndex        =   82
         Tag             =   "XYZ"
         Top             =   960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cboTrans 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frmRFPOther.frx":5A65
         Left            =   1440
         List            =   "frmRFPOther.frx":5A67
         Locked          =   -1  'True
         TabIndex        =   9
         Tag             =   "XYZ"
         Top             =   240
         Width           =   9972
      End
      Begin MSFlexGridLib.MSFlexGrid flxVoucher 
         Height          =   1332
         Left            =   120
         TabIndex        =   10
         Tag             =   "XYZ"
         Top             =   600
         Width           =   14412
         _ExtentX        =   25426
         _ExtentY        =   2355
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         Enabled         =   -1  'True
         AllowUserResizing=   3
         FormatString    =   $"frmRFPOther.frx":5A69
      End
      Begin VB.Label Label19 
         Caption         =   "TOTAL CREDIT:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   45
         Tag             =   "XYZ"
         Top             =   1920
         Width           =   1452
      End
      Begin VB.Label lblTotCredit 
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1680
         TabIndex        =   44
         Tag             =   "XYZ"
         Top             =   1920
         Width           =   2412
      End
      Begin VB.Label Label18 
         Caption         =   "TOTAL DEBIT:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   4560
         TabIndex        =   43
         Tag             =   "XYZ"
         Top             =   1920
         Width           =   1332
      End
      Begin VB.Label lblTotDebit 
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   6000
         TabIndex        =   42
         Tag             =   "XYZ"
         Top             =   1920
         Width           =   2412
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "TRANSACTION:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   41
         Tag             =   "XYZ"
         Top             =   270
         Width           =   1305
      End
   End
   Begin VB.Frame fraHoldSignatory 
      Height          =   1050
      Left            =   120
      TabIndex        =   33
      Tag             =   "xyz"
      Top             =   9120
      Width           =   14652
      Begin VB.ComboBox cboStatus 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmRFPOther.frx":5B80
         Left            =   11520
         List            =   "frmRFPOther.frx":5B82
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Tag             =   "XYF"
         Top             =   600
         Width           =   3012
      End
      Begin VB.TextBox txtPrepBy 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Tag             =   "XYZ"
         Top             =   240
         Width           =   3132
      End
      Begin VB.TextBox txtApprvdBy 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   6480
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Tag             =   "XYZ"
         Top             =   600
         Width           =   3132
      End
      Begin VB.TextBox txtChckedBy 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Tag             =   "XYZ"
         Top             =   600
         Width           =   3132
      End
      Begin VB.TextBox txtNotedBy 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   6480
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Tag             =   "XYZ"
         Top             =   240
         Width           =   3132
      End
      Begin MSMask.MaskEdBox txtDocDateRec 
         Height          =   315
         Left            =   11520
         TabIndex        =   88
         Tag             =   "XYZ"
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12632256
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DOC DATE REC:"
         Height          =   375
         Left            =   9960
         TabIndex        =   90
         Tag             =   "xyz"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label28 
         Caption         =   "MM/DD/YYYY"
         Height          =   255
         Left            =   13440
         TabIndex        =   89
         Tag             =   "XYZ"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "STATUS:"
         Height          =   255
         Left            =   10440
         TabIndex        =   39
         Tag             =   "xyz"
         Top             =   615
         Width           =   975
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "CHECKED BY:"
         Height          =   252
         Left            =   120
         TabIndex        =   38
         Tag             =   "xyz"
         Top             =   648
         Width           =   1332
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "PREPARED BY:"
         Height          =   252
         Left            =   120
         TabIndex        =   37
         Tag             =   "xyz"
         Top             =   288
         Width           =   1452
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "APPROVED BY:"
         Height          =   252
         Left            =   5160
         TabIndex        =   36
         Tag             =   "xyz"
         Top             =   612
         Width           =   1452
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "NOTED BY:"
         Height          =   252
         Left            =   5160
         TabIndex        =   35
         Tag             =   "xyz"
         Top             =   240
         Width           =   1332
      End
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   15600
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "txtHoldGrayBackColor"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   15600
      TabIndex        =   24
      Text            =   "HoldBackColor"
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   372
      Index           =   4
      Left            =   5520
      TabIndex        =   23
      Tag             =   "xyz"
      Top             =   120
      Width           =   1188
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "E&xit Form"
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
      Image           =   "frmRFPOther.frx":5B84
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Tag             =   "xyz"
      Top             =   120
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   661
      Caption         =   "&New"
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
      Image           =   "frmRFPOther.frx":611E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   372
      Index           =   1
      Left            =   1320
      TabIndex        =   20
      Tag             =   "xyz"
      Top             =   120
      Width           =   1080
      _ExtentX        =   1905
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
      Image           =   "frmRFPOther.frx":66B8
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   372
      Index           =   2
      Left            =   2520
      TabIndex        =   21
      Tag             =   "xyz"
      Top             =   120
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   661
      Caption         =   "&Print Preview"
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
      Image           =   "frmRFPOther.frx":6C52
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   372
      Index           =   3
      Left            =   3960
      TabIndex        =   22
      Tag             =   "xyz"
      Top             =   120
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "View &Summary"
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
      Image           =   "frmRFPOther.frx":C444
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   372
      Index           =   6
      Left            =   15600
      TabIndex        =   19
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
      Image           =   "frmRFPOther.frx":C9DE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   372
      Index           =   5
      Left            =   15600
      TabIndex        =   18
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
      Image           =   "frmRFPOther.frx":CF78
      cBack           =   -2147483633
   End
   Begin MSMask.MaskEdBox txtCheckDate 
      Height          =   312
      Left            =   4560
      TabIndex        =   13
      Tag             =   "XYZ"
      Top             =   8760
      Width           =   1932
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   12632256
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Frame framHoldTop 
      Height          =   1332
      Left            =   120
      TabIndex        =   26
      Tag             =   "xyz"
      Top             =   480
      Width           =   14652
      Begin VB.ComboBox cboCompany 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         ItemData        =   "frmRFPOther.frx":D512
         Left            =   5520
         List            =   "frmRFPOther.frx":D51C
         TabIndex        =   86
         Tag             =   "XYZ"
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox cboRFPType 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   5640
         TabIndex        =   83
         Tag             =   "XYZ"
         Top             =   960
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtPayee 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   80
         Tag             =   "XYZ"
         Top             =   240
         Width           =   5775
      End
      Begin VB.ComboBox cboStatutory 
         BackColor       =   &H00C0C0C0&
         Height          =   288
         ItemData        =   "frmRFPOther.frx":D55D
         Left            =   11160
         List            =   "frmRFPOther.frx":D55F
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "XYZ"
         Top             =   960
         Width           =   1932
      End
      Begin VB.ComboBox cboBillTo 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   2
         Tag             =   "XYZ"
         Top             =   600
         Width           =   5775
      End
      Begin VB.ComboBox cboSupName2 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   1
         Tag             =   "XYZ"
         Top             =   240
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.TextBox txtRFPNo 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   11160
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Tag             =   "XYZ"
         ToolTipText     =   "Request for Payment Number"
         Top             =   240
         Width           =   1935
      End
      Begin MSMask.MaskEdBox txtRFPDate 
         Height          =   312
         Left            =   11160
         TabIndex        =   5
         Tag             =   "XYZ"
         ToolTipText     =   "Request for Payment Date"
         Top             =   600
         Width           =   1932
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12632256
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDueDate 
         Height          =   312
         Left            =   2040
         TabIndex        =   3
         Tag             =   "XYZ"
         Top             =   960
         Width           =   1212
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12632256
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label27 
         Caption         =   "COMPANY:"
         Height          =   255
         Left            =   4680
         TabIndex        =   87
         Tag             =   "XYZ"
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label23 
         Caption         =   "RFP TYPE:"
         Height          =   255
         Left            =   4800
         TabIndex        =   84
         Tag             =   "XYZ"
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "PAYEE:"
         Height          =   255
         Left            =   360
         TabIndex        =   81
         Tag             =   "XYZ"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "DUE DATE:"
         Height          =   372
         Left            =   360
         TabIndex        =   75
         Tag             =   "xyz"
         Top             =   972
         Width           =   1452
      End
      Begin VB.Label Label22 
         Caption         =   "MM/DD/YYYY"
         Height          =   252
         Left            =   3360
         TabIndex        =   74
         Tag             =   "XYZ"
         Top             =   960
         Width           =   1092
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "STATUTORY:"
         Height          =   252
         Left            =   9960
         TabIndex        =   32
         Tag             =   "XYZ"
         Top             =   960
         Width           =   1212
      End
      Begin VB.Label Label5 
         Caption         =   "MM/DD/YYYY"
         Height          =   252
         Left            =   13150
         TabIndex        =   31
         Tag             =   "XYZ"
         Top             =   600
         Width           =   1092
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "RFP DATE:"
         Height          =   252
         Left            =   9960
         TabIndex        =   30
         Tag             =   "XYZ"
         Top             =   600
         Width           =   852
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "RFP NO. :"
         Height          =   252
         Left            =   9960
         TabIndex        =   29
         Tag             =   "XYZ"
         Top             =   240
         Width           =   1212
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "BILL TO:"
         Height          =   252
         Left            =   360
         TabIndex        =   28
         Tag             =   "XYZ"
         Top             =   600
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER'S NAME:"
         Height          =   252
         Left            =   360
         TabIndex        =   27
         Tag             =   "XYZ"
         Top             =   240
         Visible         =   0   'False
         Width           =   1812
      End
   End
   Begin lvButton.lvButtons_H cmdSync 
      Height          =   372
      Left            =   11160
      TabIndex        =   73
      Tag             =   "XYZ"
      Top             =   8520
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&SYNC ACCT. TITLE"
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
      Image           =   "frmRFPOther.frx":D561
      cBack           =   -2147483633
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "REMARKS:"
      Height          =   192
      Left            =   120
      TabIndex        =   66
      Tag             =   "XYZ"
      Top             =   8400
      Width           =   828
   End
   Begin VB.Label Label9 
      Caption         =   "MM/DD/YYYY"
      Height          =   252
      Left            =   6580
      TabIndex        =   56
      Tag             =   "XYZ"
      Top             =   8760
      Width           =   1092
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "CHECK NO.:"
      Height          =   252
      Left            =   120
      TabIndex        =   50
      Tag             =   "xyz"
      Top             =   8760
      Width           =   972
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "CHECK  DATE:"
      Height          =   372
      Left            =   3480
      TabIndex        =   49
      Tag             =   "xyz"
      Top             =   8772
      Width           =   1452
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   7920
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   15165
      m_MDI_Height    =   10830
   End
   Begin VB.Menu mnuPopUpFormat 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Row"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuDel 
         Caption         =   "Delete Row"
         Shortcut        =   ^D
      End
   End
End
Attribute VB_Name = "frmRFPOther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpRs As New ADODB.Recordset
Dim formAction As String
Dim didRecordEntry As Boolean
Dim Qry As String
Dim ans As VbMsgBoxResult
Dim level As String
Dim message As String
Dim RFPBranchCode As String

Dim DtlsCol As Integer
Dim DtlsRow As Integer, i As Integer, j As Integer
Dim flexGridObject As String
Dim checkdate As String
Public RFPID As Long

Dim AllArrBillTo() As String
Dim arrBillTo() As String

Dim AllArrSupplier() As String
Dim arrSupplier() As String
Dim arrBRANCH() As String

Dim supplierid As Long

Dim SearchcboSup As Boolean
Dim SearchcboBillTo As Boolean

Dim TransID As Long
Dim pTransID As Long
Dim arrDesc() As String
Dim arrAcctTitles() As String
Dim arrAcctCode() As String
Dim SearchcboDesc As Boolean
Dim SearchcboAcctTitles As Boolean
Dim SearchcboAcctCode As Boolean
Dim SearchcboParticular As Boolean

Public PrintCopy As Integer

Dim AcctCode As String
Dim AcctTitle As String
Dim AcctType As String

Dim StatSrch As Boolean
Dim strPartial As String, strTotal As String
Dim PrtBln As Boolean, PrtSrch As Boolean

'Sub DisableRighClick(DataEntry As Boolean)
'On Error GoTo er
'
'    If DataEntry = True Then
'        ' Set the control's new WindowProc.
'        OldWindowProc = SetWindowLong(txtRFPNo.hwnd, GWL_WNDPROC, AddressOf NoPopupWindowProc)
'        OldWindowProc = SetWindowLong(txtFlxBrchDpt.hwnd, GWL_WNDPROC, AddressOf NoPopupWindowProc)
'        OldWindowProc = SetWindowLong(txtFlxRefNo.hwnd, GWL_WNDPROC, AddressOf NoPopupWindowProc)
'        OldWindowProc = SetWindowLong(txtFlxParticular.hwnd, GWL_WNDPROC, AddressOf NoPopupWindowProc)
'        OldWindowProc = SetWindowLong(txtFlxValueDeduct.hwnd, GWL_WNDPROC, AddressOf NoPopupWindowProc)
'        OldWindowProc = SetWindowLong(txtRemarks.hwnd, GWL_WNDPROC, AddressOf NoPopupWindowProc)
'        OldWindowProc = SetWindowLong(txtCheckNo.hwnd, GWL_WNDPROC, AddressOf NoPopupWindowProc)
'        OldWindowProc = SetWindowLong(txtChckedBy.hwnd, GWL_WNDPROC, AddressOf NoPopupWindowProc)
'        OldWindowProc = SetWindowLong(txtNotedBy.hwnd, GWL_WNDPROC, AddressOf NoPopupWindowProc)
'        OldWindowProc = SetWindowLong(txtApprvdBy.hwnd, GWL_WNDPROC, AddressOf NoPopupWindowProc)
'    Else
'        ' Restore the original WindowProc.
'        SetWindowLong txtRFPNo.hwnd, GWL_WNDPROC, OldWindowProc
'        SetWindowLong txtFlxBrchDpt.hwnd, GWL_WNDPROC, OldWindowProc
'        SetWindowLong txtFlxRefNo.hwnd, GWL_WNDPROC, OldWindowProc
'        SetWindowLong txtFlxParticular.hwnd, GWL_WNDPROC, OldWindowProc
'        SetWindowLong txtFlxValueDeduct.hwnd, GWL_WNDPROC, OldWindowProc
'        SetWindowLong txtRemarks.hwnd, GWL_WNDPROC, OldWindowProc
'        SetWindowLong txtCheckNo.hwnd, GWL_WNDPROC, OldWindowProc
'        SetWindowLong txtChckedBy.hwnd, GWL_WNDPROC, OldWindowProc
'        SetWindowLong txtNotedBy.hwnd, GWL_WNDPROC, OldWindowProc
'        SetWindowLong txtApprvdBy.hwnd, GWL_WNDPROC, OldWindowProc
'    End If
'
'Exit Sub
'er:
'    Call modObjTools.ErrorMessage(Err.Description, "Error->DisableRighClick()")
'End Sub

Public Sub ShowRecord()
On Error GoTo er
    Dim i As Integer
    
    Call modAPI.sp_load_rfpother_info(RFPID, tmpRs)
    If tmpRs.State = 0 Then Exit Sub
    
    If tmpRs.RecordCount <= 0 Then
        tmpRs.Close
        Exit Sub
    End If
    
    tmpRs.MoveFirst
    'SUPPLIERNAME, ADDRESS, SOLDTO, RFPNO, RFPDATE, POTYPE, REMARKS, CHECKNO, CHECKAMNT, PREPBY, REVBY, APPBY, STATUS, PrintCopy
            txtPayee.Text = tmpRs!SUPPLIERNAME
            'txtSupAddress.Text = tmpRs!Address
            cboBillTo.Text = tmpRs!SOLDTO
            RFPID = tmpRs!RFPID
            txtRFPNo.Text = tmpRs!rfpno
            txtRFPDate.Text = tmpRs!RFPDATE
            cboCompany.Text = IIf(IsNull(tmpRs!COMPANY), "LYR MOTORPARTS (RETAIL)", tmpRs!COMPANY)
            
            If tmpRs!DueDate <> "" Then
                txtDueDate.Text = tmpRs!DueDate
            End If
            
            Call modObjTools.changeListIndex(cboStatutory, tmpRs!POTYPE)
            txtRemarks.Text = tmpRs!Remarks
            
            If tmpRs!CheckNo <> "" Then
                txtCheckNo.Text = tmpRs!CheckNo
            End If
            
            If tmpRs!checkdate <> "" Then txtCheckDate.Text = Format(tmpRs!checkdate, "mm/dd/YYYY")
            If tmpRs!docdaterec <> "" Then txtDocDateRec.Text = Format(tmpRs!docdaterec, "mm/dd/YYYY")
            
            txtPrepBy.Text = tmpRs!PREPBY
            txtChckedBy.Text = tmpRs!REVBY
            txtNotedBy.Text = tmpRs!NOTEDBY
            txtApprvdBy.Text = tmpRs!APPBY
            
            cboStatus.Clear
            cboStatus.AddItem "FOR CHECKING"
            cboStatus.AddItem "FOR APPROVAL"
            cboStatus.AddItem "APPROVED WITH CHECK NO."
            cboStatus.AddItem "CANCELLED"
            Call modObjTools.changeListIndex(cboStatus, tmpRs!Status)
            
            PrintCopy = tmpRs!PrintCopy
            
        While Not tmpRs.EOF
    'SHIPPEDTO, INVNO, PARTICULARS, INVAMNT
          If flxRFP.TextMatrix(1, 1) = "" Then
                flxRFP.TextMatrix(1, 1) = tmpRs!SHIPPEDTO
                flxRFP.TextMatrix(1, 2) = tmpRs!InvNo
                flxRFP.TextMatrix(1, 3) = tmpRs!Particulars
                flxRFP.TextMatrix(1, 4) = Format(tmpRs!INVAMNT, "#,##0.00")

            Else
                flxRFP.AddItem ""
                flxRFP.TextMatrix(flxRFP.Rows - 1, 1) = tmpRs!SHIPPEDTO
                flxRFP.TextMatrix(flxRFP.Rows - 1, 2) = tmpRs!InvNo
                flxRFP.TextMatrix(flxRFP.Rows - 1, 3) = tmpRs!Particulars
                flxRFP.TextMatrix(flxRFP.Rows - 1, 4) = Format(tmpRs!INVAMNT, "#,##0.00")

            End If
            
        tmpRs.MoveNext
    Wend
    
    
    If tmpRs.State <> 0 Then tmpRs.Close

    DoEvents
    Call ShowDeduction(RFPID)
    DoEvents
    Call ShowAcctTitles(RFPID)
    DoEvents
    Call EnabledSelectedButton
    DoEvents
    Call ComputeTotalNetPay
    DoEvents
    Call ComputeTotAddLess
    
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
            
        While Not tmpRs.EOF

            If flxDeduct.TextMatrix(1, 1) = "" Then
                flxDeduct.TextMatrix(1, 1) = tmpRs!DTYPE
                flxDeduct.TextMatrix(1, 2) = tmpRs!Type
                flxDeduct.TextMatrix(1, 3) = tmpRs!Description
                flxDeduct.TextMatrix(1, 4) = Format(tmpRs!REBDUCT, "#,##0.00")

  
            Else
                flxDeduct.AddItem _
                vbTab & tmpRs!DTYPE & _
                vbTab & tmpRs!Type & _
                vbTab & tmpRs!Description & _
                vbTab & Format(tmpRs!REBDUCT, "#,##0.00")

            End If
            
            
        tmpRs.MoveNext
    Wend
    
    
    If tmpRs.State <> 0 Then tmpRs.Close
    
    Call ComputeTotAddLess
    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->ShowDeduction()")
End Sub


Public Sub ShowAcctTitles(ByVal pRFPID As Long)
On Error GoTo er
    Dim i As Integer
    
    Call modAPI.sp_load_acct_title(pRFPID, tmpRs)
    If tmpRs.State = 0 Then Exit Sub
    
    If tmpRs.RecordCount <= 0 Then
        tmpRs.Close
        Exit Sub
    End If
'TransNAme, code, Description, Type, Amount, TotDebit, TotCredit
    tmpRs.MoveFirst
        
        cboTrans.Text = tmpRs!TransNAme
            
        While Not tmpRs.EOF

          If flxVoucher.TextMatrix(1, 2) = "" Then
                flxVoucher.TextMatrix(1, 1) = IIf(IsNull(tmpRs!branchName), "", tmpRs!branchName)
                flxVoucher.TextMatrix(1, 2) = tmpRs!code
                flxVoucher.TextMatrix(1, 3) = tmpRs!Description
                flxVoucher.TextMatrix(1, 4) = tmpRs!Type
                flxVoucher.TextMatrix(1, 5) = Format(tmpRs!Amount, "#,##0.00")
            Else
                flxVoucher.AddItem _
                vbTab & IIf(IsNull(tmpRs!branchName), "", tmpRs!branchName) & _
                vbTab & tmpRs!code & _
                vbTab & tmpRs!Description & _
                vbTab & tmpRs!Type & _
                vbTab & Format(tmpRs!Amount, "#,##0.00")
            End If
            
        tmpRs.MoveNext
    Wend
    
    
    If tmpRs.State <> 0 Then tmpRs.Close
    
    Call ComputeTotalDebitCredit
    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->ShowAcctTitles()")
End Sub

Private Sub cboCompany_Change()
On Error GoTo er
    Dim i As Integer, j As Integer
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
    If SearchcboBillTo Then
        SearchcboBillTo = False
        Exit Sub
    End If

    With cboCompany
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters

        If i <> -1 And i <= .ListCount Then
            strTotal = .List(i)
    
            DoEvents
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
               SearchcboBillTo = True
               .SelText = Right$(strTotal, j)
               .SelStart = Len(strPartial)
               .SelLength = j
            End If

        End If

    End With
        
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error ->cboCompany_Change()")
End Sub

Private Sub cboCompany_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Or KeyCode = 46 Then
        SearchcboBillTo = True
    End If
End Sub

Private Sub cboCompany_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboCompany, KeyAscii, ValidChars)
    If KeyAscii = 13 Then cboStatutory.SetFocus
End Sub

Private Sub cboFlxBranch_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboFlxBranch
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboFlxBranch-Change()")
End Sub

Private Sub cboFlxBranch_Click()
 Call cboFlxBranch_LostFocus
End Sub

Private Sub cboFlxBranch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboFlxBranch_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cboFlxBranch_LostFocus
End Sub

Private Sub cboFlxBranch_LostFocus()
    If cboFlxBranch.ListIndex < 0 Then
        flxVoucher.TextMatrix(DtlsRow, 1) = ""
        cboFlxBranch.Visible = False
        Exit Sub
    End If
    flxVoucher.TextMatrix(DtlsRow, 1) = arrBRANCH(cboFlxBranch.ListIndex, 1)
    flxVoucher.TextMatrix(DtlsRow, 6) = arrBRANCH(cboFlxBranch.ListIndex, 0)
    cboFlxBranch.Visible = False
    flxVoucher.SetFocus
    flxVoucher.col = 2
End Sub

Private Sub cboFlxMainBranch_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboFlxMainBranch
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboFlxMainBranch-Change()")
End Sub

Private Sub cboFlxMainBranch_Click()
 Call cboFlxMainBranch_LostFocus
End Sub

Private Sub cboFlxMainBranch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboFlxMainBranch_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cboFlxMainBranch_LostFocus
End Sub

Private Sub cboFlxMainBranch_LostFocus()
    If cboFlxMainBranch.ListIndex < 0 Then
        flxRFP.TextMatrix(DtlsRow, 1) = ""
        cboFlxMainBranch.Visible = False
        Exit Sub
    End If
    RFPBranchCode = ""
    flxRFP.TextMatrix(DtlsRow, 1) = arrBRANCH(cboFlxMainBranch.ListIndex, 1)
    RFPBranchCode = arrBRANCH(cboFlxMainBranch.ListIndex, 0)
    cboFlxMainBranch.Visible = False
    flxRFP.SetFocus
    flxRFP.col = 2
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

Private Sub cboFlxParticular_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboFlxParticular, KeyAscii, ValidSQLChars)
    If KeyAscii = 13 Then
        Dim LessValue As Double
        
        If Val(lblTotPOAmnt.Caption) = 0 Then
            cboFlxParticular.Visible = False
            Exit Sub
        End If
        
        If flxDeduct.TextMatrix(DtlsRow, 3) = "EXPANDED WITH HOLDING TAX 1%" Then
        
            LessValue = (Val(Replace(lblTotPOAmnt.Caption, ",", "")) / 1.12) * 0.01
            flxDeduct.TextMatrix(DtlsRow, 4) = Format(LessValue, "#,##0.00")
            
        ElseIf flxDeduct.TextMatrix(DtlsRow, 3) = "EXPANDED WITH HOLDING TAX 2%" Then
        
            LessValue = (Val(Replace(lblTotPOAmnt.Caption, ",", "")) / 1.12) * 0.02
            flxDeduct.TextMatrix(DtlsRow, 4) = Format(LessValue, "#,##0.00")
        ElseIf flxDeduct.TextMatrix(DtlsRow, 3) = "EXPANDED WITH HOLDING TAX 5%" Then
        
            LessValue = (Val(Replace(lblTotPOAmnt.Caption, ",", "")) / 1.12) * 0.05
            flxDeduct.TextMatrix(DtlsRow, 4) = Format(LessValue, "#,##0.00")
        End If
        
        flxDeduct.TextMatrix(DtlsRow, 3) = cboFlxParticular.Text
        cboFlxParticular.Visible = False
        flxDeduct.SetFocus
        Call ComputeTotAddLess
        Call ComputeTotalNetPay
        
    End If
End Sub

Private Sub cboFlxParticular_LostFocus()
    Call cboFlxParticular_KeyPress(13)
End Sub

Private Sub cboFlxType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxDeduct.TextMatrix(DtlsRow, 2) = cboFlxType.Text
        cboFlxType.Visible = False
        flxDeduct.SetFocus
    End If
End Sub

Private Sub cboFlxType_LostFocus()
    Call cboFlxType_KeyPress(13)
End Sub


Private Sub cboTrans_LostFocus()
    If pTransID = TransID Then Exit Sub
    Call GetAccountTitles(TransID)

    flxVoucher.col = 4
    flxVoucher.row = 1
    flxVoucher.SetFocus
End Sub

Private Sub cmdHistory_Click()
Dim i As Integer
    'ssTabRBhist.Visible = True
    'txtTransName.Text = flxRBhist.TextMatrix(ActRow, 2)
    'txtReference.Text = flxRBhist.TextMatrix(ActRow, 3)
    With ssTabRBhist
        .Top = (Me.Height - .Height) \ 2
        .Left = (Me.Width - .Width) \ 2
        .Visible = True
        .ZOrder
    End With
    
    With flxRBTransHist
        '.Clear
        .Rows = 2
        .Cols = 5
        .WordWrap = True
        .AllowUserResizing = flexResizeBoth
        '               0     1        2      3           4
        .FormatString = " |^DATE |<TRANSACTION NAME |^STATUS |<RESPONSIBLE"
        
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.21
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.3
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.15
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.3
    End With
    
    Qry = "call sp_mrcwarehouse.sp_transaction_hist('" & txtRFPNo.Text & "')"
    If rsAnyUse.State = adStateOpen Then rsAnyUse.Close
    rsAnyUse.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    
    With rsAnyUse
        If .RecordCount = 0 Then   'No records to show
            flxRBTransHist.Clear
            flxRBTransHist.Rows = 2
            flxRBTransHist.Cols = 5
            'Call RXResizer1_ResizeComplete
            Exit Sub
        Else
            .MoveFirst
            'Call RXResizer1_ResizeComplete
            flxRBTransHist.Rows = .RecordCount + 1
            
            For i = 1 To flxRBTransHist.Rows - 1
                flxRBTransHist.TextMatrix(i, 1) = IIf(IsNull(!Date), "", Format(!Date, "mm/dd/yyyy hh:mm:ss AM/PM"))
                flxRBTransHist.TextMatrix(i, 2) = IIf(IsNull(!TransNAme), "", UCase(!TransNAme))
                flxRBTransHist.TextMatrix(i, 3) = IIf(IsNull(!Status), "", !Status)
                flxRBTransHist.TextMatrix(i, 4) = IIf(IsNull(!RESPONSIBLE), "", !RESPONSIBLE)
                .MoveNext
            Next i
        End If
        .Close
    End With

'    frmTransHist.Trans = "RFPOTHER"
'    frmTransHist.RefID = RFPID
'    frmTransHist.Show vbModal
End Sub

Private Sub cmdSync_Click()
On Error GoTo er

    ans = MsgBox("Synchronize account titles?", vbYesNo + vbQuestion, Me.Caption)
    If ans = vbNo Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
        
    DoEvents
    Screen.MousePointer = vbHourglass
    tmrSync.Enabled = True
    
    DoEvents
    Call Load(frm_OnProcess2)
    frm_OnProcess2.Show vbModal
    
Exit Sub
er:
    tmrSync.Enabled = False
    Screen.MousePointer = vbDefault
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdSync_Click()")
End Sub

Private Sub flxDeduct_Click()
    Dim i As Integer
    
    If didRecordEntry = False Then Exit Sub
    
    DtlsRow = flxDeduct.row
    DtlsCol = flxDeduct.col
    flexGridObject = "flxDeduct"
    
    If DtlsCol = 0 Then 'Check Icon
    
        With flxDeduct
            If .col >= 0 And .row > 0 Then
                If .TextMatrix(.row, 0) = "" Then
                    .FillStyle = flexFillSingle
                    .CellFontName = "Webdings"
                    .TextMatrix(.row, 0) = "a"
                Else
                    .FillStyle = flexFillSingle
                    .TextMatrix(.row, 0) = ""
                    .Highlight = flexHighlightNever
                End If
            End If
        End With
        
    End If

End Sub

Private Sub flxDeduct_DblClick()
On Error GoTo er

    If didRecordEntry = False Then Exit Sub
   
    DtlsCol = flxDeduct.col
    DtlsRow = flxDeduct.row
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

Private Sub flxDeduct_GotFocus()
    flexGridObject = "flxDeduct"
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

Private Sub flxRFP_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If didRecordEntry = False Then Exit Sub
    If Button = vbRightButton Then
        PopupMenu mnuPopUpFormat, vbPopupMenuRightButton
        flexGridObject = "flxRFP"
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
   ' Call DisableRighClick(False)
End Sub

Private Sub lblClose_Click()
ssTabRBhist.Visible = False
End Sub

Private Sub mnuAdd_Click()
On Error GoTo er
    
    If didRecordEntry = False Then Exit Sub
    
    If flexGridObject = "flxRFP" Then
        flxRFP.AddItem ""
        flxRFP.col = 1
        flxRFP.SetFocus
        
    ElseIf flexGridObject = "flxDeduct" Then
        flxDeduct.AddItem ""
        flxDeduct.col = 1
        flxDeduct.SetFocus
    
    ElseIf flexGridObject = "flxVoucher" Then
        flxVoucher.Rows = flxVoucher.Rows + 1
    
        For i = flxVoucher.Rows - 1 To flxVoucher.row Step -1
            If i = flxVoucher.row Then
                For j = 0 To flxVoucher.Cols - 1
                flxVoucher.TextMatrix(i + 1, j) = ""

                Next j
            Else
                For j = 0 To flxVoucher.Cols - 1
                flxVoucher.TextMatrix(i, j) = flxVoucher.TextMatrix(i - 1, j)
                Next j
            End If
        Next i

'        flxVoucher.AddItem "" & _
'        vbTab & flxVoucher.TextMatrix(flxVoucher.Rows - 1, 1) & _
'        vbTab & flxVoucher.TextMatrix(flxVoucher.Rows - 1, 2) & _
'        vbTab & flxVoucher.TextMatrix(flxVoucher.Rows - 1, 3) & _
'        vbTab & flxVoucher.TextMatrix(flxVoucher.Rows - 1, 4)
'
'        If InStr(flxVoucher.TextMatrix(flxVoucher.Rows - 2, 2), "CASH IN BANK") Then
'            flxVoucher.TextMatrix(flxVoucher.Rows - 2, 1) = ""
'            flxVoucher.TextMatrix(flxVoucher.Rows - 2, 2) = ""
'            flxVoucher.TextMatrix(flxVoucher.Rows - 2, 3) = ""
'            flxVoucher.TextMatrix(flxVoucher.Rows - 2, 4) = ""
'        End If
    
        DoEvents
        flxVoucher.SetFocus
        flxVoucher.col = 2
        flxVoucher.row = flxVoucher.Rows - 2
    End If

Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->mnuADD_Click()")
End Sub

Private Sub mnuDel_Click()
On Error GoTo er
    Dim i As Integer
    
    If didRecordEntry = False Then Exit Sub
    
    ans = MsgBox("Delete row?", vbYesNo + vbQuestion, Me.Caption)
    If ans = vbNo Then Exit Sub
    
    If flexGridObject = "flxRFP" Then
        For i = 1 To flxRFP.Rows - 1
            If i <= flxRFP.Rows - 1 Then
                If flxRFP.TextMatrix(i, 0) = "a" Then
                    If flxRFP.Rows = 2 Then
                        flxRFP.Rows = 1
                        flxRFP.Rows = 2
                        
                    Else
                        flxRFP.RemoveItem (i)
                        i = i - 1
                    End If
                End If
            End If
        Next i
        
    ElseIf flexGridObject = "flxDeduct" Then
        For i = 1 To flxDeduct.Rows - 1
            If i <= flxDeduct.Rows - 1 Then
                If flxDeduct.TextMatrix(i, 0) = "a" Then
                    If flxDeduct.Rows = 2 Then
                        flxDeduct.Rows = 1
                        flxDeduct.Rows = 2
                        
                    Else
                        flxDeduct.RemoveItem (i)
                        i = i - 1
                    End If
                End If
            End If
        Next i
        Call ComputeTotAddLess
        
    ElseIf flexGridObject = "flxVoucher" Then
        For i = 1 To flxVoucher.Rows - 1
            If i <= flxVoucher.Rows - 1 Then
                If flxVoucher.TextMatrix(i, 0) = "a" Then
                    If flxVoucher.Rows = 2 Then
                        flxVoucher.Rows = 1
                        flxVoucher.Rows = 2
                    
                    Else
                        flxVoucher.RemoveItem (i)
                        i = i - 1
                    End If
                End If
            End If
        Next i
        Call ComputeTotalDebitCredit
    End If
    
    Call ComputeTotalNetPay
    
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error ->mnuDel_Click()"
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
    
    lblTotAdd.Caption = Format(TotalAdd, "#,##0.00")
    lblTotLess.Caption = Format(TotalLess, "#,##0.00")
    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Erro->ComputeTotAddLess()")
End Sub


Function ValidateEntry() As String
    Dim vMsg As String, Qry As String
    Dim i As Integer, Cntr As Integer
    
    vMsg = ""
    
    If Replace(txtPayee.Text, " ", "") = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field supplier name."
    'If supplierid = 0 Then vMsg = vMsg & vbCrLf & "Supplier name doesn't exist!"
    
    If Replace(cboBillTo.Text, " ", "") = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field sold to."

    If Not IsDate(txtRFPDate.Text) Then vMsg = vMsg & vbCrLf & "Invalid RFP date format."
    
    If Val(txtCheckDate.Text) > 0 Then
        If Not IsDate(txtCheckDate.Text) Then vMsg = vMsg & vbCrLf & "Invalid Check date format."
    End If
    
    If Val(txtDueDate.Text) > 0 Then
        If Not IsDate(txtDueDate.Text) Then vMsg = vMsg & vbCrLf & "Invalid Due date format."
    End If
    
    If modObjTools.getListIndex(cboBillTo, cboBillTo.Text) = 0 Then
         vMsg = vMsg & vbCrLf & "Bill to name doesn't exist!"
    End If
    
    For i = 1 To flxRFP.Rows - 1
        'If flxRFP.TextMatrix(i, 1) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Branch/Department at row " & i
        If flxRFP.TextMatrix(i, 2) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Ref. no. at row " & i
        'If flxRFP.TextMatrix(i, 3) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Particulars at row " & i
        If flxRFP.TextMatrix(i, 4) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Amount at row " & i
    Next i
    
    For i = 1 To flxVoucher.Rows - 1
        If flxVoucher.TextMatrix(i, 2) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Account Code to at row " & i
        If flxVoucher.TextMatrix(i, 3) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Account Title to at row " & i
        If flxVoucher.TextMatrix(i, 4) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Account Type to at row " & i
    Next i
    
    If Replace(cboTrans.Text, " ", "") = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Transaction."
    If Replace(txtPrepBy.Text, " ", "") = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field prepared by name."
    If Replace(txtApprvdBy.Text, " ", "") = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field approved by name."
    'If Replace(txtRemarks.Text, " ", "") = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on remarks."
    If Val(TransID) <= 0 Then vMsg = vMsg & vbCrLf & "Transaction not found in the List."
    
    If Val(lblTotCredit.Caption) <> Val(lblTotDebit) Then vMsg = vMsg & vbCrLf & "Debit and Credit was not equal."
    If Val(txtDocDateRec.Text) <= 0 Then
        If Not IsDate(txtDocDateRec.Text) Then vMsg = vMsg & vbCrLf & "Invalid Document date receive format."
    End If
    
    ValidateEntry = vMsg
    
End Function


Private Sub cbobillto_Change()
On Error GoTo er
    Dim i As Integer, j As Integer
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
    If SearchcboBillTo Then
        SearchcboBillTo = False
        Exit Sub
    End If

    With cboBillTo
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters

        If i <> -1 And i <= .ListCount Then
            strTotal = .List(i)
    
            DoEvents
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
               SearchcboBillTo = True
               .SelText = Right$(strTotal, j)
               .SelStart = Len(strPartial)
               .SelLength = j
            End If

        End If

    End With
        
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error ->cboSoldTo_Change()")
End Sub

Private Sub cbobillto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Or KeyCode = 46 Then
        SearchcboBillTo = True
    End If
End Sub


Private Sub cbobillto_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboBillTo, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtDueDate.SetFocus
End Sub

Private Sub cboFlxAcctCode_Change()
On Error GoTo er
    Dim i As Integer, j As Integer
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
    
    If SearchcboAcctCode Then
        SearchcboAcctCode = False
        Exit Sub
    End If
    
    With cboFlxAcctCode
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters

        If i <> -1 And i <= .ListCount Then
            strTotal = .List(i)
            
            'Acct_Code,Acct_Titles,`type`
            AcctCode = arrAcctCode(i, 0)
            AcctTitle = arrAcctCode(i, 1)
            AcctType = arrAcctCode(i, 2)
    
            DoEvents
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
               SearchcboAcctCode = True
               .SelText = Right$(strTotal, j)
               .SelStart = Len(strPartial)
               .SelLength = j
            End If

        End If

    End With
        
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error ->cboFlxAcctCode_Change()")
End Sub

Private Sub cboFlxAcctCode_Click()
    'Acct_Code,Acct_Titles,`type`
    AcctCode = arrAcctCode(cboFlxAcctCode.ListIndex, 0)
    AcctTitle = arrAcctCode(cboFlxAcctCode.ListIndex, 1)
    AcctType = arrAcctCode(cboFlxAcctCode.ListIndex, 2)
End Sub

Private Sub cboFlxAcctCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Or KeyCode = 46 Then
        SearchcboAcctCode = True
    End If
      
End Sub

Private Sub cboFlxAcctCode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cboFlxAcctCode_LostFocus
End Sub

Private Sub cboFlxAcctCode_LostFocus()
    flxVoucher.TextMatrix(flxVoucher.row, 2) = AcctCode
    flxVoucher.TextMatrix(flxVoucher.row, 3) = AcctTitle
    flxVoucher.TextMatrix(flxVoucher.row, 4) = AcctType
    cboFlxAcctCode.Visible = False
    flxVoucher.SetFocus
End Sub

Private Sub cboFlxAcctTitles_Change()
On Error GoTo er
    Dim i As Integer, j As Integer
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
    
    If SearchcboAcctTitles Then
        SearchcboAcctTitles = False
        Exit Sub
    End If
    
    With cboFlxAcctTitles
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters

        If i <> -1 And i <= .ListCount Then
            strTotal = .List(i)
            
            'Acct_Code,Acct_Titles,`type`,seqno
            AcctCode = arrAcctTitles(i, 0)
            AcctTitle = arrAcctTitles(i, 1)
            AcctType = arrAcctTitles(i, 2)
            
    
            DoEvents
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
               SearchcboAcctTitles = True
               .SelText = Right$(strTotal, j)
               .SelStart = Len(strPartial)
               .SelLength = j
            End If

        End If

    End With
        
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error ->cboFlxAcctTitles_Change()")
End Sub

Sub ComputeTotalDebitCredit()
On Error GoTo er
    Dim i As Integer
    Dim Debit As Double
    Dim Credit As Double
    
    For i = 1 To flxVoucher.Rows - 1
        
        If flxVoucher.TextMatrix(i, 4) = "DEBIT" Then
            Debit = Debit + Val(Replace(flxVoucher.TextMatrix(i, 5), ",", ""))
        Else
            Credit = Credit + Val(Replace(flxVoucher.TextMatrix(i, 5), ",", ""))
        End If
    Next
    
    lblTotDebit.Caption = Format(Debit, "#,##0.00")
    lblTotCredit.Caption = Format(Credit, "#,##0.00")
    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->ComputeTotalDebitCredit()")
End Sub


Private Sub cboFlxAcctTitles_Click()
    'Acct_Code,Acct_Titles,`type`
    AcctCode = arrAcctTitles(cboFlxAcctTitles.ListIndex, 0)
    AcctTitle = arrAcctTitles(cboFlxAcctTitles.ListIndex, 1)
    AcctType = arrAcctTitles(cboFlxAcctTitles.ListIndex, 2)
End Sub

Private Sub cboFlxAcctTitles_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Or KeyCode = 46 Then
        SearchcboAcctTitles = True
    End If
End Sub

Private Sub cboFlxAcctTitles_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cboFlxAcctTitles_LostFocus
End Sub

Private Sub cboFlxAcctTitles_LostFocus()
    flxVoucher.TextMatrix(flxVoucher.row, 2) = AcctCode
    flxVoucher.TextMatrix(flxVoucher.row, 3) = AcctTitle
    flxVoucher.TextMatrix(flxVoucher.row, 4) = AcctType
    cboFlxAcctTitles.Visible = False
    flxVoucher.col = 4
    flxVoucher.SetFocus
End Sub



Private Sub cboFlxAcctType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxVoucher.TextMatrix(flxVoucher.row, 3) = cboFlxAcctType.Text
        cboFlxAcctType.Visible = False
        flxVoucher.SetFocus
    End If
End Sub

Private Sub cboFlxAcctType_LostFocus()
    flxVoucher.TextMatrix(flxVoucher.row, 4) = cboFlxAcctType.Text
    cboFlxAcctType.Visible = False
    flxVoucher.SetFocus
End Sub

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboStatus, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        If cmdTrans(5).Visible = True Then
            cmdTrans(5).SetFocus
        Else
            txtPayee.SetFocus
        End If
    End If
End Sub

Private Sub tmrGetRecord_Timer()
On Error GoTo er
    If tmrGetRecord.Enabled = False Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    tmrGetRecord.Enabled = False

        DoEvents
        Call ShowRecord
        
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

Private Sub tmrSync_Timer()
On Error GoTo er
Dim level As String
Dim message As String

    tmrSync.Enabled = False
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    

    DoEvents
        Call modAPI.sp_sync_acct_titles(level, message)
    DoEvents

    
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    
    DoEvents
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    
    If Trim(level) = "" Then
        Exit Sub
    ElseIf level = "Granted" Then
        MsgBox message, vbInformation, "SYNCRONIZED."

        Call modAPI.sp_transaction_list(tmpRs)
        If tmpRs.State = 1 Then
            If tmpRs.RecordCount > 0 Then
                DoEvents
                Call modObjTools.fillArray(AllArrAcctTitle, tmpRs)
                DoEvents
                Call modObjTools.addToListObjectUsingArrayWithCols(cboTrans, arrDesc, AllArrAcctTitle, 1, 1)
            End If
        End If
        If tmpRs.State <> 0 Then tmpRs.Close
        
        Call modAPI.sp_AccountTitles_List(tmpRs)
        If tmpRs.State = 1 Then
            If tmpRs.RecordCount > 0 Then
                DoEvents
                Call modObjTools.fillArray(AllArrAcctTitle2, tmpRs)
                DoEvents
                Call modObjTools.addToListObjectUsingArrayWithCols(cboFlxAcctCode, arrAcctTitles, AllArrAcctTitle2, 2, 0)
                DoEvents
                Call modObjTools.addToListObjectUsingArrayWithCols(cboFlxAcctTitles, arrAcctCode, AllArrAcctTitle2, 2, 1)
            End If
        End If
        If tmpRs.State <> 0 Then tmpRs.Close
    
    Else
        MsgBox message, vbCritical, "Error!"
         Exit Sub
    End If
    
Exit Sub
er:
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    Call modObjTools.ErrorMessage(Err.Description, "Error->tmrSync_Timer()")
End Sub

Private Sub txtApprvdBy_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtApprvdBy, KeyAscii, ValidChars)
    If KeyAscii = 13 Then cboStatus.SetFocus
End Sub

Private Sub txtApprvdBy_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Exit Sub
End Sub


Private Sub txtChckedBy_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtChckedBy, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtNotedBy.SetFocus
End Sub

Private Sub txtChckedBy_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Exit Sub
End Sub


Private Sub txtCheckDate_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtCheckDate, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then txtChckedBy.SetFocus
End Sub

Private Sub txtCheckNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtCheckNo, KeyAscii, ValidCodeChars)
    If KeyAscii = 13 Then txtCheckDate.SetFocus
End Sub

Private Sub txtCheckNo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Exit Sub
End Sub


Private Sub txtDocDateRec_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtDocDateRec, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then cboStatus.SetFocus
End Sub

Private Sub txtDueDate_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtDueDate, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then txtRFPNo.SetFocus
End Sub


Private Sub txtFlxAmnt_GotFocus()
    DtlsRow = flxVoucher.row
End Sub

Private Sub txtFlxAmnt_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFlxAmnt, KeyAscii, NumbersWithDecimalPoint)
    If KeyAscii = 13 Then txtFlxAmnt_LostFocus
End Sub

Private Sub txtFlxAmnt_LostFocus()
    flxVoucher.TextMatrix(flxVoucher.row, 5) = Format(txtFlxAmnt.Text, "#,##0.00")
    txtFlxAmnt.Visible = False
    flxVoucher.SetFocus
    DoEvents
    Call ComputeTotalDebitCredit
    DoEvents
    Call ComputeTotalNetPay
End Sub

Private Sub txtFlxBrchDpt_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFlxBrchDpt, KeyAscii, ValidSQLChars)
    If KeyAscii = 13 Then
        flxRFP.col = 2
        flxRFP.SetFocus
    End If
End Sub

Private Sub txtFlxBrchDpt_LostFocus()
    flxRFP.TextMatrix(DtlsRow, 1) = txtFlxBrchDpt.Text
    txtFlxBrchDpt.Visible = False
End Sub


Private Sub cboStatutory_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboStatutory, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        flxRFP.col = 1
        flxRFP.row = 1
        flxRFP.SetFocus
    End If
End Sub




Private Sub cboTrans_Change()
On Error GoTo er
    Dim i As Integer, j As Integer
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
    
    If SearchcboDesc Then
        SearchcboDesc = False
        Exit Sub
    End If

    
    With cboTrans
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters

        If i <> -1 And i <= .ListCount Then
            strTotal = .List(i)
            
            TransID = arrDesc(i, 0)
    
            DoEvents
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
               SearchcboDesc = True
               .SelText = Right$(strTotal, j)
               .SelStart = Len(strPartial)
               .SelLength = j
            End If

        End If

    End With
        
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error ->cboTrans_Change()")
End Sub

Private Sub cboTrans_Click()
    TransID = arrDesc(cboTrans.ListIndex, 0)
    Call cboTrans_KeyPress(13)
End Sub


Private Sub cboTrans_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Or KeyCode = 46 Then
        SearchcboDesc = True
    End If
       
End Sub


Private Sub cboTrans_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboTrans, KeyAscii, ValidChars)
    If KeyAscii = 13 Then flxVoucher.SetFocus
End Sub


Sub GetAccountTitles(Pid As Long)
On Error GoTo er
    
    Call modAPI.sp_get_AccountTitles_Auto(Pid, tmpRs)
    If tmpRs.State = 0 Then Exit Sub
    
    If tmpRs.RecordCount <= 0 Then
        tmpRs.Close
        Exit Sub
    End If
    
    lblTotCredit.Caption = ""
    lblTotDebit.Caption = ""
    flxVoucher.Rows = 1
    flxVoucher.Rows = 2
    tmpRs.MoveFirst
    While Not tmpRs.EOF

        If flxVoucher.TextMatrix(1, 2) = "" Then
            flxVoucher.TextMatrix(1, 2) = tmpRs!AcctCode
            flxVoucher.TextMatrix(1, 3) = tmpRs!AcctTitles
            flxVoucher.TextMatrix(1, 4) = tmpRs!Type
            flxVoucher.TextMatrix(1, 5) = ""

        Else
            flxVoucher.AddItem vbTab & _
            vbTab & tmpRs!AcctCode & _
            vbTab & tmpRs!AcctTitles & _
            vbTab & tmpRs!Type & _
            vbTab & ""
        End If

    tmpRs.MoveNext
    Wend
    
    pTransID = TransID
    flxVoucher.TextMatrix(flxVoucher.Rows - 1, 5) = Format(lblNetPay.Caption, "#,##0.00")
    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->GetAccountTitles()")
End Sub

Private Sub cmdTrans_Click(Index As Integer)
On Error GoTo er
Dim checkdate As String
'cmdTrans Legend
'0 ->NEW
'1 ->EDIT
'2 ->PRINT
'3 ->SUMMRY
'4 ->EXIT
'5 ->SAVE
'6 ->CANCEL

    If Index = 0 Then 'NEW
        formAction = "NEW"
        didRecordEntry = True
        Call SetButtonVisibility(True)
        Call LockedEntryObject(True)
        Call ClearEntryForm(False)
        
        txtPayee.SetFocus
        txtRFPDate.Text = Format(Now(), "mm/dd/YYYY")
        txtPrepBy.Text = EmployeesName
        
        Qry = "Select `VALUE` From settings Where DESCRIPTION = 'REQUEST FOR PAYMENT' and REMARKS = 'SECTION HEAD' "
        txtChckedBy.Text = GetFieldValue(Qry)
    
        Qry = "Select `VALUE` From settings Where DESCRIPTION = 'REQUEST FOR PAYMENT' and REMARKS = 'PAP CLERK' "
        txtPrepBy.Text = GetFieldValue(Qry)
        
        Qry = "Select `VALUE` From settings Where DESCRIPTION = 'REQUEST FOR PAYMENT' and REMARKS = 'DIVISION HEAD' "
        txtNotedBy.Text = GetFieldValue(Qry)
    
        Qry = "Select `VALUE` From settings Where DESCRIPTION = 'REQUEST FOR PAYMENT' and REMARKS = 'CFO' "
        txtApprvdBy.Text = GetFieldValue(Qry)
        
    ElseIf Index = 1 Then 'EDIT
    
        If txtPayee.Text = "" Then
            MsgBox "Please Select Record to be edit", vbCritical + vbOKOnly, "Access Denied!"
            Exit Sub
        End If
        
        If PrintCopy > 0 Then
            MsgBox "Unable to edit because RFP was already printed!", vbExclamation + vbOKOnly, "Access Denied!"
            Exit Sub
        End If
        
        If cboStatus.Text = "CANCELLED" Then
            MsgBox "Unable to edit because RFP was cancelled!", vbExclamation + vbOKOnly, "Access Denied"
            Exit Sub
        End If
    
        formAction = "EDIT"
        didRecordEntry = True
        Call SetButtonVisibility(True)
        Call LockedEntryObject(True)
        txtPrepBy.Text = EmployeesName
        
        txtPayee.SetFocus
        
    ElseIf Index = 2 Then 'PRINT
    
        If txtPayee.Text = "" Then
            MsgBox "No available record to print!", vbCritical + vbOKOnly, "Warning"
            Exit Sub
        End If
    
        If cboStatus.Text = "CANCELLED" Then
            MsgBox "Unable to print because RFP was cancelled!", vbExclamation + vbOKOnly, "Access Denied"
            Exit Sub
        End If
    
       Dim Qry1, Qry2, Qry3 As String
       Dim tmpRs1 As New ADODB.Recordset
       Dim tmpRs2 As New ADODB.Recordset
       Dim tmpRs3 As New ADODB.Recordset
       Screen.MousePointer = vbHourglass
       Dim rpt As New rptRFPOther
       

        Qry1 = "Call sp_mrcwarehouse.sp_get_rfp_other(" & RFPID & ");"
        DoEvents
        'InputBox Qry, Qry, Qry
        If tmpRs1.State <> 0 Then tmpRs1.Close
        Set tmpRs1 = LYRcon.Execute(Qry1)
        
        
        Qry2 = "SELECT DESCRIPTION as DTYPE,DEDUCTTYPE as TYPE, DEDUCTION_REB as REBDUCT, PARTICULARS as DESCRIPTION FROM rfp_deduct where RFPID = " & RFPID & ";"
        DoEvents
        'InputBox Qry, Qry, Qry
        If tmpRs2.State <> 0 Then tmpRs2.Close
        Set tmpRs2 = LYRcon.Execute(Qry2)

        Qry3 = "SELECT b.BRANCHNAME, ch.`TRANSACTION` as TransNAme,vou.code,vou.Description,vou.Type,vou.Amount,vou.TotDebit, vou.TotCredit FROM voucherdtls as vou " & _
              "LEFT JOIN branch as b ON b.BRANCHCODE = vou.BCODE INNER JOIN accounting_corp.trans_legend as ch ON vou.TransID = ch.TRANS_LEGEND_ID WHERE vou.RFPID = " & RFPID & " " & _
              "GROUP by vou.SeqNO Order by vou.SeqNo asc;"
              
        DoEvents
        'InputBox Qry, Qry, Qry
        If tmpRs3.State <> 0 Then tmpRs3.Close
        Set tmpRs3 = LYRcon.Execute(Qry3)

        DoEvents
        'Main report
        With rpt
            .DiscardSavedData
             DoEvents
            .Database.SetDataSource tmpRs1
            .ParameterFields(1).ClearCurrentValueAndRange
            .ParameterFields(1).AddCurrentValue txtNotedBy.Text
            .ParameterFields(2).ClearCurrentValueAndRange
            .ParameterFields(2).AddCurrentValue cboCompany.Text
        End With

        'Subreport Deductions
        With rpt.Subreport1.OpenSubreport
            .DiscardSavedData
             DoEvents
            .Database.SetDataSource tmpRs2
            .ParameterFields(1).ClearCurrentValueAndRange
            .ParameterFields(1).AddCurrentValue lblTotLess.Caption
            .ParameterFields(2).ClearCurrentValueAndRange
            .ParameterFields(2).AddCurrentValue lblTotAdd.Caption
            .ParameterFields(3).ClearCurrentValueAndRange
            .ParameterFields(3).AddCurrentValue lblNetPay.Caption
        End With

        'Subreport Account Titles
        With rpt.Subreport.OpenSubreport
            .DiscardSavedData
             DoEvents
            .Database.SetDataSource tmpRs3
            .ParameterFields(1).ClearCurrentValueAndRange
            .ParameterFields(1).AddCurrentValue lblTotCredit.Caption
            .ParameterFields(2).ClearCurrentValueAndRange
            .ParameterFields(2).AddCurrentValue lblTotDebit.Caption
        End With
        
       rpt.ReadRecords
       
       frmReport.CRViewer1.ReportSource = rpt
       frmReport.CRViewer1.ViewReport
       frmReport.CRViewer1.Zoom 100
       frmReport.CRViewer1.EnableExportButton = True
       frmReport.RFPID = RFPID
       frmReport.MySource = "RFPOTHER"
       Screen.MousePointer = vbDefault
       
       DoEvents
       frmReport.WindowState = vbMaximized
       frmReport.Show vbModal
        
    ElseIf Index = 3 Then 'SUMMARY
        didRecordEntry = False
        frmHSummary.MySource = "RFPOTHER"
        frmHSummary.Show
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
'
'
        Screen.MousePointer = vbHourglass

        Call CreatTempTable
        
        If txtCheckDate.Text = "__/__/____" Then
            checkdate = ""
        Else
            checkdate = txtCheckDate.Text
        End If


        DoEvents
        If formAction = "EDIT" Then

            ans = MsgBox("DO YOU WANT TO SAVE YOUR CHANGES?", vbQuestion + vbYesNo, Me.Caption)
            If ans = vbNo Then Exit Sub

            Call modAPI.sp_rfp_other_update(RFPID, Trim(txtRFPNo.Text), txtRFPDate.Text, txtDueDate.Text, cboStatutory.Text, txtPayee.Text, txtPrepBy.Text _
                                     , txtChckedBy.Text, txtNotedBy.Text, txtApprvdBy.Text, txtCheckNo.Text, Format(checkdate, "YYYY-MM-DD"), txtRemarks.Text, cboStatus.Text, cboTrans.Text, cboBillTo.Text, cboCompany.Text, txtDocDateRec.Text, level, message)

        ElseIf formAction = "NEW" Then

            ans = MsgBox("PLEASE MAKE SURE THAT THE DATA IS CORRECT!" & _
                vbCrLf & _
                vbCrLf & "DO YOU WISH TO CONTINUE?", vbYesNo + vbQuestion, Me.Caption)
            If ans = vbNo Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass

            Call modAPI.sp_rfp_other_create(Trim(txtRFPNo.Text), txtRFPDate.Text, txtDueDate.Text, cboStatutory.Text, txtPayee.Text, txtPrepBy.Text _
                                     , txtChckedBy.Text, txtNotedBy.Text, txtApprvdBy.Text, txtCheckNo.Text, Format(checkdate, "YYYY-MM-DD"), txtRemarks.Text, cboStatus.Text, cboTrans.Text, cboBillTo.Text, cboCompany.Text, txtDocDateRec.Text, level, message)

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
        level = ""
        message = ""
        
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
    Screen.MousePointer = vbDefault
    DoEvents
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdTrans_Click()")
End Sub


Sub ClearEntryForm(Focus1stObject As Boolean)
    
    txtPayee.Text = ""
    'txtSupAddress.Text = ""
    
    cboBillTo.Text = ""
    txtRFPNo.Text = ""
    cboStatutory.ListIndex = 0
    cboCompany.Text = ""
    
    txtRFPDate.Mask = ""
    txtRFPDate.Text = ""
    txtRFPDate.Mask = "##/##/####"
    
    txtDocDateRec.Mask = ""
    txtDocDateRec.Text = ""
    txtDocDateRec.Mask = "##/##/####"
    
    txtDueDate.Mask = ""
    txtDueDate.Text = ""
    txtDueDate.Mask = "##/##/####"
    
    cboTrans.Text = ""
    lblTotPOAmnt.Caption = "0.00"
    lblNetPay.Caption = "0.00"
    
    cboStatus.Clear
    cboStatus.AddItem "FOR CHECKING"
    cboStatus.AddItem "FOR APPROVAL"
    cboStatus.AddItem "APPROVED WITH CHECK NO."
    cboStatus.ListIndex = 0

    txtChckedBy.Text = ""
    txtPrepBy.Text = ""
    txtNotedBy.Text = ""
    txtApprvdBy.Text = ""
    
    txtCheckNo.Text = ""
    txtCheckDate.Mask = ""
    txtCheckDate.Text = ""
    txtCheckDate.Mask = "##/##/####"
    
    flxRFP.Rows = 1
    flxRFP.Rows = 2
    
    flxDeduct.Rows = 1
    flxDeduct.Rows = 2
    
    flxVoucher.Rows = 1
    flxVoucher.Rows = 2
    
    lblTotAdd.Caption = "0.00"
    lblTotLess.Caption = "0.00"
    
    lblTotCredit.Caption = "0.00"
    lblTotDebit.Caption = "0.00"
    
    txtRemarks.Text = ""
    pTransID = 0
    RFPID = 0

    If Focus1stObject = True Then cmdTrans(0).SetFocus
                              
End Sub



Private Sub ComputeTotalNetPay()
 On Error GoTo er
    Dim i As Integer
    Dim j As Integer
    Dim TotalInvoice As Double
    
    For i = 1 To flxRFP.Rows - 1
        TotalInvoice = TotalInvoice + Val(Replace(flxRFP.TextMatrix(i, 4), ",", ""))
    Next
    
    lblTotPOAmnt.Caption = Format(TotalInvoice, "#,##0.00")
    
    TotalInvoice = TotalInvoice + Val(Replace(lblTotAdd.Caption, ",", "")) - Val(Replace(lblTotLess.Caption, ",", ""))
    
    lblNetPay.Caption = Format(TotalInvoice, "#,##0.00")
    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->ComputeTotalNetPay()")
End Sub

Sub LockedEntryObject(DataEntry As Boolean)

    DataEntry = Not DataEntry
    
        txtPayee.Locked = DataEntry
        cboBillTo.Locked = DataEntry
        txtRFPNo.Locked = DataEntry
        txtCheckNo.Locked = DataEntry
        txtChckedBy.Locked = DataEntry
        txtNotedBy.Locked = DataEntry
        txtApprvdBy.Locked = DataEntry
        cboTrans.Locked = DataEntry
        txtRemarks.Locked = DataEntry
        cboStatus.Locked = DataEntry
        cboStatutory.Locked = DataEntry
        cboRFPType.Locked = DataEntry
        cboCompany.Locked = DataEntry
    If txtPayee.Locked Then
    
        txtRFPDate.Enabled = False
        txtRFPDate.BackColor = txtHoldGrayBackColor.BackColor
        txtDueDate.Enabled = False
        txtDueDate.BackColor = txtHoldGrayBackColor.BackColor
        txtCheckDate.Enabled = False
        txtCheckDate.BackColor = txtHoldGrayBackColor.BackColor
        txtDocDateRec.Enabled = False
        txtDocDateRec.BackColor = txtHoldGrayBackColor.BackColor
        
        cboStatutory.BackColor = txtHoldGrayBackColor.BackColor
        txtPayee.BackColor = txtHoldGrayBackColor.BackColor
        cboBillTo.BackColor = txtHoldGrayBackColor.BackColor
        txtRFPNo.BackColor = txtHoldGrayBackColor.BackColor
        txtCheckNo.BackColor = txtHoldGrayBackColor.BackColor
        txtChckedBy.BackColor = txtHoldGrayBackColor.BackColor
        txtNotedBy.BackColor = txtHoldGrayBackColor.BackColor
        txtApprvdBy.BackColor = txtHoldGrayBackColor.BackColor
        cboTrans.BackColor = txtHoldGrayBackColor.BackColor
        txtRemarks.BackColor = txtHoldGrayBackColor.BackColor
        cboStatus.BackColor = txtHoldGrayBackColor.BackColor
        cboRFPType.BackColor = txtHoldGrayBackColor.BackColor
        cboCompany.BackColor = txtHoldGrayBackColor.BackColor
    Else
        
        txtRFPDate.Enabled = True
        txtRFPDate.BackColor = txtHoldWhiteBackColor.BackColor
        txtDueDate.Enabled = True
        txtDueDate.BackColor = txtHoldWhiteBackColor.BackColor
        txtCheckDate.Enabled = True
        txtCheckDate.BackColor = txtHoldWhiteBackColor.BackColor
        txtDocDateRec.Enabled = True
        txtDocDateRec.BackColor = txtHoldWhiteBackColor.BackColor
        
        cboStatutory.BackColor = txtHoldWhiteBackColor.BackColor
        txtPayee.BackColor = txtHoldWhiteBackColor.BackColor
        cboBillTo.BackColor = txtHoldWhiteBackColor.BackColor
        txtRFPNo.BackColor = txtHoldWhiteBackColor.BackColor
        txtChckedBy.BackColor = txtHoldWhiteBackColor.BackColor
        txtNotedBy.BackColor = txtHoldWhiteBackColor.BackColor
        txtCheckNo.BackColor = txtHoldWhiteBackColor.BackColor
        cboTrans.BackColor = txtHoldWhiteBackColor.BackColor
        txtRemarks.BackColor = txtHoldWhiteBackColor.BackColor
        cboStatus.BackColor = txtHoldWhiteBackColor.BackColor
        txtApprvdBy.BackColor = txtHoldWhiteBackColor.BackColor
        cboRFPType.BackColor = txtHoldWhiteBackColor.BackColor
        cboCompany.BackColor = txtHoldWhiteBackColor.BackColor
    End If

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
        cmdTrans(0).SetFocus
    End If
    
End Sub


Private Sub flxRFP_Click()
    Dim i As Integer
   
    'DtlsRow = flxRFP.Row
    DtlsCol = flxRFP.col
    flexGridObject = "flxRFP"
    
    If didRecordEntry = False Then Exit Sub
    
    If flxRFP.TextMatrix(1, 1) = "" Then Exit Sub
    
    If DtlsCol = 0 Then 'Check Icon
        With flxRFP
            If .col >= 0 And .row > 0 Then
                If .TextMatrix(.row, 0) = "" Then
                    .FillStyle = flexFillSingle
                    .CellFontName = "Webdings"
                    .TextMatrix(.row, 0) = "a"
                Else
                    .FillStyle = flexFillSingle
                    .TextMatrix(.row, 0) = ""
                    .Highlight = flexHighlightNever
                End If
            End If
        End With
    End If
End Sub

Private Sub flxRFP_DblClick()
On Error GoTo er

    If didRecordEntry = False Then Exit Sub
    
    DtlsRow = flxRFP.row
    DtlsCol = flxRFP.col
    
    If DtlsCol = 1 Then
        If cboRFPType.Text <> "GASOLINE EXPENSE" Then
            With txtFlxBrchDpt
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
        Else
            With cboFlxMainBranch
             .Top = framHoldMid.Top + flxRFP.CellTop + flxRFP.Top
             .Left = framHoldMid.Left + flxRFP.CellLeft + flxRFP.Left
             .Width = flxRFP.CellWidth
             .Text = flxRFP.TextMatrix(flxRFP.row, DtlsCol)
             .Visible = True
             .SetFocus
            End With
        End If
    ElseIf DtlsCol = 2 Then
       
         With txtFlxRefNo
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
    
    ElseIf DtlsCol = 3 Then
       
         With txtFlxParticular
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
        
    ElseIf DtlsCol = 4 Then
       
         With txtFlxRFPAmnt
          .Top = framHoldMid.Top + flxRFP.CellTop + flxRFP.Top
          .Left = framHoldMid.Left + flxRFP.CellLeft + flxRFP.Left
          .Width = flxRFP.CellWidth
          .Height = flxRFP.CellHeight
          .Text = Val(Replace(flxRFP.TextMatrix(flxRFP.row, DtlsCol), ",", ""))
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


Private Sub flxRFP_GotFocus()
    flexGridObject = "flxRFP"
End Sub

Private Sub flxRFP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxRFP_DblClick
End Sub

Private Sub flxVoucher_Click()
    DtlsRow = flxVoucher.row
    DtlsCol = flxVoucher.col
    txtFlxAmnt.Visible = False
    
    If didRecordEntry = False Then Exit Sub
    
    If DtlsCol = 0 Then 'Check Icon

        With flxVoucher
            If .col >= 0 And .row > 0 Then
              If flxVoucher.TextMatrix(DtlsRow, 0) = "" Then
                    For i = 1 To flxVoucher.Rows - 1
                        flxVoucher.TextMatrix(i, 0) = ""
                    Next

                    .FillStyle = flexFillSingle
                    .CellFontName = "Webdings"
                    .TextMatrix(DtlsRow, 0) = "a"
                Else
                    .FillStyle = flexFillSingle
                    .TextMatrix(DtlsRow, 0) = ""
                    .Highlight = flexHighlightNever
                End If
            End If
        End With
        
    End If
End Sub

Private Sub flxVoucher_DblClick()
On Error GoTo er
Dim i As Integer
    
    If didRecordEntry = False Then Exit Sub
    
    DtlsRow = flxVoucher.row
    DtlsCol = flxVoucher.col
    '   1          2            3            4
    '|^CODE |<ACCOUNT TITLES |^ACCOUNT TYPE |>AMOUNT"
        If flxVoucher.col = 1 Then 'Branch

             With cboFlxBranch
              .Top = flxVoucher.CellTop + flxVoucher.Top
              .Left = flxVoucher.CellLeft + flxVoucher.Left
              .Width = flxVoucher.CellWidth
              .Visible = True
              .Text = flxVoucher.TextMatrix(flxVoucher.row, 1)
              .SetFocus
            End With
        ElseIf flxVoucher.col = 2 Then 'Account Code

             With cboFlxAcctCode
              .Top = fraHoldAcctTitles.Top + flxVoucher.CellTop + flxVoucher.Top
              .Left = fraHoldAcctTitles.Left + flxVoucher.CellLeft + flxVoucher.Left
              .Width = flxVoucher.CellWidth
              .Visible = True
              .Text = flxVoucher.TextMatrix(flxVoucher.row, 2)
              .SetFocus
            End With
        
        ElseIf flxVoucher.col = 3 Then 'Account Titles

             With cboFlxAcctTitles
              .Top = fraHoldAcctTitles.Top + flxVoucher.CellTop + flxVoucher.Top
              .Left = fraHoldAcctTitles.Left + flxVoucher.CellLeft + flxVoucher.Left
              .Width = flxVoucher.CellWidth
              .Visible = True
              .Text = flxVoucher.TextMatrix(flxVoucher.row, 3)
              .SetFocus
            End With
            
        ElseIf flxVoucher.col = 4 Then 'Account Type

             With cboFlxAcctType
              .Top = fraHoldAcctTitles.Top + flxVoucher.CellTop + flxVoucher.Top
              .Left = fraHoldAcctTitles.Left + flxVoucher.CellLeft + flxVoucher.Left
              .Width = flxVoucher.CellWidth
              .Visible = True
              .SetFocus
            End With
                
        ElseIf flxVoucher.col = 5 Then 'AMOUNT
        
            With txtFlxAmnt
              .Top = fraHoldAcctTitles.Top + flxVoucher.CellTop + flxVoucher.Top
              .Left = fraHoldAcctTitles.Left + flxVoucher.CellLeft + flxVoucher.Left
              .Width = flxVoucher.CellWidth
              .Visible = True
              .Text = Val(Replace(flxVoucher.TextMatrix(flxVoucher.row, 5), ",", ""))
              .SelStart = 0
              .SelLength = Len(.Text)
              .ZOrder
              .SetFocus
            End With
    
        End If

Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "flxVoucher_DblClick()")
End Sub


Private Sub flxVoucher_GotFocus()
    flexGridObject = "flxVoucher"
End Sub


Private Sub flxVoucher_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call flxVoucher_DblClick
End Sub


Private Sub flxVoucher_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If didRecordEntry = False Then Exit Sub
    If Button = vbRightButton Then
        PopupMenu mnuPopUpFormat, vbPopupMenuRightButton
        flexGridObject = "flxVoucher"
    End If
End Sub


Private Sub Form_Load()
On Error GoTo er
    
    Me.Caption = "Request for Payment in Other Payables"
    Warehouse.sBar.Panels(6).Text = Me.Caption
    
    cboStatutory.Clear
    cboStatutory.AddItem ""
    cboStatutory.AddItem "DECLARED"
    cboStatutory.AddItem "UNDECLARED"
    
    cboFlxParticular.Clear
    cboFlxParticular.AddItem "DEBIT MEMO"
    cboFlxParticular.AddItem "CREDIT MEMO"
    cboFlxParticular.AddItem "EXPANDED WITH HOLDING TAX 1%"
    cboFlxParticular.AddItem "EXPANDED WITH HOLDING TAX 2%"
    cboFlxParticular.AddItem "EXPANDED WITH HOLDING TAX 5%"
    
    cboRFPType.Clear
    cboRFPType.AddItem ""
    cboRFPType.AddItem "GASOLINE EXPENSE"
    
    DoEvents
    Call LoadCboRecord
    
    'Call DisableRighClick(True)
    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error-Form_Load()")
End Sub



Sub LoadCboRecord()
On Error GoTo er
    
    flxVoucher.Rows = 1
    flxVoucher.Rows = 2
    
'    DoEvents
'    Call modAPI.sp_supplier_list(tmpRs)
'    If tmpRs.State = 1 Then
'        If tmpRs.RecordCount > 0 Then
'            DoEvents
'            Call modObjTools.fillArray(AllArrSupplier, tmpRs)
'            DoEvents
'            Call modObjTools.addToListObjectUsingArrayWithCols(txtPayee, arrSupplier, AllArrSupplier, 3, 1)
'        End If
'    End If
'    If tmpRs.State <> 0 Then tmpRs.Close
    
    DoEvents
    Call modAPI.sp_company_list(tmpRs)
    If tmpRs.State = 1 Then
        If tmpRs.RecordCount > 0 Then
            DoEvents
            Call modObjTools.fillArray(AllArrBillTo, tmpRs)
            DoEvents
            Call modObjTools.addToListObjectUsingArrayWithCols(cboBillTo, arrBillTo, AllArrBillTo, 0, 0)
        End If
    End If
    If tmpRs.State <> 0 Then tmpRs.Close
    
    
    Call modAPI.sp_transaction_list(tmpRs)
    If tmpRs.State = 1 Then
        If tmpRs.RecordCount > 0 Then
            DoEvents
            Call modObjTools.fillArray(AllArrAcctTitle, tmpRs)
            DoEvents
            Call modObjTools.addToListObjectUsingArrayWithCols(cboTrans, arrDesc, AllArrAcctTitle, 1, 1)
        End If
    End If
    If tmpRs.State <> 0 Then tmpRs.Close
    
    Call modAPI.sp_AccountTitles_List(tmpRs)
    If tmpRs.State = 1 Then
        If tmpRs.RecordCount > 0 Then
            DoEvents
            Call modObjTools.fillArray(AllArrAcctTitle2, tmpRs)
            DoEvents
            Call modObjTools.addToListObjectUsingArrayWithCols(cboFlxAcctCode, arrAcctCode, AllArrAcctTitle2, 2, 0)
            DoEvents
            Call modObjTools.addToListObjectUsingArrayWithCols(cboFlxAcctTitles, arrAcctTitles, AllArrAcctTitle2, 2, 1)
        End If
    End If
    If tmpRs.State <> 0 Then tmpRs.Close
    
    DoEvents
    Call modMotoR.sp_get_Branch(tmpRs)
    If tmpRs.State = 1 Then
        DoEvents
        Call modObjTools.addToListObjectWithArrayGivenField(cboFlxBranch, tmpRs, arrBRANCH, "BRANCHCODE", "BRANCHNAME")
    End If
    If tmpRs.State <> 0 Then tmpRs.Close
    
    DoEvents
    Call modMotoR.sp_get_Branch(tmpRs)
    If tmpRs.State = 1 Then
        DoEvents
        Call modObjTools.addToListObjectWithArrayGivenField(cboFlxMainBranch, tmpRs, arrBRANCH, "BRANCHCODE", "BRANCHNAME")
    End If
    If tmpRs.State <> 0 Then tmpRs.Close
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->LoadCboRecord()")
End Sub


Private Sub RXResizer1_ResizeComplete()
On Error Resume Next
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->PRINT
'3 ->SUMMARY
'4 ->EXIT
'5 ->SAVE
'6 ->CANCEL

    'Align Save Button
    cmdTrans(5).Left = cmdTrans(0).Left
    cmdTrans(5).Top = cmdTrans(0).Top

    'Align Cancel Button
    cmdTrans(6).Left = Me.Width - 1600
    cmdTrans(6).Top = cmdTrans(0).Top
    
    If flxRFP.TextMatrix(flxRFP.row, 1) <> "" Then Exit Sub
    
    With flxRFP
        .Clear
        .Rows = 2
        .Cols = 5
        .FixedCols = 0
        
        '               0            1                         2                             3             4
        .FormatString = " |<BRANCH/DEPARTMENT  |^REF. NO. / BILLING NO. / ACCT. NO.   |<PARTICULARS    |>AMOUNT "
        
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.2
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.19
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.4
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.18


    End With
    
    With flxDeduct
        .Clear
        .Rows = 2
        .Cols = 5
        .FixedCols = 0
        
        '               0       1           2           3              4
        .FormatString = "  |<DESCRIPTION |^TYPE |<PARTICULARS |>DEDUCTION/REBATE "
        
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.26
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.1
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.47
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.15
    End With

    With flxVoucher
        .Clear
        .Rows = 2
        .Cols = 7
        .FixedCols = 0
        
        .FormatString = "|<BRANCH |^CODE |<ACCOUNT TITLES |^ACCOUNT TYPE |>AMOUNT "
        
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.3  'CODE
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.15  'CODE
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.3   'DESCRIPTION
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.1  'ACCOUNT TYPE
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.1  'AMOUNT
        .ColWidth(6) = 0 '(.Width - .ColWidth(0)) * 0.1  'AMOUNT
    End With

    With flxRBTransHist
        '.Clear
        .Rows = 2
        .Cols = 5
        .WordWrap = True
        .AllowUserResizing = flexResizeBoth
        '               0     1        2      3           4
        .FormatString = " |^DATE |<TRANSACTION NAME |^STATUS |<RESPONSIBLE"
        
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.2
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.3
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.15
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.4
    End With
End Sub

Private Sub CreatTempTable()
On Error GoTo er
Dim Qry As String

        Qry = "Drop Temporary Table IF  Exists sp_mrcwarehouse.tmp_rfpdtls;"
        Call LYRcon.Execute(Qry)
        
        
        DoEvents
        Qry = "Create Temporary Table IF Not Exists sp_mrcwarehouse.tmp_rfpdtls( " & _
        "   SOLDTO          varchar(60), " & _
        "   SHIPPEDTO       varchar(60), " & _
        "   INVNO           varchar(30), " & _
        "   PARTICULARS     LONGTEXT, " & _
        "   INVAMNT         Decimal(16,2), " & _
        "   REBDUCT         Decimal(16, 2), " & _
        "   NETPAY          Decimal(16, 2), " & _
        "   SEQNO           INT); "
        
        Call LYRcon.Execute(Qry)
        
        
        Qry = "Drop Temporary Table IF  Exists sp_mrcwarehouse.tmp_rfpdeduct;"
        Call LYRcon.Execute(Qry)
        
        DoEvents
        Qry = "Create Temporary Table IF Not Exists sp_mrcwarehouse.tmp_rfpdeduct( " & _
        "   DTYPE           varchar(100), " & _
        "   TYPE            varchar(10),   " & _
        "   DESCRIPTION     varchar(100),   " & _
        "   REBDUCT         decimal(16, 2), " & _
        "   TOTALLESS       decimal(16, 2),     " & _
        "   TOTALADD        decimal(16, 2),    " & _
        "   SEQNO           Int(11));"
        
        'Text1.Text = Qry DTYPE,TYPE,DESCRIPTION,REBDUCT
        Call LYRcon.Execute(Qry)
        
        
        Qry = "Drop Temporary Table IF  Exists sp_mrcwarehouse.tmp_voucherdtls;"
        Call LYRcon.Execute(Qry)
        
        DoEvents
        Qry = "Create Temporary Table IF Not Exists sp_mrcwarehouse.tmp_voucherdtls( " & _
        "   seqno       int(10) , " & _
        "   TransID     int(10),   " & _
        "   code        varchar(20),  " & _
        "   Description varchar(500),   " & _
        "   Type        varchar(6),  " & _
        "   Amount      decimal(16,2),  " & _
        "   TotDebit    decimal(16,2),  " & _
        "   TotCredit   decimal(16,2), " & _
        "   bcode       Varchar(1000));"
        
        'Text1.Text = Qry
        Call LYRcon.Execute(Qry)

    
    Call InsertTempTable
                
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CreatTempTable()")
End Sub

Private Sub InsertTempTable()
On Error GoTo er
Dim Qry As String


        Qry = "Insert Into sp_mrcwarehouse.tmp_rfpdtls(SOLDTO,SHIPPEDTO,INVNO,PARTICULARS,INVAMNT,NETPAY,SEQNO) Values "

        DtlsRow = 0
        For DtlsRow = 1 To flxRFP.Rows - 1

                If DtlsRow > 1 Then Qry = Qry & ","

                Qry = Qry & "('" & cboBillTo.Text & "','" & flxRFP.TextMatrix(DtlsRow, 1) & "','" & flxRFP.TextMatrix(DtlsRow, 2) & _
                            "','" & flxRFP.TextMatrix(DtlsRow, 3) & _
                            "'," & Val(Replace(flxRFP.TextMatrix(DtlsRow, 4), ",", "")) & "," & Val(Replace(lblNetPay.Caption, ",", "")) & "," & DtlsRow & ")"
         Next

        'InputBox Qry, Qry, Qry
        Call LYRcon.Execute(Qry)
        
        
        Qry = "Insert Into sp_mrcwarehouse.tmp_rfpdeduct(DTYPE,TYPE,DESCRIPTION,REBDUCT,TOTALLESS,TOTALADD,SEQNO) Values "

        DtlsRow = 0
        For DtlsRow = 1 To flxDeduct.Rows - 1
            
                If DtlsRow > 1 Then Qry = Qry & ","
    
                Qry = Qry & "('" & flxDeduct.TextMatrix(DtlsRow, 1) & "','" & flxDeduct.TextMatrix(DtlsRow, 2) & _
                            "','" & flxDeduct.TextMatrix(DtlsRow, 3) & "','" & Val(Replace(Trim(flxDeduct.TextMatrix(DtlsRow, 4)), ",", "")) & "','" & Val(Replace(Trim(lblTotLess.Caption), ",", "")) & "','" & Val(Replace(Trim(lblTotAdd.Caption), ",", "")) & "'," & DtlsRow & ")"
         Next
    
        'InputBox Qry, Qry, Qry
        Call LYRcon.Execute(Qry)
        
    
        Qry = "Insert Into sp_mrcwarehouse.tmp_voucherdtls(seqno, TransID, code, Description, Type, Amount, TotDebit, TotCredit,BCODE) Values "

        DtlsRow = 0
        For DtlsRow = 1 To flxVoucher.Rows - 1
            
                If DtlsRow > 1 Then Qry = Qry & ","
    
                Qry = Qry & "(" & DtlsRow & "," & TransID & ",'" & flxVoucher.TextMatrix(DtlsRow, 2) & "','" & flxVoucher.TextMatrix(DtlsRow, 3) & _
                            "','" & flxVoucher.TextMatrix(DtlsRow, 4) & "', '" & Val(Replace(flxVoucher.TextMatrix(DtlsRow, 5), ",", "")) & _
                            "','" & Val(Replace(lblTotDebit.Caption, ",", "")) & "', '" & Val(Replace(lblTotCredit.Caption, ",", "")) & "','" & flxVoucher.TextMatrix(DtlsRow, 6) & "')"
         Next
    
        'InputBox Qry, Qry, Qry
        Call LYRcon.Execute(Qry)
        
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->InsertTempTable()")
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
    
    If txtPayee.Text <> "" Then
        If cboStatus.Text = "APPROVED WITH CHECK NO." Or cboStatus.Text = "CANCELLED" Then
            cmdTrans(1).Enabled = False 'Edit Button
        Else
            cmdTrans(1).Enabled = True 'Edit Button
        End If
        cmdTrans(2).Enabled = True 'PRINT Button
        cmdHistory.Enabled = True 'Show History Button
    Else
        cmdTrans(1).Enabled = False 'Edit Button
        cmdTrans(2).Enabled = False 'PRINT Button
    End If
 
End Sub

Private Sub txtFlxBrchDpt_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Exit Sub
End Sub

Private Sub txtFlxDeductReb_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFlxDeductReb, KeyAscii, NumbersWithDecimalPoint)
    If KeyAscii = 13 Then
    
        If Val(lblTotPOAmnt.Caption) = 0 Then
            txtFlxDeductReb.Visible = False
            Exit Sub
        End If

        flxDeduct.TextMatrix(DtlsRow, 4) = Format(txtFlxDeductReb.Text, "#,##0.00")
        flxDeduct.SetFocus
        txtFlxDeductReb.Visible = False
        Call ComputeTotalNetPay
        Call ComputeTotAddLess
    End If
End Sub

Private Sub txtFlxDeductReb_LostFocus()
    Call txtFlxDeductReb_KeyPress(13)
End Sub

Private Sub txtFlxDeductReb_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Exit Sub
End Sub

Private Sub txtFlxParticular_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFlxParticular, KeyAscii, ValidSQLChars)
    If KeyAscii = 13 Then
        flxRFP.col = 4
        flxRFP.SetFocus
    End If
End Sub


Private Sub txtFlxParticular_LostFocus()
    flxRFP.TextMatrix(DtlsRow, 3) = txtFlxParticular.Text
    txtFlxParticular.Visible = False
End Sub


Private Sub txtFlxParticular_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Exit Sub
End Sub

Private Sub txtFlxRefNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFlxRefNo, KeyAscii, ValidSQLChars)
    If KeyAscii = 13 Then
        flxRFP.col = 3
        flxRFP.SetFocus
    End If
End Sub


Private Sub txtFlxRefNo_LostFocus()
    If cboRFPType.Text = "GASOLINE EXPENSE" Then
        
    Else
        flxRFP.TextMatrix(DtlsRow, 2) = txtFlxRefNo.Text
        txtFlxRefNo.Visible = False
    End If
End Sub


Private Sub txtFlxRefNo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Exit Sub
End Sub

Private Sub txtFlxRFPAmnt_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFlxRFPAmnt, KeyAscii, NumbersWithDecimalPoint)
    If KeyAscii = 13 Then flxRFP.SetFocus
End Sub

Private Sub txtFlxRFPAmnt_LostFocus()
    flxRFP.TextMatrix(DtlsRow, 4) = Format(Val(txtFlxRFPAmnt.Text), "#,##0.00")
    txtFlxRFPAmnt.Visible = False
    
    DoEvents
    Call ComputeTotalNetPay
    
    If flxRFP.row = flxRFP.Rows - 1 Then
    
        ans = MsgBox("Add new row?", vbYesNo + vbQuestion, Me.Caption)
        If ans = vbYes Then
            flxRFP.AddItem ""
            flxRFP.col = 1
            flxRFP.row = flxRFP.Rows - 1
        End If
    
    End If
    
    flxRFP.SetFocus
    
End Sub

Private Sub txtFlxRFPAmnt_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Exit Sub
End Sub

Private Sub txtFlxValueDeduct_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFlxValueDeduct, KeyAscii, ValidSQLChars)
    If KeyAscii = 13 Then
        flxDeduct.TextMatrix(DtlsRow, 1) = txtFlxValueDeduct.Text
        txtFlxValueDeduct.Visible = False
        flxDeduct.SetFocus
    End If
End Sub

Private Sub txtFlxValueDeduct_LostFocus()
    Call txtFlxValueDeduct_KeyPress(13)
End Sub


Private Sub txtNotedBy_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtNotedBy, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtApprvdBy.SetFocus
End Sub

Private Sub txtNotedBy_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Exit Sub
End Sub

Private Sub txtPayee_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtPayee, KeyAscii, ValidChars)
    If KeyAscii = 13 Then cboBillTo.SetFocus
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtRemarks, KeyAscii, ValidSQLChars)
    If KeyAscii = 13 Then txtCheckNo.SetFocus
End Sub

Private Sub txtRemarks_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Exit Sub
End Sub


Private Sub txtRFPDate_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtRFPDate, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then cboStatutory.SetFocus
End Sub


Private Sub txtRFPNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtRFPNo, KeyAscii, ValidCodeChars)
    If KeyAscii = 13 Then txtRFPDate.SetFocus
End Sub




