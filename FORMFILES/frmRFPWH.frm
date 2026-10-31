VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRFPWH 
   Caption         =   "Request For Payment"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   14910
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   38
      Tag             =   "XYZ"
      Top             =   9120
      Width           =   14655
      Begin VB.ComboBox cboCompany 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         ItemData        =   "frmRFPWH.frx":0000
         Left            =   6720
         List            =   "frmRFPWH.frx":000A
         TabIndex        =   85
         Tag             =   "XYZ"
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtRemarks 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   59
         Tag             =   "XYZ"
         Top             =   240
         Width           =   13095
      End
      Begin VB.TextBox txtCheckedBy 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   11760
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   51
         Tag             =   "XYZ"
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtPrepBy 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   42
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
         TabIndex        =   41
         Tag             =   "XYZ"
         Top             =   1080
         Width           =   2775
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
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtCheckNo 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   39
         Tag             =   "XYZ"
         Top             =   1080
         Width           =   1935
      End
      Begin MSMask.MaskEdBox txtChckNoDate 
         Height          =   315
         Left            =   1440
         TabIndex        =   43
         Tag             =   "XYZ"
         Top             =   1440
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
      Begin MSMask.MaskEdBox txtDocDateRec 
         Height          =   315
         Left            =   6720
         TabIndex        =   82
         Tag             =   "XYZ"
         Top             =   1320
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
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Caption         =   "COMPANY:"
         Height          =   255
         Left            =   5640
         TabIndex        =   86
         Tag             =   "XYZ"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label28 
         Caption         =   "MM/DD/YYYY"
         Height          =   255
         Left            =   8760
         TabIndex        =   84
         Tag             =   "XYZ"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DOC DATE REC:"
         Height          =   375
         Left            =   5280
         TabIndex        =   83
         Tag             =   "xyz"
         Top             =   1335
         Width           =   1455
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "REMARKS:"
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
         TabIndex        =   60
         Tag             =   "XYZ"
         Top             =   285
         Width           =   1455
      End
      Begin VB.Label Label20 
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
         TabIndex        =   50
         Tag             =   "XYZ"
         Top             =   765
         Width           =   1455
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
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
         Left            =   10320
         TabIndex        =   49
         Tag             =   "XYZ"
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label18 
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
         Left            =   10680
         TabIndex        =   48
         Tag             =   "XYZ"
         Top             =   1080
         Width           =   1095
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
         Left            =   10320
         TabIndex        =   47
         Tag             =   "XYZ"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "CHECK  DATE:"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Tag             =   "XYZ"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "MM/DD/YYYY"
         Height          =   255
         Left            =   3480
         TabIndex        =   45
         Tag             =   "XYZ"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "CHECK NO.:"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Tag             =   "XYZ"
         Top             =   1100
         Width           =   975
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   7
      Left            =   2280
      TabIndex        =   79
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "&RECONCILIATION"
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
      Image           =   "frmRFPWH.frx":004B
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Timer tmrSync 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9600
      Top             =   0
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
      Height          =   2775
      Left            =   120
      TabIndex        =   65
      Tag             =   "xyz"
      Top             =   6360
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
         TabIndex        =   81
         Tag             =   "XYZ"
         Top             =   960
         Visible         =   0   'False
         Width           =   855
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
         TabIndex        =   77
         Tag             =   "XYZ"
         Top             =   960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtFlxAmnt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   11880
         TabIndex        =   76
         Tag             =   "XYZ"
         Text            =   "txtFlxAmnt"
         Top             =   960
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
         Left            =   3120
         TabIndex        =   75
         Tag             =   "XYZ"
         Top             =   960
         Visible         =   0   'False
         Width           =   6015
      End
      Begin VB.ComboBox cboFlxAcctType 
         Height          =   315
         ItemData        =   "frmRFPWH.frx":05E5
         Left            =   9600
         List            =   "frmRFPWH.frx":05EF
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
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
         Height          =   345
         ItemData        =   "frmRFPWH.frx":0602
         Left            =   1440
         List            =   "frmRFPWH.frx":0604
         Locked          =   -1  'True
         TabIndex        =   66
         Tag             =   "XYZ"
         Top             =   240
         Width           =   13095
      End
      Begin lvButton.lvButtons_H cmdSync 
         Height          =   375
         Left            =   12480
         TabIndex        =   68
         Tag             =   "XYZ"
         Top             =   2235
         Width           =   2040
         _ExtentX        =   3598
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
         Image           =   "frmRFPWH.frx":0606
         cBack           =   -2147483633
      End
      Begin MSFlexGridLib.MSFlexGrid flxVoucher 
         Height          =   1575
         Left            =   120
         TabIndex        =   67
         Tag             =   "XYZ"
         Top             =   600
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   2778
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         Enabled         =   -1  'True
         AllowUserResizing=   3
         FormatString    =   $"frmRFPWH.frx":5DF8
      End
      Begin VB.Label Label25 
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
         TabIndex        =   73
         Tag             =   "XYZ"
         Top             =   270
         Width           =   1305
      End
      Begin VB.Label lblTotDebit 
         Alignment       =   2  'Center
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   72
         Tag             =   "XYF"
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label23 
         Caption         =   "TOTAL DEBIT:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   71
         Tag             =   "XYF"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblTotCredit 
         Alignment       =   2  'Center
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   70
         Tag             =   "XYF"
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label22 
         Caption         =   "TOTAL CREDIT:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Tag             =   "XYF"
         Top             =   2160
         Width           =   1335
      End
   End
   Begin VB.ComboBox cboFlxRemarks 
      Height          =   315
      ItemData        =   "frmRFPWH.frx":5F0A
      Left            =   5880
      List            =   "frmRFPWH.frx":5F0C
      TabIndex        =   64
      Top             =   4920
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.ComboBox cboFlxParticular 
      Height          =   315
      ItemData        =   "frmRFPWH.frx":5F0E
      Left            =   4080
      List            =   "frmRFPWH.frx":5F10
      TabIndex        =   58
      Top             =   4920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox cboFlxType 
      Height          =   315
      ItemData        =   "frmRFPWH.frx":5F12
      Left            =   3240
      List            =   "frmRFPWH.frx":5F1C
      Style           =   2  'Dropdown List
      TabIndex        =   57
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtFlxDeductReb 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   7080
      TabIndex        =   56
      Top             =   4920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtFlxReference 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   480
      TabIndex        =   55
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFlxDesc 
      Height          =   285
      Left            =   1800
      TabIndex        =   54
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFlxValueRFP 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   720
      TabIndex        =   53
      Top             =   2280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSMask.MaskEdBox txtFlxDate 
      Height          =   315
      Left            =   3120
      TabIndex        =   52
      Tag             =   "XYZ"
      Top             =   2280
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
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   15120
      TabIndex        =   35
      Text            =   "HoldBackColor"
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   15120
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "txtHoldGrayBackColor"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer tmrGetRecord 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10800
      Top             =   0
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
      Height          =   2055
      Left            =   120
      TabIndex        =   24
      Tag             =   "xyz"
      Top             =   4320
      Width           =   14655
      Begin VB.TextBox txtTotLess 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Tag             =   "XYZ"
         Text            =   "0.00"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox txtTotAdd 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   4440
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Tag             =   "XYZ"
         Text            =   "0.00"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox txtTotalDed 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   9360
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Tag             =   "XYZ"
         Text            =   "0.00"
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtNetPay 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   12600
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Tag             =   "XYZ"
         Text            =   "0.00"
         Top             =   1560
         Width           =   1935
      End
      Begin MSFlexGridLib.MSFlexGrid flxDeduct 
         Height          =   1215
         Left            =   120
         TabIndex        =   29
         Tag             =   "XYZ"
         Top             =   240
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   2143
         _Version        =   393216
         Cols            =   7
         RowHeightMin    =   315
         BackColor       =   16777215
         BackColorBkg    =   12632256
         HighLight       =   2
         GridLines       =   2
         AllowUserResizing=   3
         FormatString    =   " |^REFERENCE      |<DESCRIPTION     |^TYPE       |<PARTICULARS            |^REMARKS         |^DEDUCTION/REB                "
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
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL LESS:"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Tag             =   "XYZ"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL ADD:"
         Height          =   255
         Left            =   3480
         TabIndex        =   32
         Tag             =   "XYZ"
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL ADJUSTMENT:"
         Height          =   255
         Left            =   7560
         TabIndex        =   31
         Tag             =   "XYZ"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL NET PAY:"
         Height          =   255
         Left            =   11280
         TabIndex        =   30
         Tag             =   "XYZ"
         Top             =   1560
         Width           =   1335
      End
   End
   Begin VB.Frame framHoldMid 
      Height          =   2655
      Left            =   120
      TabIndex        =   20
      Tag             =   "xyz"
      Top             =   1680
      Width           =   14655
      Begin MSFlexGridLib.MSFlexGrid flxRFP 
         Height          =   2055
         Left            =   120
         TabIndex        =   22
         Tag             =   "XYZ"
         Top             =   240
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   10
         RowHeightMin    =   315
         BackColor       =   16777215
         BackColorBkg    =   12632256
         WordWrap        =   -1  'True
         HighLight       =   2
         GridLines       =   2
         AllowUserResizing=   3
         FormatString    =   $"frmRFPWH.frx":5F2B
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
      Begin VB.TextBox txtTotInvAmnt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   12600
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Tag             =   "XYZ"
         Text            =   "0.00"
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CommandButton cmdAltA 
         Caption         =   "&A"
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton cmdAltD 
         Caption         =   "&D"
         Height          =   255
         Left            =   720
         TabIndex        =   62
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL:"
         Height          =   255
         Left            =   12000
         TabIndex        =   23
         Tag             =   "XYZ"
         Top             =   2355
         Width           =   615
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
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
      Image           =   "frmRFPWH.frx":5FFC
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   0
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
      Image           =   "frmRFPWH.frx":6596
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   4200
      TabIndex        =   2
      Top             =   0
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
      Image           =   "frmRFPWH.frx":6B30
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5280
      TabIndex        =   3
      Top             =   0
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
      Image           =   "frmRFPWH.frx":C322
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6600
      TabIndex        =   4
      Top             =   0
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
      Image           =   "frmRFPWH.frx":C8BC
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   15120
      TabIndex        =   36
      Top             =   360
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
      Image           =   "frmRFPWH.frx":120AE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   15120
      TabIndex        =   37
      Top             =   0
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
      Image           =   "frmRFPWH.frx":12648
      cBack           =   -2147483633
   End
   Begin VB.Frame framHoldTop 
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Tag             =   "xyz"
      Top             =   360
      Width           =   14655
      Begin VB.ComboBox txtTerm 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         ItemData        =   "frmRFPWH.frx":12BE2
         Left            =   2040
         List            =   "frmRFPWH.frx":12BF8
         Locked          =   -1  'True
         TabIndex        =   78
         Tag             =   "XYZ"
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox chkFinal 
         Caption         =   "FINAL"
         Enabled         =   0   'False
         Height          =   255
         Left            =   6720
         TabIndex        =   63
         Tag             =   "XY"
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox cboRFPType 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         ItemData        =   "frmRFPWH.frx":12C2D
         Left            =   4680
         List            =   "frmRFPWH.frx":12C37
         Locked          =   -1  'True
         TabIndex        =   19
         Tag             =   "XYZ"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtTtlPayable 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   11280
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Tag             =   "XYZ"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtPayee 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   100
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
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
         TabIndex        =   7
         Tag             =   "XYZ"
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox cboSupName 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   6
         Tag             =   "XYZ"
         Top             =   600
         Width           =   4575
      End
      Begin MSMask.MaskEdBox txtRFPDate 
         Height          =   315
         Left            =   11280
         TabIndex        =   8
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
      Begin VB.Label lblHeadingRecon 
         BackStyle       =   0  'Transparent
         Caption         =   "RECONCILIATION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   6720
         TabIndex        =   80
         Tag             =   "XYZ"
         Top             =   240
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "RFP TYPE:"
         Height          =   255
         Left            =   3720
         TabIndex        =   18
         Tag             =   "XYZ"
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL PAYABLE:"
         Height          =   255
         Left            =   9840
         TabIndex        =   17
         Tag             =   "XYZ"
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "PAYEE:"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Tag             =   "XYZ"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER NAME:"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Tag             =   "XYZ"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "TERM:"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Tag             =   "XYZ"
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "RFP NUMBER:"
         Height          =   255
         Left            =   10080
         TabIndex        =   11
         Tag             =   "XYZ"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DATE PROCESSED:"
         Height          =   255
         Left            =   9720
         TabIndex        =   10
         Tag             =   "XYZ"
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "MM/DD/YYYY"
         Height          =   255
         Left            =   13320
         TabIndex        =   9
         Tag             =   "XYZ"
         Top             =   675
         Width           =   1095
      End
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   10200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   15150
      m_MDI_Height    =   11520
   End
   Begin VB.Menu mnuPopUpFormat 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuADD 
         Caption         =   "&ADD ROW"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "&DELETE ROW"
      End
   End
End
Attribute VB_Name = "frmRFPWH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpRs As New ADODB.Recordset

Dim formAction As String
Dim flexGridObject As String
Dim ans As String
Dim Qry As String, MySupp As String

Dim supplierid As Long, myRFPID As Long, cntDuplicate As Long
Dim i As Long, j As Long
Dim DtlsCol As Integer
Dim DtlsRow As Integer
Dim TtlLessValue As Double
Dim TtlAddValue As Double

Dim didRecordEntry As Boolean
Dim SearchcboSup As Boolean
Dim SearchcboShipTo As Boolean
Dim SearchcboCheckedBy As Boolean
Dim SearchcboParticular As Boolean

Dim TransID As Long
Dim pTransID As Long
Dim arrDesc() As String
Dim arrAcctTitles() As String
Dim arrAcctCode() As String
Dim arrBRANCH() As String

Dim SearchcboDesc As Boolean
Dim SearchcboAcctTitles As Boolean
Dim SearchcboAcctCode As Boolean
Dim RECON As Boolean

Dim StatSrch As Boolean
Dim strPartial As String, strTotal As String
Dim PrtBln As Boolean, PrtSrch As Boolean, consign As Boolean
Dim checkedByPosition As String, notedByPosition As String, preparedByPosition As String, approvedByposition As String

Public Sub flxCBDOpt()
    If txtTerm.Text = "CBD" Then
        With flxRFP
            .Rows = 2
            .Cols = 13
            .WordWrap = True
            .AllowUserResizing = flexResizeBoth
                '               0     1          2         3        4          5            6         7            8          9           10          11
            .FormatString = " |<SPCRR |<INVOICE DATE |^INV. NO |^REF. NO |^PACKING NO |^DATE REC |^P.O NO. |^P.O AMOUNT |^INV. AMNT|^SPCRR AMOUNT |^DUE DATE "
            .ColWidth(0) = 200
            .ColWidth(1) = (.Width - .ColWidth(0)) * 0.09
            .ColWidth(2) = (.Width - .ColWidth(0)) * 0.08
            .ColWidth(3) = (.Width - .ColWidth(0)) * 0.09
            .ColWidth(4) = (.Width - .ColWidth(0)) * 0.15
            .ColWidth(5) = (.Width - .ColWidth(0)) * 0.08
            .ColWidth(6) = (.Width - .ColWidth(0)) * 0.08
            .ColWidth(7) = (.Width - .ColWidth(0)) * 0.09
            .ColWidth(8) = (.Width - .ColWidth(0)) * 0.13
            .ColWidth(9) = (.Width - .ColWidth(0)) * 0.13
            .ColWidth(10) = (.Width - .ColWidth(0)) * 0.13
            .ColWidth(11) = (.Width - .ColWidth(0)) * 0.08
            .ColWidth(12) = 0 '(.Width - .ColWidth(0)) * 0.08
        End With
    Else
        With flxRFP
            '.Clear
            .Rows = 2
            .Cols = 11
            .WordWrap = True
            .AllowUserResizing = flexResizeBoth
            '               0     1          2              3           4             5             6             7             8                 9         10
            .FormatString = " |<SPCRR |<INVOICE DATE |^INVOICE NO |^PACKING NO |^DATE RECEIVED  |^P.O NO. |^INVOICE AMOUNT |^SPCRR AMOUNT |^DUE DATE "
            .ColWidth(0) = 200
            .ColWidth(1) = (.Width - .ColWidth(0)) * 0.1
            .ColWidth(2) = (.Width - .ColWidth(0)) * 0.1
            .ColWidth(3) = (.Width - .ColWidth(0)) * 0.11
            .ColWidth(4) = (.Width - .ColWidth(0)) * 0.1
            .ColWidth(5) = (.Width - .ColWidth(0)) * 0.1
            .ColWidth(6) = (.Width - .ColWidth(0)) * 0.11
            .ColWidth(7) = (.Width - .ColWidth(0)) * 0.13
            .ColWidth(8) = (.Width - .ColWidth(0)) * 0.13
            .ColWidth(9) = (.Width - .ColWidth(0)) * 0.1
            .ColWidth(10) = 0 '(.Width - .ColWidth(0)) * 0.1
        End With
    End If
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
        lblTotCredit.Caption = Format(tmpRs!TotCredit, "#,##0.00")
        lblTotDebit.Caption = Format(tmpRs!TotDebit, "#,##0.00")
            
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
    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->ShowAcctTitles()")
End Sub



Private Sub cboCompany_Change()
On Error GoTo er
    Dim i As Integer, j As Integer
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
    'If didRecordEntry = False Then Exit Sub
    
    If SearchcboSup Then
        SearchcboSup = False
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
               SearchcboSup = True
               .SelText = Right$(strTotal, j)
               .SelStart = Len(strPartial)
               .SelLength = j
            End If
        End If

    End With
        
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error ->cboCompany_change()")
End Sub

Private Sub cboCompany_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Or KeyCode = 46 Then
        SearchcboSup = True
    End If
End Sub

Private Sub cboCompany_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboCompany, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtDocDateRec.SetFocus
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
On Error GoTo er
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
    
        If pTransID = TransID Then Exit Sub
        Call GetAccountTitles(TransID)
    
        flxVoucher.col = 4
        flxVoucher.row = flxVoucher.Rows - 1
        flxVoucher.SetFocus
        
    End If
    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboTrans_KeyPress()")
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
    
    flxVoucher.TextMatrix(flxVoucher.Rows - 1, 5) = Format(txtNetPay.Text, "#,##0.00")
    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->GetAccountTitles()")
End Sub

Private Sub cmdSync_Click()
On Error GoTo er
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



Public Sub ZeroSPCRRAMNT()
Dim intLoopCount As Long, myOldSPCRR As String

    If txtFlxValueRFP.Text = flxRFP.TextMatrix(flxRFP.row, 1) Then
        Exit Sub
    End If
    
    For intLoopCount = 1 To flxRFP.Rows - 1
        myOldSPCRR = flxRFP.TextMatrix(intLoopCount, 1)
        cntDuplicate = 0
        If myOldSPCRR = txtFlxValueRFP.Text Then 'flxPARTS.TextMatrix(actrow, 1) Then
            flxRFP.TextMatrix(flxRFP.row, 9) = 0
            cntDuplicate = cntDuplicate + 1
            Exit Sub
        End If
    Next intLoopCount
End Sub

Public Sub ShowDeduction(ByVal pRFPID As Long)
On Error GoTo er
    Dim i As Integer
    
    Qry = "call sp_acctpayinvcng.sp_load_rfp_deduction(" & Val(pRFPID) & ")"
    With tmpRs
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenStatic, adLockOptimistic
        
        If .RecordCount <= 0 Then
            .Close
            Exit Sub
        End If
        
        If .RecordCount > 0 Then
            flxDeduct.Rows = 1
            flxDeduct.Rows = 2
            flxDeduct.Rows = .RecordCount + 1
            .MoveFirst
            For i = 1 To flxDeduct.Rows - 1
                flxDeduct.TextMatrix(i, 1) = IIf(IsNull(!REFERENCE), "", !REFERENCE)
                flxDeduct.TextMatrix(i, 2) = IIf(IsNull(!Description), "", !Description)
                flxDeduct.TextMatrix(i, 3) = IIf(IsNull(!DEDUCTTYPE), "", !DEDUCTTYPE)
                flxDeduct.TextMatrix(i, 4) = IIf(IsNull(!Particulars), "", !Particulars)
                flxDeduct.TextMatrix(i, 5) = IIf(IsNull(!DEDUCTREMARKS), "", !DEDUCTREMARKS)
                flxDeduct.TextMatrix(i, 6) = IIf(IsNull(!DEDUCTION_REB), "", Format(!DEDUCTION_REB, "#,###,###,###.00"))
                .MoveNext
            Next i
        End If
        If .State <> 0 Then .Close
    End With
    Call ComputeTotAddLess
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->ShowDeduction()")
End Sub

Public Sub ShowRecord(ByVal pRFPNO As String)
Dim rs As New ADODB.Recordset
Dim reconVal As String
On Error GoTo er
    Dim i As Integer
    
    myRFPID = 0
    
    Qry = "call sp_acctpayinvcng.sp_load_rfp_info('" & pRFPNO & "')"
    With rs
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic

        If .RecordCount <= 0 Then
            .Close
            Exit Sub
        End If
        
        If .RecordCount > 0 Then
            .MoveFirst
            myRFPID = rs!RFPID
            If !RFPSTATUS = "PENDING" Then
                chkFinal.Value = 0
            Else
                chkFinal.Value = 1
            End If
            txtPayee.Text = IIf(IsNull(!payee), "", !payee)
            cboSupName.Text = IIf(IsNull(!SUPPLIERNAME), "", !SUPPLIERNAME)
            txtRFPNo.Text = IIf(IsNull(!rfpno), "", !rfpno)
            txtRFPDate.Text = Format(!RFPDATE, "mm/dd/yyyy")
            txtTerm.Text = IIf(IsNull(!Terms), "", !Terms)
            cboRFPType.Text = IIf(IsNull(!RFPTYPE), "", !RFPTYPE)
            txtPrepBy.Text = IIf(IsNull(!PreparedBy), "", !PreparedBy)
            txtCheckedBy.Text = IIf(IsNull(!CHECKEDBY), "", !CHECKEDBY)
            txtNotedBy.Text = IIf(IsNull(!NOTEDBY), "", !NOTEDBY)
            txtApprvdBy.Text = IIf(IsNull(!ApprovedBy), "", !ApprovedBy)
            txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
            txtCheckNo.Text = IIf(IsNull(!CheckNo), "", !CheckNo)
            If IsNull(!checkdate) = False Then txtChckNoDate.Text = Format(!checkdate, "mm/dd/yyyy")
            reconVal = IIf(IsNull(!RECONCILIATION), "N", !RECONCILIATION)
            If IsNull(!docdaterec) = False Then txtDocDateRec.Text = Format(!docdaterec, "mm/dd/yyyy")
            cboCompany.Text = IIf(IsNull(!COMPANY), "LYR MOTORPARTS (RETAIL)", !COMPANY)
            supplierid = GetFieldValue("SELECT SUPPLIERID FROM SUPPLIER WHERE SUPPLIERNAME = '" & cboSupName.Text & "' ")
            Call flxCBDOpt
            flxRFP.Rows = .RecordCount + 1
            For i = 1 To flxRFP.Rows - 1
                flxRFP.TextMatrix(i, 1) = IIf(IsNull(!SPCRRNO), "", !SPCRRNO)
                flxRFP.TextMatrix(i, 2) = IIf(IsNull(!INVOICEDATE), "", Format(!INVOICEDATE, "mm/dd/yyyy"))
                flxRFP.TextMatrix(i, 3) = IIf(IsNull(!InvoiceNO), "", !InvoiceNO)
                
                If txtTerm.Text <> "CBD" Then
                    flxRFP.TextMatrix(i, 4) = IIf(IsNull(!PACKINGNO), "", !PACKINGNO)
                    flxRFP.TextMatrix(i, 5) = IIf(IsNull(!DATEREC), "", Format(!DATEREC, "mm/dd/yyyy"))
                    flxRFP.TextMatrix(i, 6) = IIf(IsNull(!PONO), "", !PONO)
                    flxRFP.TextMatrix(i, 7) = IIf(IsNull(!INVAMNT), "", Format(!INVAMNT, "#,###,###,###.00"))
                    flxRFP.TextMatrix(i, 8) = IIf(IsNull(!SPCRRAMNT), "", Format(!SPCRRAMNT, "#,###,###,###.00"))
                    flxRFP.TextMatrix(i, 9) = IIf(IsNull(!DueDate), "", Format(!DueDate, "mm/dd/yyyy"))
                Else
                    flxRFP.TextMatrix(i, 4) = IIf(IsNull(!RefNo), "", !RefNo)
                    flxRFP.TextMatrix(i, 5) = IIf(IsNull(!PACKINGNO), "", !PACKINGNO)
                    flxRFP.TextMatrix(i, 6) = IIf(IsNull(!DATEREC), "", Format(!DATEREC, "mm/dd/yyyy"))
                    flxRFP.TextMatrix(i, 7) = IIf(IsNull(!PONO), "", !PONO)
                    flxRFP.TextMatrix(i, 8) = IIf(IsNull(!POAMNT), "", Format(!POAMNT, "#,###,###,###.00"))
                    flxRFP.TextMatrix(i, 9) = IIf(IsNull(!INVAMNT), "", Format(!INVAMNT, "#,###,###,###.00"))
                    flxRFP.TextMatrix(i, 10) = IIf(IsNull(!SPCRRAMNT), "", Format(!SPCRRAMNT, "#,###,###,###.00"))
                    flxRFP.TextMatrix(i, 11) = IIf(IsNull(!DueDate), "", Format(!DueDate, "mm/dd/yyyy"))
                End If
                .MoveNext
            Next i
        End If
        If .State <> 0 Then .Close
    End With

    If reconVal = "Y" And txtTerm.Text = "CBD" Then
        lblHeadingRecon.Visible = True
        lblHeadingRecon.Caption = "RECONCILED"
        RECON = True
    ElseIf reconVal = "N" And txtTerm.Text = "CBD" Then
        lblHeadingRecon.Visible = True
        lblHeadingRecon.Caption = "UNRECONCILED"
    Else
        lblHeadingRecon.Visible = False
    End If

    DoEvents
    Call ShowDeduction(myRFPID)
    Call ShowAcctTitles(myRFPID)
    Call EnabledSelectedButton
    Call ComputeTotalNetPay
    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->ShowRecord()")
End Sub

Private Sub InsertTempTable()
On Error GoTo er
Dim Qry As String, myREFNO As String
Dim myPOAmnt As Currency, myInvAmnt As Currency, mySPCRRAmnt As Currency

        Qry = "Insert Into sp_acctpayinvcng.tmp_rfpdtls(SPCRRNO, INVOICEDATE, INVOICENO, PACKINGNO, DATERECEIVED, PONO, INVOICEAMNT, DUEDATE, SEQNO, POAMNT, SPCRRAMNT, REFNO, CONSIGNEE) Values "

        i = 0
        For i = 1 To flxRFP.Rows - 1
        
            If txtTerm.Text = "CBD" Then
                If flxRFP.TextMatrix(i, 2) = "__/__/____" Then
                    flxRFP.TextMatrix(i, 2) = ""
                End If
                myPOAmnt = Val(Replace(flxRFP.TextMatrix(i, 8), ",", ""))
                myInvAmnt = Val(Replace(flxRFP.TextMatrix(i, 9), ",", ""))
                mySPCRRAmnt = Val(Replace(flxRFP.TextMatrix(i, 10), ",", ""))
                
                If i > 1 Then Qry = Qry & ","
                
                         Qry = Qry & "('" & flxRFP.TextMatrix(i, 1) & "','" & Format(flxRFP.TextMatrix(i, 2), "YYYY-mm-dd") & "','" & flxRFP.TextMatrix(i, 3) & _
                                       "','" & Trim(flxRFP.TextMatrix(i, 5)) & "','" & Format(flxRFP.TextMatrix(i, 6), "YYYY-mm-dd") & "','" & Trim(flxRFP.TextMatrix(i, 7)) & _
                                       "','" & Val(myInvAmnt) & "','" & Format(flxRFP.TextMatrix(i, 11), "YYYY-mm-dd") & "'," & i & ",'" & Val(myPOAmnt) & "','" & Val(mySPCRRAmnt) & "','" & flxRFP.TextMatrix(i, 4) & "','" & flxRFP.TextMatrix(i, 12) & "')"

            Else
                myREFNO = ""
                myPOAmnt = 0
                myInvAmnt = Val(Replace(flxRFP.TextMatrix(i, 7), ",", ""))
                mySPCRRAmnt = Val(Replace(flxRFP.TextMatrix(i, 8), ",", ""))
                If flxRFP.TextMatrix(i, 2) = "__/__/____" Then
                    flxRFP.TextMatrix(i, 2) = ""
                End If
                If i > 1 Then Qry = Qry & ","
                
                         Qry = Qry & "('" & flxRFP.TextMatrix(i, 1) & "','" & Format(flxRFP.TextMatrix(i, 2), "YYYY-mm-dd") & "','" & flxRFP.TextMatrix(i, 3) & _
                                       "','" & Trim(flxRFP.TextMatrix(i, 4)) & "','" & Format(flxRFP.TextMatrix(i, 5), "YYYY-mm-dd") & "','" & Trim(flxRFP.TextMatrix(i, 6)) & _
                                       "','" & Val(myInvAmnt) & "','" & Format(flxRFP.TextMatrix(i, 9), "YYYY-mm-dd") & "'," & i & ",'" & Val(myPOAmnt) & "','" & Val(mySPCRRAmnt) & "','" & myREFNO & "','" & flxRFP.TextMatrix(i, 10) & "')"
                      
            End If
                
        Next
    
        'InputBox Qry, Qry, Qry
        Call LYRcon.Execute(Qry)
        
        Qry = "Insert Into sp_acctpayinvcng.tmp_rfpdeduct(REFERENCE, DESCRIPTION, DEDUCTTYPE,  PARTICULARS, REMARKS, DEDUCTREB, TOTALLESS, TOTALADD, SEQNO) Values "

        i = 0
        For i = 1 To flxDeduct.Rows - 1
            If i > 1 Then Qry = Qry & ","

            Qry = Qry & "('" & flxDeduct.TextMatrix(i, 1) & "','" & flxDeduct.TextMatrix(i, 2) & "','" & flxDeduct.TextMatrix(i, 3) & "','" & flxDeduct.TextMatrix(i, 4) & _
                        "','" & flxDeduct.TextMatrix(i, 5) & "','" & Val(Replace(Trim(flxDeduct.TextMatrix(i, 6)), ",", "")) & _
                        "','" & Val(Replace(txtTotLess.Text, ",", "")) & "', '" & Val(Replace(txtTotAdd.Text, ",", "")) & "', " & i & ")"
        Next
    
        'InputBox Qry, Qry, Qry
        Call LYRcon.Execute(Qry)
        
        Qry = "Insert Into sp_acctpayinvcng.tmp_voucherdtls(seqno, TransID, code, Description, Type, Amount, TotDebit, TotCredit,BCODE) Values "

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
    Call modObjTools.ErrorMessage(Err.Description, "Error->InsertTempTable()")
End Sub


Private Sub CreatTempTable()
On Error GoTo er
Dim Qry As String

        Qry = "Drop Temporary Table IF  Exists sp_acctpayinvcng.tmp_rfpdtls;"
        Call LYRcon.Execute(Qry)
        
        DoEvents
        Qry = "Create Temporary Table IF Not Exists sp_acctpayinvcng.tmp_rfpdtls( " & _
        "   SPCRRNO         varchar(60),        " & _
        "   INVOICEDATE     varchar(60),               " & _
        "   INVOICENO       varchar(60),        " & _
        "   PACKINGNO       varchar(60),        " & _
        "   DATERECEIVED    varchar(60),               " & _
        "   PONO            varchar(60),        " & _
        "   INVOICEAMNT     Decimal(20, 2),     " & _
        "   DUEDATE         varchar(60),               " & _
        "   SEQNO           int(11),           " & _
        "   POAMNT          Decimal(20, 2),           " & _
        "   SPCRRAMNT       Decimal(20, 2),    " & _
        "   REFNO           varchar(60),      " & _
        "   CONSIGNEE       varchar(60));      "
        
        'InputBox Qry, Qry, Qry
        Call LYRcon.Execute(Qry)
        

        Qry = "Drop Temporary Table IF  Exists sp_acctpayinvcng.tmp_rfpdeduct;"
        Call LYRcon.Execute(Qry)
        
        DoEvents
        Qry = "Create Temporary Table IF Not Exists sp_acctpayinvcng.tmp_rfpdeduct( " & _
        "   REFERENCE       varchar(60),        " & _
        "   DESCRIPTION     varchar(100),       " & _
        "   DEDUCTTYPE      varchar(30),        " & _
        "   PARTICULARS     varchar(100),       " & _
        "   REMARKS         varchar(300),       " & _
        "   DEDUCTREB       decimal(20, 2),     " & _
        "   TOTALLESS       decimal(20, 2),     " & _
        "   TOTALADD        decimal(20, 2),    " & _
        "   SEQNO           int(11));            "
        
        'Text1.Text = Qry
        Call LYRcon.Execute(Qry)
'

        Qry = "Drop Temporary Table IF  Exists sp_acctpayinvcng.tmp_voucherdtls;"
        Call LYRcon.Execute(Qry)
        
        DoEvents
        Qry = "Create Temporary Table IF Not Exists sp_acctpayinvcng.tmp_voucherdtls( " & _
        "   seqno       int(10) , " & _
        "   TransID     int(10),   " & _
        "   code        varchar(20),  " & _
        "   Description varchar(500),   " & _
        "   Type        varchar(6),  " & _
        "   Amount      decimal(16,2),  " & _
        "   TotDebit    decimal(16,2),  " & _
        "   TotCredit   decimal(16,2), " & _
        "   bcode       Varchar(1000));"
    
        Call LYRcon.Execute(Qry)
        
    Call InsertTempTable
                
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CreatTempTable()")
End Sub


Function ValidateEntry() As String
    Dim vMsg As String, Qry As String
    Dim i As Integer, Cntr As Integer
    
    vMsg = ""
    
    If cboSupName.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field supplier name."
    If supplierid = 0 Then vMsg = vMsg & vbCrLf & "Supplier name doesn't exist!"
    If txtPayee.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field PAYEE"

    If Not IsDate(txtRFPDate.Text) Then
        vMsg = vMsg & vbCrLf & "Invalid RFP date."
    End If
    

    If Not IsDate(txtChckNoDate.Text) Then
        'vMsg = vMsg & vbCrLf & "Invalid due date."
        txtChckNoDate.Text = "00/00/0000"
    End If
    
    If txtTerm.Text = "CBD" Then
        For i = 1 To flxRFP.Rows - 1
            If flxRFP.TextMatrix(i, 7) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field p.o no."
            If flxRFP.TextMatrix(i, 8) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field invoice amount."
        Next i
    Else
        For i = 1 To flxRFP.Rows - 1
            If flxRFP.TextMatrix(i, 1) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field spcrr no."
            If flxRFP.TextMatrix(i, 2) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field invoice date"
            If flxRFP.TextMatrix(i, 3) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field invoice no."
            If flxRFP.TextMatrix(i, 6) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field p.o no."
            If flxRFP.TextMatrix(i, 7) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field invoice amount."
        Next i
    End If
    
    If txtPrepBy.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field prepared by name."
    If txtCheckedBy.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field checked by name."
    If txtNotedBy.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field noted by name."
    If txtApprvdBy.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field approved by name."
    If Val(TransID) <= 0 Then vMsg = vMsg & vbCrLf & "Transaction not found in the List"
    If Val(txtDocDateRec.Text) <= 0 Then
        If Not IsDate(txtDocDateRec.Text) Then vMsg = vMsg & vbCrLf & "Invalid Document date receive format."
    End If
    
    ValidateEntry = vMsg
    
End Function


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
'        If txtCheckNo.Text = "" Then
'            cmdTrans(1).Enabled = True 'Edit Button
''        Else
''            cmdTrans(1).Enabled = False 'Edit Button
'        End If
        cmdTrans(1).Enabled = True 'Edit Button
        cmdTrans(2).Enabled = True 'PRINT Button
    Else
        cmdTrans(1).Enabled = False 'Edit Button
        cmdTrans(2).Enabled = False 'PRINT Button
    End If
        
    If txtTerm.Text = "CBD" And Val(chkFinal.Value) = 1 Then
        cmdTrans(7).Enabled = True
    Else
        cmdTrans(7).Enabled = False
    End If
End Sub

Private Sub ComputeTotAddLess()
On Error GoTo er
    Dim i As Integer
    Dim TotalAdd As Double
    Dim TotalLess As Double
    
    For i = 1 To flxDeduct.Rows - 1
        If flxDeduct.TextMatrix(i, 3) = "ADD" Then
            TotalAdd = TotalAdd + Val(Replace(flxDeduct.TextMatrix(i, 6), ",", ""))
        ElseIf flxDeduct.TextMatrix(i, 3) = "LESS" Then
            TotalLess = TotalLess + Val(Replace(flxDeduct.TextMatrix(i, 6), ",", ""))
        End If
    Next
    
    txtTotAdd.Text = Format(TotalAdd, "#,##0.00")
    txtTotLess.Text = Format(TotalLess, "#,##0.00")
    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Erro->ComputeTotAddLess()")
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
               SearchcboParticular = True
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

Private Sub cboFlxParticular_Click()
Call cboFlxParticular_KeyPress(13)
End Sub

Private Sub cboFlxParticular_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Or KeyCode = 46 Then
        SearchcboParticular = True
    End If
End Sub

Private Sub cboFlxParticular_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        Dim LessValue As Double

        flxDeduct.TextMatrix(flxDeduct.row, 4) = cboFlxParticular.Text
        If flxDeduct.TextMatrix(flxDeduct.row, 3) = "LESS" Then
            If flxDeduct.TextMatrix(flxDeduct.row, 4) = "EXPANDED WITH HOLDING TAX 1%" Then
                LessValue = (Val(Replace(txtTotInvAmnt.Text, ",", "")) / 1.12) * 0.01
                flxDeduct.TextMatrix(flxDeduct.row, 6) = Format(LessValue, "#,###,##0.00")
            End If
        ElseIf flxDeduct.TextMatrix(flxDeduct.row, 3) = "ADD" Then
            If cboFlxParticular.Text = "VAT(12%)" Then
                LessValue = (Val(Replace(txtTotInvAmnt.Text, ",", "")) * 0.12) * 0.01
                flxDeduct.TextMatrix(flxDeduct.row, 6) = Format(LessValue, "#,###,##0.00")
            End If
        End If
        
        cboFlxParticular.Visible = False
        flxDeduct.SetFocus
        Call ComputeTotAddLess
        Call ComputeTotalNetPay
        
        
        If flxDeduct.TextMatrix(flxDeduct.row, 4) = "" Then
            flxDeduct.col = 4
            flxDeduct.row = DtlsRow
        Else
            flxDeduct.col = 5
            flxDeduct.row = DtlsRow
        End If
        
    End If
End Sub

Private Sub cboFlxParticular_LostFocus()
    cboFlxParticular.Visible = False
    cboFlxParticular.Text = ""
End Sub




'Private Sub cboFlxRemarks_Change()
'On Error GoTo er
'    Dim i As Integer, j As Integer
'    Dim strPartial As String, strTotal As String
'    Dim x As Integer
'
'    If didRecordEntry = False Then Exit Sub
'
'    If SearchcboParticular Then
'        SearchcboParticular = False
'        Exit Sub
'    End If
'
'    With cboFlxRemarks
'        strPartial = .Text
'        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
'
'        If i <> -1 And i <= .ListCount Then
'            strTotal = .List(i)
'
'            DoEvents
'            j = Len(strTotal) - Len(strPartial)
'            If j <> 0 Then
'               SearchcboParticular = True
'               .SelText = Right$(strTotal, j)
'               .SelStart = Len(strPartial)
'               .SelLength = j
'            End If
'
'        End If
'
'    End With
'
'Exit Sub
'er:
'    Call modObjTools.ErrorMessage(Err.Description, "Error ->cboFlxRemarks_Change()")
'End Sub

Private Sub cboFlxRemarks_Click()
Call cboFlxRemarks_KeyPress(13)
End Sub


'Private Sub cboFlxRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 8 Or KeyCode = 46 Then
'        SearchcboParticular = True
'    End If
'End Sub

Private Sub cboFlxRemarks_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        DoEvents
        DoEvents
        flxDeduct.TextMatrix(DtlsRow, 5) = cboFlxRemarks.Text
        cboFlxRemarks.Visible = False
        flxDeduct.SetFocus
        If flxDeduct.TextMatrix(DtlsRow, 5) = "" Then
            DoEvents
            flxDeduct.col = 5
            flxDeduct.row = DtlsRow
        Else
            DoEvents
            flxDeduct.col = 6
            flxDeduct.row = DtlsRow
        End If
    End If
End Sub

Private Sub cboFlxRemarks_LostFocus()
On Error Resume Next
    DoEvents
    DoEvents
    flxDeduct.SetFocus
    If flxDeduct.TextMatrix(DtlsRow, 5) = "" Then
        DoEvents
        flxDeduct.col = 5
        flxDeduct.row = DtlsRow
    Else
        DoEvents
        flxDeduct.col = 6
        flxDeduct.row = DtlsRow
    End If
End Sub

Private Sub cboFlxType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxDeduct.TextMatrix(flxDeduct.row, 3) = cboFlxType.Text
        cboFlxType.Visible = False
        flxDeduct.SetFocus
        If flxDeduct.TextMatrix(flxDeduct.row, 3) = "" Then
            flxDeduct.col = 3
            flxDeduct.row = DtlsRow
        Else
            flxDeduct.col = 4
            flxDeduct.row = DtlsRow
        End If
    End If
End Sub

Private Sub cboFlxType_LostFocus()
    Call cboFlxType_KeyPress(13)
End Sub

Private Sub cboRFPType_Change()
On Error GoTo er
    Dim i As Integer, j As Integer
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
    'If didRecordEntry = False Then Exit Sub
    
    If SearchcboSup Then
        SearchcboSup = False
        Exit Sub
    End If


    With cboRFPType
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
    Call modObjTools.ErrorMessage(Err.Description, "Error ->cboRFPType_Change()")
End Sub

Private Sub cboRFPType_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboRFPType, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtRFPNo.SetFocus
End Sub

Private Sub cmdAltA_Click()
Call mnuAdd_Click
End Sub

Private Sub cmdAltD_Click()
Call mnuDel_Click
End Sub

Private Sub flxDeduct_Click()
    flexGridObject = "flxDeduct"
End Sub

Private Sub flxDeduct_DblClick()
On Error GoTo er

    If didRecordEntry = False Then Exit Sub
   
    DtlsCol = flxDeduct.col
    DtlsRow = flxDeduct.row
    If DtlsCol = 1 Then
       With txtFlxReference
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
       With txtFlxDesc
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
    ElseIf DtlsCol = 3 Then
       With cboFlxType
        .Top = framHoldButtton.Top + flxDeduct.CellTop + flxDeduct.Top
        .Left = framHoldButtton.Left + flxDeduct.CellLeft + flxDeduct.Left
        .Width = flxDeduct.CellWidth
        .Visible = True
        .ZOrder
        .SetFocus
      End With
    ElseIf DtlsCol = 4 Then
       With cboFlxParticular
        .Top = framHoldButtton.Top + flxDeduct.CellTop + flxDeduct.Top
        .Left = framHoldButtton.Left + flxDeduct.CellLeft + flxDeduct.Left
        .Width = flxDeduct.CellWidth
        .Text = flxDeduct.TextMatrix(flxDeduct.row, DtlsCol)
        .Visible = True
        .ZOrder
        .SetFocus
      End With
    ElseIf DtlsCol = 5 Then
       With cboFlxRemarks
        .Top = framHoldButtton.Top + flxDeduct.CellTop + flxDeduct.Top
        .Left = framHoldButtton.Left + flxDeduct.CellLeft + flxDeduct.Left
        .Width = flxDeduct.CellWidth
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .ZOrder
        .SetFocus
      End With
    ElseIf DtlsCol = 6 Then
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

Private Sub flxVoucher_Click()
    Dim i As Integer
    
    DtlsRow = flxVoucher.row
    DtlsCol = flxVoucher.col
    flexGridObject = "flxVoucher"
    
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
    '   1          2            3            4
    '|<BRANCH |^CODE |<ACCOUNT TITLES |^ACCOUNT TYPE |>AMOUNT"
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
              .Top = flxVoucher.CellTop + flxVoucher.Top
              .Left = flxVoucher.CellLeft + flxVoucher.Left
              .Width = flxVoucher.CellWidth
              .Visible = True
              .Text = flxVoucher.TextMatrix(flxVoucher.row, 2)
              .SetFocus
            End With
        
        ElseIf flxVoucher.col = 3 Then 'Account Titles

             With cboFlxAcctTitles
              .Top = flxVoucher.CellTop + flxVoucher.Top
              .Left = flxVoucher.CellLeft + flxVoucher.Left
              .Width = flxVoucher.CellWidth
              .Visible = True
              .Text = flxVoucher.TextMatrix(flxVoucher.row, 3)
              .SetFocus
            End With
            
        ElseIf flxVoucher.col = 4 Then 'Account Type

             With cboFlxAcctType
              .Top = flxVoucher.CellTop + flxVoucher.Top
              .Left = flxVoucher.CellLeft + flxVoucher.Left
              .Width = flxVoucher.CellWidth
              .Visible = True
              .SetFocus
            End With
                
        ElseIf flxVoucher.col = 5 Then 'AMOUNT
        
            With txtFlxAmnt
              .Top = flxVoucher.CellTop + flxVoucher.Top
              .Left = flxVoucher.CellLeft + flxVoucher.Left
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

Private Sub flxVoucher_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
       Call flxVoucher_DblClick
    End If
End Sub

Private Sub flxVoucher_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If didRecordEntry = False Then Exit Sub
    If Button = vbRightButton Then
        PopupMenu mnuPopUpFormat, vbPopupMenuRightButton
        flexGridObject = "flxVoucher"
    End If
End Sub

Private Sub tmrSync_Timer()
On Error GoTo er
Dim level As String
Dim message As String

    tmrSync.Enabled = False
    Screen.MousePointer = vbHourglass
    Warehouse.Enabled = False
    

    DoEvents
        Call modAPI.sp_sync_acct_titles(level, message)
    DoEvents

    
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    
    DoEvents
    Warehouse.Enabled = True
    Screen.MousePointer = vbDefault
    
    If Trim(level) = "" Then
        Exit Sub
    ElseIf level = "Granted" Then
        MsgBox message, vbInformation, "SYNCRONIZED."
        Call LoadCboRecord
    Else
        MsgBox message, vbCritical, "Error!"
         Exit Sub
    End If
    
Exit Sub
er:
    Warehouse.Enabled = True
    Screen.MousePointer = vbDefault
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    Call modObjTools.ErrorMessage(Err.Description, "Error->tmrRefresh_Timer()")
End Sub

Sub LoadCboRecord()
On Error GoTo er
    
    flxVoucher.Rows = 1
    flxVoucher.Rows = 2
    
    Call modAPI.sp_transaction_list(tmpRs)
    If tmpRs.State = 1 Then
        If tmpRs.RecordCount > 0 Then
            Call modObjTools.fillArray(modGlobalVar.AllArrAcctTitle, tmpRs)
        Else
            Exit Sub
        End If
    End If
    If tmpRs.State <> 0 Then tmpRs.Close
    
    DoEvents
    Call modObjTools.addToListObjectUsingArrayWithCols(cboTrans, arrDesc, modGlobalVar.AllArrAcctTitle, 1, 1)
    
    Call modAPI.sp_AccountTitles_List(tmpRs)
    If tmpRs.State = 1 Then
        If tmpRs.RecordCount > 0 Then
            Call modObjTools.fillArray(modGlobalVar.AllArrAcctTitle2, tmpRs)
        Else
            Exit Sub
        End If
    End If
    If tmpRs.State <> 0 Then tmpRs.Close
    
    DoEvents
    Call modObjTools.addToListObjectUsingArrayWithCols(cboFlxAcctCode, arrAcctTitles, modGlobalVar.AllArrAcctTitle2, 2, 0)
    DoEvents
    Call modObjTools.addToListObjectUsingArrayWithCols(cboFlxAcctTitles, arrAcctCode, modGlobalVar.AllArrAcctTitle2, 2, 1)

    DoEvents
    Call modMotoR.sp_get_Branch(tmpRs)
    If tmpRs.State = 1 Then
        Call modObjTools.addToListObjectWithArrayGivenField(cboFlxBranch, tmpRs, arrBRANCH, "BRANCHCODE", "BRANCHNAME")
    End If
    If tmpRs.State <> 0 Then tmpRs.Close

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->LoadCboRecord()")
End Sub

Private Sub txtDocDateRec_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtDocDateRec, KeyAscii, NumbersOnly)
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
    flxVoucher.TextMatrix(DtlsRow, 5) = Format(txtFlxAmnt.Text, "#,##0.00")
    txtFlxAmnt.Visible = False
    flxVoucher.SetFocus
    Call ComputeTotalDebitCredit
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
            flxVoucher.TextMatrix(flxVoucher.row, 2) = arrAcctTitles(i, 0)
            flxVoucher.TextMatrix(flxVoucher.row, 3) = arrAcctTitles(i, 1)
            flxVoucher.TextMatrix(flxVoucher.row, 4) = arrAcctTitles(i, 2)
    
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


Private Sub cboFlxAcctTitles_Click()
    'Acct_Code,Acct_Titles,`type`
    flxVoucher.TextMatrix(flxVoucher.row, 2) = arrAcctTitles(cboFlxAcctTitles.ListIndex, 0)
    flxVoucher.TextMatrix(flxVoucher.row, 3) = arrAcctTitles(cboFlxAcctTitles.ListIndex, 1)
    flxVoucher.TextMatrix(flxVoucher.row, 4) = arrAcctTitles(cboFlxAcctTitles.ListIndex, 2)
    Call cboFlxAcctTitles_LostFocus
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
    cboFlxAcctTitles.Visible = False
    flxVoucher.SetFocus
End Sub

Private Sub cboFlxAcctType_GotFocus()
    DtlsRow = flxVoucher.row
End Sub

Private Sub cboFlxAcctType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxVoucher.TextMatrix(flxVoucher.row, 4) = cboFlxAcctType.Text
        cboFlxAcctType.Visible = False
        flxVoucher.SetFocus
    End If
End Sub


Private Sub cboFlxAcctType_LostFocus()
    flxVoucher.TextMatrix(DtlsRow, 4) = cboFlxAcctType.Text
    cboFlxAcctType.Visible = False
    flxVoucher.SetFocus
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
            flxVoucher.TextMatrix(DtlsRow, 2) = arrAcctCode(i, 0)
            flxVoucher.TextMatrix(DtlsRow, 3) = arrAcctCode(i, 1)
            flxVoucher.TextMatrix(DtlsRow, 4) = arrAcctCode(i, 2)
    
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
    flxVoucher.TextMatrix(DtlsRow, 2) = arrAcctCode(cboFlxAcctCode.ListIndex, 0)
    flxVoucher.TextMatrix(DtlsRow, 3) = arrAcctCode(cboFlxAcctCode.ListIndex, 1)
    flxVoucher.TextMatrix(DtlsRow, 4) = arrAcctCode(cboFlxAcctCode.ListIndex, 2)
    Call cboFlxAcctCode_LostFocus
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
    cboFlxAcctCode.Visible = False
    flxVoucher.SetFocus
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

Private Sub mnuAdd_Click()
On Error GoTo er
    
    If didRecordEntry = False Then Exit Sub
    
    If flexGridObject = "flxRFP" Then
        If txtTerm.Text = "CBD" And RECON = False Then
            If flxRFP.TextMatrix(flxRFP.Rows - 1, 6) = "" And flxRFP.TextMatrix(flxRFP.Rows - 1, 7) = "" Then
                'MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
                flxRFP.SetFocus
                flxRFP.col = 7
                flxRFP.row = flxRFP.Rows - 1
                Exit Sub
            Else
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
        
                flxRFP.col = 6
                flxRFP.row = flxRFP.Rows - 1
                flxRFP.SetFocus
            End If
        ElseIf txtTerm.Text = "CBD" And RECON = True Then
            If flxRFP.TextMatrix(flxRFP.Rows - 1, 1) = "" And flxRFP.TextMatrix(flxRFP.Rows - 1, 3) = "" Then
                'MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
                flxRFP.SetFocus
                flxRFP.col = 1
                flxRFP.row = flxRFP.Rows - 1
                Exit Sub
            Else
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
        
                flxRFP.col = 1
                flxRFP.row = flxRFP.Rows - 1
                flxRFP.SetFocus
            End If
        ElseIf txtTerm.Text <> "CBD" Then
            If flxRFP.TextMatrix(flxRFP.Rows - 1, 1) = "" And flxRFP.TextMatrix(flxRFP.Rows - 1, 3) = "" Then
                'MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
                flxRFP.SetFocus
                flxRFP.col = 1
                flxRFP.row = flxRFP.Rows - 1
                Exit Sub
            Else
                flxRFP.AddItem "" & _
                vbTab & "" & _
                vbTab & "" & _
                vbTab & "" & _
                vbTab & "" & _
                vbTab & "" & _
                vbTab & "" & _
                vbTab & "" & _
                vbTab & ""
        
                flxRFP.col = 1
                flxRFP.row = flxRFP.Rows - 1
                flxRFP.SetFocus
            End If
        End If
    ElseIf flexGridObject = "flxDeduct" Then
        flxDeduct.AddItem "" & _
        vbTab & "" & _
        vbTab & "" & _
        vbTab & "" & _
        vbTab & "" & _
        vbTab & ""
        
        flxDeduct.col = 1
        flxDeduct.row = flxDeduct.Rows - 1
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


'
'        flxVoucher.AddItem "" & _
'        vbTab & "" & _
'        vbTab & "" & _
'        vbTab & "" & _
'        vbTab & "" & _
'        vbTab & "" & _
'        vbTab & "" 'flxVoucher.TextMatrix(flxVoucher.Rows - 1, 6)
'
'        If InStr(flxVoucher.TextMatrix(flxVoucher.Rows - 2, 2), "CASH IN BANK") Then
'            flxVoucher.TextMatrix(flxVoucher.Rows - 2, 1) = ""
'            flxVoucher.TextMatrix(flxVoucher.Rows - 2, 2) = ""
'            flxVoucher.TextMatrix(flxVoucher.Rows - 2, 3) = ""
'            flxVoucher.TextMatrix(flxVoucher.Rows - 2, 4) = ""
'            flxVoucher.TextMatrix(flxVoucher.Rows - 2, 5) = ""
'            flxVoucher.TextMatrix(flxVoucher.Rows - 2, 6) = ""
'        End If
    
        DoEvents
        flxVoucher.SetFocus
        flxVoucher.col = 2
        flxVoucher.row = flxVoucher.row + 1
    End If

Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->mnuADD_Click()")
End Sub

Public Sub GetSPCRRRecord(pSPCRRNO As String)
Dim mySPCRRValidate As Long
Dim myRFPNO As String, validateAns As String
On Error GoTo er
    
    mySPCRRValidate = 0
    myRFPNO = ""
    
    Qry = "SELECT supinv.SPRRNO,supinv.PACKINGNO,supinv.INVOICENO,supinv.TERMS,supinv.DUEDATE,supinv.INVSTATUS,supinv.INVOICETTLPAY,supinv.DATERCVD,rec.POREF as PONO " & _
          "FROM supplierinvoice as supinv " & _
          "LEFT JOIN spreceive as rec ON rec.PACKINGNO = supinv.PACKINGNO " & _
          "LEFT JOIN supplier as supp ON supp.SUPPLIERNAME = rec.SUPPLIER " & _
          "WHERE supinv.SPCRRNO = '" & pSPCRRNO & "' and supp.SUPPLIERID = " & Val(supplierid) & " Group by supinv.SPCRRNO "
    With tmpRs
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            If !INVSTATUS = "PENDING" Then
                MsgBox "SPCRR number status was still pending!", vbExclamation, Me.Caption
                Screen.MousePointer = vbDefault
                .Close
                GoTo SKIP
            End If
            If flxRFP.TextMatrix(DtlsRow, 3) <> "" Then
                mySPCRRValidate = Val(GetFieldValue("SELECT RFPID FROM rfp_dtls WHERE SPCRRNO = '" & pSPCRRNO & "' and INVOICENO = '" & flxRFP.TextMatrix(DtlsRow, 3) & "' LIMIT 1 "))
            Else
                mySPCRRValidate = Val(GetFieldValue("SELECT RFPID FROM rfp_dtls WHERE SPCRRNO = '" & pSPCRRNO & "' and INVOICENO = '" & IIf(IsNull(!InvoiceNO), "", !InvoiceNO) & "' LIMIT 1 "))
            End If
            If mySPCRRValidate > 0 Then
                myRFPNO = GetFieldValue("SELECT RFPNO FROM rfp_main WHERE RFPID = " & Val(mySPCRRValidate) & " and RFPSTATUS = 'FINAL' LIMIT 1")
                If myRFPNO <> "" Then
                    validateAns = MsgBox("This SPCRR Number is consignee? Do you want to proceed?", vbQuestion + vbYesNo, Me.Caption)
                    If validateAns = vbNo Then
                        MsgBox "SPCRR Number and Invoice Number is already encoded in RFP#:" & myRFPNO & "!", vbExclamation, Me.Caption
                        Screen.MousePointer = vbDefault
                        consign = False
                        .Close
                        GoTo SKIP
                    End If
                    consign = True
                End If
            End If
            While Not tmpRs.EOF
                If txtTerm.Text <> "CBD" Then
                    If formAction = "EDIT" Then
                        flxRFP.TextMatrix(flxRFP.row, 3) = IIf(IsNull(!InvoiceNO), flxRFP.TextMatrix(flxRFP.row, 3), !InvoiceNO)
                        flxRFP.TextMatrix(flxRFP.row, 4) = IIf(IsNull(!PACKINGNO), flxRFP.TextMatrix(flxRFP.row, 4), !PACKINGNO)
                        flxRFP.TextMatrix(flxRFP.row, 5) = IIf(IsNull(!DATERCVD), flxRFP.TextMatrix(flxRFP.row, 5), Format(!DATERCVD, "MM/DD/YYYY"))
                        flxRFP.TextMatrix(flxRFP.row, 6) = IIf(IsNull(!PONO), flxRFP.TextMatrix(flxRFP.row, 6), !PONO)
                        If txtTerm.Text <> "CBD" Then
                            flxRFP.TextMatrix(flxRFP.row, 7) = IIf(IsNull(!INVOICETTLPAY), Format(flxRFP.TextMatrix(flxRFP.row, 7), "#,###,###,###.00"), Format(!INVOICETTLPAY, "#,###,##0.00"))
                            If Val(cntDuplicate) = 0 Then
                                flxRFP.TextMatrix(flxRFP.row, 8) = IIf(IsNull(!INVOICETTLPAY), Format(flxRFP.TextMatrix(flxRFP.row, 8), "#,###,###,###.00"), Format(!INVOICETTLPAY, "#,###,##0.00"))
                            End If
                            flxRFP.TextMatrix(flxRFP.row, 9) = IIf(IsNull(!DueDate), flxRFP.TextMatrix(flxRFP.row, 9), Format(!DueDate, "MM/DD/YYYY"))
                            flxRFP.TextMatrix(flxRFP.row, 10) = "Y"
                        
                        Else
                            flxRFP.TextMatrix(flxRFP.row, 8) = IIf(IsNull(!INVOICETTLPAY), Format(flxRFP.TextMatrix(flxRFP.row, 7), "#,###,###,###.00"), Format(!INVOICETTLPAY, "#,###,##0.00"))
                            If Val(cntDuplicate) = 0 Then
                                flxRFP.TextMatrix(flxRFP.row, 9) = IIf(IsNull(!INVOICETTLPAY), Format(flxRFP.TextMatrix(flxRFP.row, 8), "#,###,###,###.00"), Format(!INVOICETTLPAY, "#,###,##0.00"))
                            End If
                            flxRFP.TextMatrix(flxRFP.row, 10) = IIf(IsNull(!DueDate), flxRFP.TextMatrix(flxRFP.row, 9), Format(!DueDate, "MM/DD/YYYY"))
                            flxRFP.TextMatrix(flxRFP.row, 12) = "Y"
                        End If
                    Else
                        If !InvoiceNO = "" Then
                            flxRFP.TextMatrix(flxRFP.row, 3) = flxRFP.TextMatrix(flxRFP.row, 3)
                        Else
                            flxRFP.TextMatrix(flxRFP.row, 3) = IIf(IsNull(!InvoiceNO), "", !InvoiceNO)
                        End If
                        flxRFP.TextMatrix(flxRFP.row, 4) = IIf(IsNull(!PACKINGNO), "", !PACKINGNO)
                        flxRFP.TextMatrix(flxRFP.row, 5) = IIf(IsNull(!DATERCVD), "", Format(!DATERCVD, "MM/DD/YYYY"))
                        flxRFP.TextMatrix(flxRFP.row, 6) = IIf(IsNull(!PONO), flxRFP.TextMatrix(flxRFP.row, 6), !PONO)
                        If txtTerm.Text <> "CBD" Then
                            flxRFP.TextMatrix(flxRFP.row, 7) = IIf(IsNull(!INVOICETTLPAY), "0.00", Format(!INVOICETTLPAY, "#,###,##0.00"))
                            If Val(cntDuplicate) = 0 Then
                                flxRFP.TextMatrix(flxRFP.row, 8) = IIf(IsNull(!INVOICETTLPAY), "0.00", Format(!INVOICETTLPAY, "#,###,##0.00"))
                            End If
                            flxRFP.TextMatrix(flxRFP.row, 9) = IIf(IsNull(!DueDate), "", Format(!DueDate, "MM/DD/YYYY"))
                            flxRFP.TextMatrix(flxRFP.row, 10) = "Y"
                        Else
                            flxRFP.TextMatrix(flxRFP.row, 8) = IIf(IsNull(!INVOICETTLPAY), "0.00", Format(!INVOICETTLPAY, "#,###,##0.00"))
                            If Val(cntDuplicate) = 0 Then
                                flxRFP.TextMatrix(flxRFP.row, 9) = IIf(IsNull(!INVOICETTLPAY), "0.00", Format(!INVOICETTLPAY, "#,###,##0.00"))
                            End If
                            flxRFP.TextMatrix(flxRFP.row, 10) = IIf(IsNull(!DueDate), "", Format(!DueDate, "MM/DD/YYYY"))
                            flxRFP.TextMatrix(flxRFP.row, 12) = "Y"
                        End If
                    End If
                Else
                    If formAction = "EDIT" Then
                        flxRFP.TextMatrix(flxRFP.row, 3) = IIf(IsNull(!InvoiceNO), flxRFP.TextMatrix(flxRFP.row, 3), !InvoiceNO)
                        flxRFP.TextMatrix(flxRFP.row, 5) = IIf(IsNull(!PACKINGNO), flxRFP.TextMatrix(flxRFP.row, 5), !PACKINGNO)
                        flxRFP.TextMatrix(flxRFP.row, 6) = IIf(IsNull(!DATERCVD), flxRFP.TextMatrix(flxRFP.row, 6), Format(!DATERCVD, "MM/DD/YYYY"))
                        flxRFP.TextMatrix(flxRFP.row, 7) = IIf(IsNull(!PONO), flxRFP.TextMatrix(flxRFP.row, 7), !PONO)
                    
                         flxRFP.TextMatrix(flxRFP.row, 9) = IIf(IsNull(!INVOICETTLPAY), Format(flxRFP.TextMatrix(flxRFP.row, 9), "#,###,###,###.00"), Format(!INVOICETTLPAY, "#,###,##0.00"))
                         If Val(cntDuplicate) = 0 Then
                             flxRFP.TextMatrix(flxRFP.row, 10) = IIf(IsNull(!INVOICETTLPAY), Format(flxRFP.TextMatrix(flxRFP.row, 10), "#,###,###,###.00"), Format(!INVOICETTLPAY, "#,###,##0.00"))
                         End If
                         flxRFP.TextMatrix(flxRFP.row, 11) = IIf(IsNull(!DueDate), flxRFP.TextMatrix(flxRFP.row, 11), Format(!DueDate, "MM/DD/YYYY"))
                         flxRFP.TextMatrix(flxRFP.row, 12) = "Y"
                    Else
                        If !InvoiceNO = "" Then
                            flxRFP.TextMatrix(flxRFP.row, 3) = flxRFP.TextMatrix(flxRFP.row, 3)
                        Else
                            flxRFP.TextMatrix(flxRFP.row, 3) = IIf(IsNull(!InvoiceNO), "", !InvoiceNO)
                        End If
                        flxRFP.TextMatrix(flxRFP.row, 5) = IIf(IsNull(!PACKINGNO), "", !PACKINGNO)
                        flxRFP.TextMatrix(flxRFP.row, 6) = IIf(IsNull(!DATERCVD), "", Format(!DATERCVD, "MM/DD/YYYY"))
                        flxRFP.TextMatrix(flxRFP.row, 7) = IIf(IsNull(!PONO), flxRFP.TextMatrix(flxRFP.row, 7), !PONO)
                     
                         flxRFP.TextMatrix(flxRFP.row, 9) = IIf(IsNull(!INVOICETTLPAY), "0.00", Format(!INVOICETTLPAY, "#,###,##0.00"))
                         If Val(cntDuplicate) = 0 Then
                             flxRFP.TextMatrix(flxRFP.row, 10) = IIf(IsNull(!INVOICETTLPAY), "0.00", Format(!INVOICETTLPAY, "#,###,##0.00"))
                         End If
                         flxRFP.TextMatrix(flxRFP.row, 11) = IIf(IsNull(!DueDate), "", Format(!DueDate, "MM/DD/YYYY"))
                         flxRFP.TextMatrix(flxRFP.row, 12) = "Y"
                    End If
                End If
                .MoveNext
            Wend
            flxRFP.TextMatrix(DtlsRow, 1) = txtFlxValueRFP.Text
        Else
            MsgBox "NO RECORD FOUND!", vbInformation, Me.Caption
            Screen.MousePointer = vbDefault
            tmpRs.Close
        End If
    End With

SKIP:
    'txtFlxValueRFP.Text = ""
    txtFlxValueRFP.Visible = False
    Screen.MousePointer = vbDefault
    
Exit Sub
er:
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->GetSPCRRRecord()")
End Sub


Private Sub ComputeTotalNetPay()
On Error GoTo er
    Dim i As Integer
    Dim j As Integer
    Dim TotalInvoice As Double
    Dim TotalDeduct As Double
    Dim TotalAdd As Currency
    Dim TotalLess As Currency

    TotalAdd = 0
    TotalLess = 0
    For i = 1 To flxRFP.Rows - 1
        If txtTerm.Text = "CBD" And RECON = False Then
            TotalInvoice = TotalInvoice + Val(Replace(flxRFP.TextMatrix(i, 8), ",", "")) 'po amount
        ElseIf txtTerm.Text = "CBD" And RECON = True Then
            TotalInvoice = TotalInvoice + Val(Replace(flxRFP.TextMatrix(i, 9), ",", "")) 'invoice amount
        Else
            TotalInvoice = TotalInvoice + Val(Replace(flxRFP.TextMatrix(i, 7), ",", "")) 'invoice amount
        End If
    Next
    
    txtTotInvAmnt.Text = Format(TotalInvoice, "#,##0.00")
    
    For j = 1 To flxDeduct.Rows - 1
        TotalDeduct = TotalDeduct + Val(Replace(flxDeduct.TextMatrix(j, 6), ",", ""))
    Next
    TotalAdd = Val(Replace(txtTotAdd.Text, ",", ""))
    TotalLess = Val(Replace(txtTotLess.Text, ",", ""))
    
    TotalDeduct = TotalLess - TotalAdd
    txtTotalDed.Text = Format(TotalDeduct, "#,##0.00")
    
    txtNetPay.Text = Format((TotalInvoice - TotalLess) + TotalAdd, "#,###,##0.00")
    txtTtlPayable.Text = Format((TotalInvoice - TotalLess) + TotalAdd, "#,###,##0.00")
    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->ComputeTotalNetPay()")
End Sub

Sub ClearEntryForm(Focus1stObject As Boolean)
    
    cboSupName.Text = ""
    cboRFPType.Text = ""
    cboCompany.Text = ""
    
    txtPayee.Text = ""
    txtRFPNo.Text = ""
    
    txtRFPDate.Mask = ""
    txtRFPDate.Text = ""
    txtRFPDate.Mask = "##/##/####"
    
    txtTtlPayable.Text = ""
    txtTotInvAmnt.Text = ""
    txtTotLess.Text = ""
    txtTotAdd.Text = ""
    txtTotalDed.Text = ""
    txtNetPay.Text = ""
    txtCheckedBy.Text = ""
    txtPrepBy.Text = ""
    txtNotedBy.Text = ""
    txtApprvdBy.Text = ""
    txtRemarks.Text = ""
    txtCheckNo.Text = ""
    
    txtChckNoDate.Mask = ""
    txtChckNoDate.Text = ""
    txtChckNoDate.Mask = "##/##/####"
    
    flxRFP.Rows = 1
    flxRFP.Rows = 2
    
    flxDeduct.Rows = 1
    flxDeduct.Rows = 2
    
    flxVoucher.Rows = 1
    flxVoucher.Rows = 2
    
    lblTotCredit.Caption = "0.00"
    lblTotDebit.Caption = "0.00"
    lblHeadingRecon.Caption = ""
    
    txtRemarks.Text = ""
    pTransID = 0
    
    myRFPID = 0
    
    If Focus1stObject = True Then cmdTrans(0).SetFocus
                              
End Sub

Sub LockedEntryObject(DataEntry As Boolean)

    DataEntry = Not DataEntry
    
        cboSupName.Locked = DataEntry
        txtPayee.Locked = DataEntry
        txtRFPNo.Locked = DataEntry
        txtCheckNo.Locked = DataEntry
        cboRFPType.Locked = DataEntry
        cboTrans.Locked = DataEntry
        txtTerm.Locked = DataEntry
        txtRemarks.Locked = DataEntry
        cboCompany.Locked = DataEntry
    If cboSupName.Locked Then
    
        cboSupName.BackColor = txtHoldGrayBackColor.BackColor
        cboRFPType.BackColor = txtHoldGrayBackColor.BackColor
        txtPayee.BackColor = txtHoldGrayBackColor.BackColor
        txtRFPNo.BackColor = txtHoldGrayBackColor.BackColor
        txtRFPDate.BackColor = txtHoldGrayBackColor.BackColor
        txtCheckNo.BackColor = txtHoldGrayBackColor.BackColor
        txtChckNoDate.BackColor = txtHoldGrayBackColor.BackColor
        txtTerm.BackColor = txtHoldGrayBackColor.BackColor
        txtRemarks.BackColor = txtHoldGrayBackColor.BackColor
        cboTrans.BackColor = txtHoldGrayBackColor.BackColor
        txtDocDateRec.BackColor = txtHoldGrayBackColor.BackColor
        cboCompany.BackColor = txtHoldGrayBackColor.BackColor
        txtDocDateRec.Enabled = False
        txtRFPDate.Enabled = False
        txtChckNoDate.Enabled = False
        chkFinal.Enabled = False
    Else
        
        cboSupName.BackColor = txtHoldWhiteBackColor.BackColor
        cboRFPType.BackColor = txtHoldWhiteBackColor.BackColor
        txtPayee.BackColor = txtHoldWhiteBackColor.BackColor
        txtRFPNo.BackColor = txtHoldWhiteBackColor.BackColor
        txtTerm.BackColor = txtHoldWhiteBackColor.BackColor
        txtRFPDate.BackColor = txtHoldWhiteBackColor.BackColor
        txtRemarks.BackColor = txtHoldWhiteBackColor.BackColor
        cboTrans.BackColor = txtHoldWhiteBackColor.BackColor
'        txtCheckNo.BackColor = txtHoldWhiteBackColor.BackColor
'        txtChckNoDate.BackColor = txtHoldWhiteBackColor.BackColor
        txtRFPDate.Enabled = True
        chkFinal.Enabled = True
        txtDocDateRec.Enabled = True
        txtDocDateRec.BackColor = txtHoldWhiteBackColor.BackColor
'        txtChckNoDate.Enabled = True
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
    For i = 0 To 7
        cmdTrans(i).Visible = False
    Next i
    
    If DataEntry Then
        cmdTrans(5).Visible = True
        cmdTrans(6).Visible = True
    Else
        For i = 0 To 4
            cmdTrans(i).Visible = True
        Next i
        cmdTrans(7).Visible = True
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


    With cboSupName
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
    Call modObjTools.ErrorMessage(Err.Description, "Error ->cboSupName_Change()")
End Sub

Private Sub cboSupName_Click()
If didRecordEntry = False Then Exit Sub

    If cboSupName.ListIndex <> -1 Then
        supplierid = GetFieldValue("SELECT SUPPLIERID FROM SUPPLIER WHERE SUPPLIERNAME = '" & cboSupName.Text & "' ")
        txtTerm.SetFocus
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
        If cboSupName.Text <> "" Then
            supplierid = Val(GetFieldValue("SELECT SUPPLIERID FROM SUPPLIER WHERE SUPPLIERNAME = '" & cboSupName.Text & "' "))
            txtTerm.SetFocus
        End If
    End If
End Sub

Private Sub cboSupName_LostFocus()
    Call cboSupName_KeyPress(13)
End Sub

Private Sub cmdTrans_Click(Index As Integer)
Dim reconVal As String
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
    Dim message As String, mySysHead As String, myRepHead As String
    
    If Index = 0 Then 'NEW
        RECON = False
        formAction = "NEW"
        didRecordEntry = True
        lblHeadingRecon.Visible = False
        
        Call SetButtonVisibility(True)
        Call LockedEntryObject(True)
        Call ClearEntryForm(False)
        
        Qry = "Select `VALUE` From settings Where DESCRIPTION = 'REQUEST FOR PAYMENT' and REMARKS = 'CHECKEDBY' "
        txtCheckedBy.Text = GetFieldValue(Qry)
        
        Qry = "Select `VALUE` From settings Where DESCRIPTION = 'REQUEST FOR PAYMENT' and REMARKS = 'PREPAREDBY' "
        txtPrepBy.Text = GetFieldValue(Qry)
        
'        Qry = "Select `VALUE` From settings Where DESCRIPTION = 'REQUEST FOR PAYMENT' and REMARKS = 'DIVISION HEAD' "
'        txtNotedBy.Text = GetFieldValue(Qry) ----' Remove as per request from acctg from Charissa to Judith 03/12/2024
        
        Qry = "Select `VALUE` From settings Where DESCRIPTION = 'REQUEST FOR PAYMENT' and REMARKS = 'NOTEDBY' "
        txtNotedBy.Text = GetFieldValue(Qry)
    
        Qry = "Select `VALUE` From settings Where DESCRIPTION = 'REQUEST FOR PAYMENT' and REMARKS = 'APPROVEDBY' "
        txtApprvdBy.Text = GetFieldValue(Qry)
        
        Qry = "Select IFNULL(MAX(RFPID),0)+1 as RFPID From rfp_main"
        txtRFPNo.Text = Year(Now()) & "-" & GetFieldValue(Qry)
        
        txtPayee.SetFocus
    ElseIf Index = 1 Then 'EDIT
        
        If Val(chkFinal.Value) = 1 Then
            MsgBox "This Transaction was already Final!.", vbInformation + vbOKOnly, "Warning"
            Exit Sub
        End If
        RECON = False
        formAction = "EDIT"
        didRecordEntry = True
        Call SetButtonVisibility(True)
        Call LockedEntryObject(True)
        cboSupName.BackColor = txtHoldGrayBackColor.BackColor
        cboSupName.Locked = True
        txtPayee.SetFocus
    ElseIf Index = 2 Then 'PRINT
        Dim QryDeductDtls As String, QryTreasury As String, QryAcctTitles As String
        If txtRFPNo.Text = "" Then
            MsgBox "No Records Found", vbExclamation + vbOKOnly, "PRINT"
            Exit Sub
        End If
        Qry = "Select `VALUE` From settings Where DESCRIPTION = 'REQUEST FOR PAYMENT' and REMARKS = 'CHECKEDBY-POSITION' "
        checkedByPosition = GetFieldValue(Qry)
        
        Qry = "Select `VALUE` From settings Where DESCRIPTION = 'REQUEST FOR PAYMENT' and REMARKS = 'NOTEDBY-POSITION' "
        notedByPosition = GetFieldValue(Qry)
        
        Qry = "Select `VALUE` From settings Where DESCRIPTION = 'REQUEST FOR PAYMENT' and REMARKS = 'PREPAREDBY-POSITION' "
        preparedByPosition = GetFieldValue(Qry)
        
        Qry = "Select `VALUE` From settings Where DESCRIPTION = 'REQUEST FOR PAYMENT' and REMARKS = 'APPROVEDBY-POSITION' "
        approvedByposition = GetFieldValue(Qry)
        
        Qry = "call sp_acctpayinvcng.sp_load_rfp_info('" & txtRFPNo.Text & "')"
        frmPrint.QryPrn = Qry
        QryDeductDtls = "Call sp_acctpayinvcng.sp_load_rfp_deduction(" & Val(myRFPID) & ")"
        frmPrint.QryPrn2 = QryDeductDtls
        QryTreasury = "call sp_acctpayinvcng.sp_load_rfp_treasury('" & txtRFPNo.Text & "')"
        frmPrint.QryPrn3 = QryTreasury
        QryAcctTitles = "SELECT b.BRANCHNAME, ch.`TRANSACTION` as TransNAme,vou.code,vou.Description,vou.Type,vou.Amount,vou.TotDebit, vou.TotCredit FROM voucherdtls as vou " & _
                        "LEFT JOIN mrcwarehouse.branch as b ON b.BRANCHCODE = vou.BCODE INNER JOIN accounting_corp.trans_legend as ch ON vou.TransID = ch.TRANS_LEGEND_ID WHERE vou.RFPID = " & Val(myRFPID) & " GROUP by vou.SeqNo Order by vou.SeqNo asc;"
        frmPrint.QryPrn4 = QryAcctTitles
        
        If cboRFPType.Text = "DECLARED" Then
            If Year(Now()) = "2019" Then
                mySysHead = "MOTO-R REPUBLIC CORPORATION (WAREHOUSE)"
            Else
                mySysHead = "LYR MOTORPARTS (RETAIL)"
            End If
        Else
            If Year(Now()) = "2019" Then
                mySysHead = "MOTO-R REPUBLIC CORPORATION (WAREHOUSE)"
            Else
                mySysHead = "LYR MOTORPARTS (RETAIL)"
            End If
        End If
        
        If Val(chkFinal) = 1 Then
            myRepHead = "RECONCILIATION"
        Else
            myRepHead = "REQUEST FOR PAYMENT"
        End If
        
        frmPrint.ParamValue1 = cboCompany.Text 'mySysHead
        frmPrint.ParamValue2 = Format(txtTotLess.Text, "#,###,###,###.00")
        frmPrint.ParamValue3 = Format(txtTotAdd.Text, "#,###,###,###.00")
        frmPrint.ParamValue4 = Format(txtTotalDed.Text, "#,###,###,###.00")
        frmPrint.ParamValue5 = Format(txtNetPay.Text, "#,###,###,###.00")
        frmPrint.ParamValue6 = myRepHead
        frmPrint.ParamValue7 = preparedByPosition
        frmPrint.ParamValue8 = checkedByPosition
        frmPrint.ParamValue9 = notedByPosition
        frmPrint.ParamValue10 = approvedByposition
        
        If txtTerm.Text <> "CBD" Then
            frmPrint.MySource = "RFP"
        Else
            frmPrint.MySource = "RFP CBD"
        End If
        frmPrint.Show

    ElseIf Index = 3 Then 'SUMMARY
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
        If formAction = "NEW" Then

            ans = MsgBox("PLEASE MAKE SURE THAT THE DATA IS CORRECT!" & _
                vbCrLf & _
                vbCrLf & "DO YOU WISH TO CONTINUE?", vbYesNo + vbQuestion, Me.Caption)
            If ans = vbNo Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            Qry = "call sp_acctpayinvcng.sp_rfp_create('" & Trim(txtRFPNo.Text) & "','" & Format(txtRFPDate.Text, "YYYY-MM-DD") & "'," & Val(supplierid) & "," & Val(chkFinal.Value) & ",'" & Trim(txtPayee.Text) & "','" & Trim(txtTerm.Text) & "', " & _
                  "'" & Trim(cboRFPType.Text) & "','" & Trim(txtCheckedBy.Text) & "','" & Trim(txtNotedBy.Text) & "','" & Trim(txtApprvdBy.Text) & "'," & Val(UserID) & ",'" & Trim(txtRemarks.Text) & "'," & Val(Replace(txtTotInvAmnt.Text, ",", "")) & ", " & _
                  "" & Val(Replace(txtTotalDed.Text, ",", "")) & "," & Val(Replace(txtNetPay.Text, ",", "")) & ", '" & cboCompany.Text & "','" & Format(txtDocDateRec.Text, "YYYY-MM-DD") & "')"

            With tmpRs
                If .State = adStateOpen Then .Close
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    If Trim(!level) = "" Then
                        Exit Sub
                    ElseIf !level = "Granted" Or !level = "Warning" Then
                        myRFPID = !RFPID
                        If !level = "Granted" Then
                            MsgBox !message, vbInformation, "SAVE."
                        Else
                            MsgBox !message, vbExclamation, "SAVE!"
                            Exit Sub
                        End If
                    Else
                        MsgBox message, vbCritical, "Error!"
                        Exit Sub
                    End If
                
                End If
            End With
        ElseIf formAction = "EDIT" Then
            ans = MsgBox("DO YOU WANT TO SAVE YOUR CHANGES?", vbQuestion + vbYesNo, Me.Caption)
            If ans = vbNo Then Exit Sub
            
            If RECON = True Then
                reconVal = "Y"
                lblHeadingRecon.Visible = True
                lblHeadingRecon.Caption = "RECONCILED"
            Else
                reconVal = "N"
                lblHeadingRecon.Visible = False
            End If
            
            Qry = "call sp_acctpayinvcng.sp_rfp_update(" & Val(myRFPID) & ",'" & Trim(txtRFPNo.Text) & "','" & Format(txtRFPDate.Text, "YYYY-MM-DD") & "'," & Val(supplierid) & "," & Val(chkFinal.Value) & ",'" & Trim(txtPayee.Text) & "','" & Trim(txtTerm.Text) & "', " & _
                  "'" & Trim(cboRFPType.Text) & "','" & Trim(txtCheckedBy.Text) & "','" & Trim(txtNotedBy.Text) & "','" & Trim(txtApprvdBy.Text) & "'," & Val(UserID) & ",'" & Trim(txtRemarks.Text) & "'," & Val(Replace(txtTotInvAmnt.Text, ",", "")) & ", " & _
                  "" & Val(Replace(txtTotalDed.Text, ",", "")) & "," & Val(Replace(txtNetPay.Text, ",", "")) & ",'" & reconVal & "','" & cboCompany.Text & "','" & Format(txtDocDateRec.Text, "YYYY-MM-DD") & "')"
            With tmpRs
                If .State = adStateOpen Then .Close
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    If Trim(!level) = "" Then
                        Exit Sub
                    ElseIf !level = "Granted" Or !level = "Warning" Then
                        If !level = "Granted" Then
                            MsgBox !message, vbInformation, "SAVE."
                        Else
                            MsgBox !message, vbExclamation, "SAVE!"
                            Exit Sub
                        End If
                    Else
                        MsgBox message, vbCritical, "Error!"
                        Exit Sub
                    End If
                
                End If
            End With
        End If
        Screen.MousePointer = vbDefault

        Call LockedEntryObject(False)
        Call SetButtonVisibility(False)
        Call EnabledSelectedButton

        didRecordEntry = False

        formAction = ""

        If reconVal = "Y" And txtTerm.Text = "CBD" Then
            lblHeadingRecon.Visible = True
            lblHeadingRecon.Caption = "RECONCILED"
        Else
            lblHeadingRecon.Visible = False
            lblHeadingRecon.Caption = "UNRECONCILED"
        End If
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
    ElseIf Index = 7 Then 'RECONCILIATION Button
        
        If Val(chkFinal.Value) = 0 And txtTerm.Text = "CBD" Then
            MsgBox "This Transaction is not Final!.", vbInformation + vbOKOnly, "Warning"
            Exit Sub
        End If
        
        If txtTerm.Text = "CBD" And Val(chkFinal.Value) = 1 And lblHeadingRecon.Caption <> "RECONCILED" Then
            ans = MsgBox("Do you want to process the Reconciliation?", vbYesNo + vbQuestion, Me.Caption)
            If ans = vbNo Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
    
            RECON = True
            formAction = "EDIT"
            didRecordEntry = True
            Call SetButtonVisibility(True)
            Call LockedEntryObject(True)
            cboSupName.BackColor = txtHoldGrayBackColor.BackColor
            cboSupName.Locked = True
            txtPayee.SetFocus
        Else
            MsgBox "This Transaction is already Reconciled", vbInformation + vbOKOnly, "Validation"
            Exit Sub
        End If
    End If
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdTrans_Click()")
End Sub

Private Sub flxRFP_Click()
    flexGridObject = "flxRFP"
End Sub

Private Sub flxRFP_DblClick()
On Error GoTo er

    If didRecordEntry = False Then Exit Sub
   
    DtlsCol = flxRFP.col
    DtlsRow = flxRFP.row
    If txtTerm.Text <> "CBD" Then
        If DtlsCol = 1 Or DtlsCol = 3 Or DtlsCol = 6 Or DtlsCol = 7 Then
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
            
            If DtlsCol = 1 Then
                .Alignment = 0
            Else
                .Alignment = 2
            End If
          End With
        ElseIf DtlsCol = 2 Then
           With txtFlxDate
            .Top = framHoldMid.Top + flxRFP.CellTop + flxRFP.Top
            .Left = framHoldMid.Left + flxRFP.CellLeft + flxRFP.Left
            .Width = flxRFP.CellWidth
            .Height = flxRFP.CellHeight
            If flxRFP.TextMatrix(flxRFP.row, 2) <> "" Then .Text = flxRFP.TextMatrix(flxRFP.row, DtlsCol)
            .Visible = True
            .BorderStyle = 1
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
          End With
        End If
    Else 'CBD
        If DtlsCol = 4 Or DtlsCol = 5 Or DtlsCol = 7 Or DtlsCol = 8 And RECON = False Then 'Add Or DtlsCol = 5 -- 25.12.03
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
            
            If DtlsCol = 1 Then
                .Alignment = 0
            Else
                .Alignment = 2
            End If
          End With
        ElseIf RECON = True Then
            If DtlsCol = 1 Or DtlsCol = 3 Or DtlsCol = 5 Or DtlsCol = 7 Or DtlsCol = 8 Or DtlsCol = 9 Then 'Add Or DtlsCol = 5 -- 25.12.03
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
                
                If DtlsCol = 1 Then
                    .Alignment = 0
                Else
                    .Alignment = 2
                End If
              End With
            ElseIf DtlsCol = 2 Then
               With txtFlxDate
                .Top = framHoldMid.Top + flxRFP.CellTop + flxRFP.Top
                .Left = framHoldMid.Left + flxRFP.CellLeft + flxRFP.Left
                .Width = flxRFP.CellWidth
                .Height = flxRFP.CellHeight
                If flxRFP.TextMatrix(flxRFP.row, 2) <> "" Then .Text = flxRFP.TextMatrix(flxRFP.row, DtlsCol)
                .Visible = True
                .BorderStyle = 1
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
              End With
            End If
        End If
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
    DoEvents
    Qry = "SELECT supplier.SUPPLIERNAME, supplier.SUPPLIERID From supplier Where supplier.SUPPLIERNAME not like '%INACTIVE%' Order By supplier.SUPPLIERNAME;"
    With tmpRs
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
                .MoveFirst
                If cboSupName.Text <> "" Then MySupp = cboSupName.Text
                cboSupName.Clear
                Do While Not .EOF
                    cboSupName.AddItem !SUPPLIERNAME
                    cboSupName.ItemData(cboSupName.ListCount - 1) = !supplierid
                    .MoveNext
                Loop
        End If
        .Close
        If MySupp <> "" Then cboSupName.Text = MySupp
    End With
    
    Qry = "SELECT DISTINCT DEDUCTREMARKS FROM rfp_deduct WHERE DEDUCTREMARKS IS NOT NULL"
    With tmpRs
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
                .MoveFirst
                cboFlxRemarks.Clear
                Do While Not .EOF
                    cboFlxRemarks.AddItem !DEDUCTREMARKS
                    .MoveNext
                Loop
        End If
        .Close
    End With
    
    cboFlxParticular.AddItem "DEBIT MEMO"
    cboFlxParticular.AddItem "CREDIT MEMO"
    cboFlxParticular.AddItem "VAT(12%)"
    cboFlxParticular.AddItem "EXPANDED WITH HOLDING TAX 1%"
    
    DoEvents
    Call LoadCboRecord
    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->Form_Load()")
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
        '.Clear
        .Rows = 2
        .Cols = 10
        .WordWrap = True
        .AllowUserResizing = flexResizeBoth
        
        '               0     1          2              3           4             5             6             7             8                 9         10
        .FormatString = " |<SPCRR |<INVOICE DATE |^INVOICE NO |^PACKING NO |^DATE RECEIVED  |^P.O NO. |^INVOICE AMOUNT |^SPCRR AMOUNT |^DUE DATE "
        
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.1
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.1
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.11
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.1
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.1
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.11
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.13
        .ColWidth(8) = (.Width - .ColWidth(0)) * 0.13
        .ColWidth(9) = (.Width - .ColWidth(0)) * 0.1
    End With
    
    With flxDeduct
        '.Clear
        .Rows = 2
        .Cols = 7
        '               0       1           2           3           4          5            6
        .FormatString = " |^REFERENCE |<DESCRIPTION  |^TYPE  |<PARTICULARS |^REMARKS |>DEDUCTION/REB                "
        
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.12
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.4
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.08
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.16
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.1
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.12
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
    
End Sub

Private Sub tmrGetRecord_Timer()
On Error GoTo er
    If tmrGetRecord.Enabled = False Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    tmrGetRecord.Enabled = False
    
    DoEvents
    Call ShowRecord(txtRFPNo.Text)
    
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

Private Sub txtFlxDate_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(txtFlxValueRFP, KeyAscii, ValidCodeChars)
    If KeyAscii = 13 Then
        
        If txtFlxDate.Text <> "" Then
            flxRFP.TextMatrix(DtlsRow, 2) = txtFlxDate.Text
            flxRFP.SetFocus
            flxRFP.col = 3
            flxRFP.row = DtlsRow
        End If
    End If
End Sub

Private Sub txtFlxDate_LostFocus()
Call txtFlxDate_KeyPress(13)

txtFlxDate.Visible = False
txtFlxDate.Text = "__/__/____"
End Sub

Private Sub txtFlxDeductReb_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFlxDeductReb, KeyAscii, NumbersWithDecimalPoint)
    If KeyAscii = 13 Then
        If flxDeduct.TextMatrix(flxDeduct.row, 3) = "LESS" Then
            TtlLessValue = TtlLessValue + Val(txtFlxDeductReb.Text)
        Else
            TtlAddValue = TtlAddValue + Val(txtFlxDeductReb.Text)
        End If
        flxDeduct.TextMatrix(flxDeduct.row, 6) = Format(txtFlxDeductReb.Text, "#,###,###,##0.00")
        flxDeduct.SetFocus
        txtFlxDeductReb.Visible = False
        Call ComputeTotAddLess
        Call ComputeTotalNetPay
    End If
End Sub


Private Sub txtFlxDeductReb_LostFocus()
txtFlxDeductReb.Text = ""
End Sub

Private Sub txtFlxDesc_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFlxDesc, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        flxDeduct.TextMatrix(flxDeduct.row, 2) = txtFlxDesc.Text
        flxDeduct.SetFocus
        If flxDeduct.TextMatrix(flxDeduct.row, 2) = "" Then
            flxDeduct.col = 2
            flxDeduct.row = DtlsRow
        Else
            flxDeduct.col = 3
            flxDeduct.row = DtlsRow
        End If
        
    End If
End Sub

Private Sub txtFlxDesc_LostFocus()
    Call txtFlxDesc_KeyPress(13)
    txtFlxDesc.Visible = False
    txtFlxDesc.Text = ""
End Sub

Private Sub txtFlxReference_KeyPress(KeyAscii As Integer)
Dim cntExists As Long
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFlxReference, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        If txtFlxReference.Text <> "" Then
            For i = 1 To flxRFP.Rows - 1
                If txtFlxReference.Text = flxRFP.TextMatrix(i, 3) Then
                    cntExists = cntExists + 1
                    GoTo SKIP
                End If
            Next i
            If cntExists = 0 Then
                MsgBox "Invoice Number not found in Request For Payment Details!", vbInformation + vbOKOnly, "Validation"
                flxDeduct.SetFocus
                Exit Sub
            End If
        End If
        
SKIP:
        flxDeduct.TextMatrix(flxDeduct.row, 1) = txtFlxReference.Text
        flxDeduct.SetFocus
        If flxDeduct.TextMatrix(flxDeduct.row, 1) = "" Then
            flxDeduct.col = 1
            flxDeduct.row = DtlsRow
        Else
            flxDeduct.col = 2
            flxDeduct.row = DtlsRow
        End If
    End If
End Sub

Private Sub txtFlxReference_LostFocus()
    Call txtFlxReference_KeyPress(13)
    txtFlxReference.Visible = False
    txtFlxReference.Text = ""
End Sub

Private Sub txtFlxValueRFP_KeyPress(KeyAscii As Integer)
Dim chckExistsRFID As Long, POSupplierID As Long
Dim chckExistsRFPNO As String, validateAns As String

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(txtFlxValueRFP, KeyAscii, ValidCodeChars)
    If KeyAscii = 13 Then
        POSupplierID = 0
        If txtTerm.Text <> "CBD" Then
            If DtlsCol = 1 And txtFlxValueRFP.Text <> "" Then 'SPCRRNO
                consign = False
                Call ZeroSPCRRAMNT
                Call GetSPCRRRecord(Trim(txtFlxValueRFP.Text))
                Call ComputeTotalNetPay
                flxRFP.SetFocus
                If flxRFP.TextMatrix(DtlsRow, 1) = "" Then
                    flxRFP.col = 1
                    flxRFP.row = DtlsRow
                Else
                    flxRFP.col = 2
                    flxRFP.row = DtlsRow
                End If
            ElseIf DtlsCol = 3 Then 'INVOICE NO
                If consign = False Then
                    chckExistsRFID = 0
                    chckExistsRFID = Val(GetFieldValue("SELECT RFPID FROM rfp_dtls WHERE SPCRRNO = '" & flxRFP.TextMatrix(DtlsRow, 1) & "' AND INVOICENO = '" & txtFlxValueRFP.Text & "' "))
                    If chckExistsRFID > 0 Then
                        chckExistsRFPNO = GetFieldValue("SELECT RFPNO FROM rfp_main WHERE RFPID = " & Val(chckExistsRFID) & " and RFPSTATUS IN ('FINAL','PENDING') LIMIT 1")
                        If chckExistsRFPNO <> "" Then
                            MsgBox "SPCRR Number and Invoice Number is already encoded in RFP#:" & chckExistsRFPNO & "!", vbExclamation, Me.Caption
                            txtFlxValueRFP.Text = ""
                            txtFlxValueRFP.Visible = False
                            flxRFP.TextMatrix(DtlsRow, 1) = ""
                            flxRFP.TextMatrix(DtlsRow, 2) = ""
                            flxRFP.TextMatrix(DtlsRow, 3) = ""
                            flxRFP.TextMatrix(DtlsRow, 4) = ""
                            flxRFP.TextMatrix(DtlsRow, 5) = ""
                            flxRFP.TextMatrix(DtlsRow, 6) = ""
                            flxRFP.TextMatrix(DtlsRow, 7) = ""
                            flxRFP.TextMatrix(DtlsRow, 8) = ""
                            flxRFP.TextMatrix(DtlsRow, 9) = ""
                            Exit Sub
                        End If
                    End If
                End If
                flxRFP.TextMatrix(DtlsRow, 3) = txtFlxValueRFP.Text
                flxRFP.SetFocus
                If flxRFP.TextMatrix(DtlsRow, 3) = "" Then
                    flxRFP.col = 3
                    flxRFP.row = DtlsRow
                Else
                    flxRFP.col = 6
                    flxRFP.row = DtlsRow
                End If
            ElseIf DtlsCol = 6 Then 'PO NO
  
                POSupplierID = Val(GetFieldValue("Select SUPPLIERID From sporder Where REFERENCE = '" & txtFlxValueRFP.Text & "' "))
                
                If Val(supplierid) <> Val(POSupplierID) Then
                    MsgBox "This PO Number is not applicable to this Supplier!", vbInformation + vbOKOnly, "Validation"
                    txtFlxValueRFP.Text = ""
                    Exit Sub
                End If
                flxRFP.TextMatrix(DtlsRow, 6) = txtFlxValueRFP.Text
                flxRFP.SetFocus
                If flxRFP.TextMatrix(DtlsRow, 6) = "" Then
                    flxRFP.col = 6
                    flxRFP.row = DtlsRow
                Else
                    flxRFP.col = 7
                    flxRFP.row = DtlsRow
                End If
            ElseIf DtlsCol = 7 Then 'INVOICE AMOUNT
                flxRFP.TextMatrix(DtlsRow, 7) = Format(txtFlxValueRFP.Text, "#,###,###,###.00")
                flxRFP.SetFocus
                If flxRFP.TextMatrix(DtlsRow, 7) = "" Then
                    flxRFP.col = 7
                    flxRFP.row = DtlsRow
                End If
                Call mnuAdd_Click
                Call ComputeTotalNetPay
            End If
        Else
            If DtlsCol = 1 And txtFlxValueRFP.Text <> "" Then 'SPCRRNO
                consign = False
                Call ZeroSPCRRAMNT
                Call GetSPCRRRecord(Trim(txtFlxValueRFP.Text))
                Call ComputeTotalNetPay
                flxRFP.SetFocus
                If flxRFP.TextMatrix(DtlsRow, 1) = "" Then
                    flxRFP.col = 1
                    flxRFP.row = DtlsRow
                Else
                    flxRFP.col = 2
                    flxRFP.row = DtlsRow
                End If
            ElseIf DtlsCol = 3 Then 'INVOICE NO
                If consign = False Then
                    chckExistsRFID = 0
                    chckExistsRFID = Val(GetFieldValue("SELECT RFPID FROM rfp_dtls WHERE SPCRRNO = '" & flxRFP.TextMatrix(DtlsRow, 1) & "' AND INVOICENO = '" & txtFlxValueRFP.Text & "' "))
                    If chckExistsRFID > 0 Then
                        chckExistsRFPNO = GetFieldValue("SELECT RFPNO FROM rfp_main WHERE RFPID = " & Val(chckExistsRFID) & " and RFPSTATUS = 'FINAL' LIMIT 1")
                        If chckExistsRFPNO <> "" Then
                            MsgBox "SPCRR Number and Invoice Number is already encoded in RFP#:" & chckExistsRFPNO & "!", vbExclamation, Me.Caption
                            txtFlxValueRFP.Text = ""
                            txtFlxValueRFP.Visible = False
                            flxRFP.TextMatrix(DtlsRow, 1) = ""
                            flxRFP.TextMatrix(DtlsRow, 2) = ""
                            flxRFP.TextMatrix(DtlsRow, 3) = ""
                            flxRFP.TextMatrix(DtlsRow, 4) = ""
                            flxRFP.TextMatrix(DtlsRow, 5) = ""
                            flxRFP.TextMatrix(DtlsRow, 6) = ""
                            flxRFP.TextMatrix(DtlsRow, 7) = ""
                            flxRFP.TextMatrix(DtlsRow, 8) = ""
                            flxRFP.TextMatrix(DtlsRow, 9) = ""
                            Exit Sub
                        End If
                    End If
                End If
                flxRFP.TextMatrix(DtlsRow, 3) = txtFlxValueRFP.Text
                flxRFP.SetFocus
                If flxRFP.TextMatrix(DtlsRow, 3) = "" Then
                    flxRFP.col = 3
                    flxRFP.row = DtlsRow
                Else
                    flxRFP.col = 6
                    flxRFP.row = DtlsRow
                End If
            ElseIf DtlsCol = 4 Then 'REF NO

                flxRFP.TextMatrix(DtlsRow, 4) = txtFlxValueRFP.Text
                flxRFP.SetFocus
                If flxRFP.TextMatrix(DtlsRow, 4) = "" Then
                    flxRFP.col = 4
                    flxRFP.row = DtlsRow
                Else
                    flxRFP.col = 5
                    flxRFP.row = DtlsRow
                End If
            ElseIf DtlsCol = 5 Then 'PACKING NO -- 25.12.03
            
                flxRFP.TextMatrix(DtlsRow, 5) = txtFlxValueRFP.Text
                flxRFP.SetFocus
                If flxRFP.TextMatrix(DtlsRow, 5) = "" Then
                    flxRFP.col = 5
                    flxRFP.row = DtlsRow
                Else
                    flxRFP.col = 7
                    flxRFP.row = DtlsRow
                End If
            ElseIf DtlsCol = 7 Then 'PO NO

                POSupplierID = Val(GetFieldValue("Select SUPPLIERID From sporder Where REFERENCE = '" & txtFlxValueRFP.Text & "' "))
                
                If Val(supplierid) <> Val(POSupplierID) Then
                    MsgBox "This PO Number is not applicable to this Supplier!", vbInformation + vbOKOnly, "Validation"
                    txtFlxValueRFP.Text = ""
                    Exit Sub
                End If
                
                flxRFP.TextMatrix(DtlsRow, 7) = txtFlxValueRFP.Text
                flxRFP.SetFocus
                If flxRFP.TextMatrix(DtlsRow, 7) = "" Then
                    flxRFP.col = 7
                    flxRFP.row = DtlsRow
                Else
                    flxRFP.col = 8
                    flxRFP.row = DtlsRow
                End If
            ElseIf DtlsCol = 8 Then 'PO AMOUNT
                flxRFP.TextMatrix(DtlsRow, 8) = Format(txtFlxValueRFP.Text, "#,###,###,###.00")
                flxRFP.SetFocus
                If flxRFP.TextMatrix(DtlsRow, 8) = "" Then
                    flxRFP.col = 8
                    flxRFP.row = DtlsRow
                End If
                If RECON = False Then
                    Call mnuAdd_Click
                    Call ComputeTotalNetPay
                End If
            End If
            If RECON = True Then
                If DtlsCol = 9 Then 'INVOICE AMOUNT
                    flxRFP.TextMatrix(DtlsRow, 9) = Format(txtFlxValueRFP.Text, "#,###,###,###.00")
                    flxRFP.SetFocus
                    If flxRFP.TextMatrix(DtlsRow, 9) = "" Then
                    flxRFP.col = 9
                    flxRFP.row = DtlsRow
                    End If
                    Call mnuAdd_Click
                    Call ComputeTotalNetPay
                End If
            End If
        End If
        If flxRFP.row <> flxRFP.Rows - 1 Then
            flxRFP.row = DtlsRow + 1
        End If
        txtFlxValueRFP.Visible = False
    End If
End Sub

Private Sub txtFlxValueRFP_LostFocus()
txtFlxValueRFP.Visible = False
End Sub


Private Sub txtPayee_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtPayee, KeyAscii, ValidChars)
    If KeyAscii = 13 Then cboSupName.SetFocus
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtRemarks, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtPayee.SetFocus
End Sub

Private Sub txtRFPDate_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(txtRFPDate, KeyAscii, ValidChars)
    If KeyAscii = 13 Then flxRFP.SetFocus
End Sub

Private Sub txtRFPNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtRFPNo, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtRFPDate.SetFocus
End Sub

Private Sub txtTerm_Change()
On Error GoTo er
    Dim i As Integer, j As Integer
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
    'If didRecordEntry = False Then Exit Sub
    
    If SearchcboSup Then
        SearchcboSup = False
        Exit Sub
    End If


    With txtTerm
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
    Call modObjTools.ErrorMessage(Err.Description, "Error ->txtTerm_Change()")
End Sub

Private Sub txtTerm_Click()
Call txtTerm_KeyPress(13)
End Sub

Private Sub txtTerm_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtTerm, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Call flxCBDOpt
        cboRFPType.SetFocus
    End If
End Sub
