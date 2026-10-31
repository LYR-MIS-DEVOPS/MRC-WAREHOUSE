VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSIMSDiscrepancy 
   Caption         =   "Processing of Branch Discrepancy"
   ClientHeight    =   9150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   13785
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtWHSDiscID 
      Height          =   285
      Left            =   13920
      TabIndex        =   50
      Top             =   3000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtRem 
      Height          =   285
      Left            =   13920
      TabIndex        =   49
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   13920
      TabIndex        =   48
      Top             =   2280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtPARTNO 
      Height          =   285
      Left            =   13920
      TabIndex        =   47
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSDiscID 
      Height          =   285
      Left            =   13920
      TabIndex        =   46
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   13920
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   13920
      TabIndex        =   44
      Text            =   "HoldBackColor"
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtRecom 
      Height          =   285
      Left            =   13920
      TabIndex        =   43
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      Height          =   1215
      Left            =   120
      TabIndex        =   36
      Tag             =   "XY"
      Top             =   7920
      Width           =   13575
      Begin VB.TextBox txtApprvBy 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9120
         TabIndex        =   39
         Tag             =   "XY"
         Top             =   360
         Width           =   4215
      End
      Begin VB.TextBox txtPrepBy 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   38
         Tag             =   "XY"
         Top             =   360
         Width           =   4215
      End
      Begin VB.TextBox txtAudBy 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4680
         TabIndex        =   37
         Tag             =   "XY"
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "APPROVED BY"
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
         Left            =   10680
         TabIndex        =   42
         Tag             =   "XY"
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblAuditBy 
         Caption         =   "AUDITED BY"
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
         Left            =   6240
         TabIndex        =   41
         Tag             =   "XY"
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblPrepBy 
         Caption         =   "PREPARED BY"
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
         Left            =   1680
         TabIndex        =   40
         Tag             =   "XY"
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   120
      TabIndex        =   29
      Tag             =   "XY"
      Top             =   6960
      Width           =   13575
      Begin VB.TextBox txtAuditFindings 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5280
         TabIndex        =   35
         Tag             =   "XY"
         Top             =   550
         Width           =   8175
      End
      Begin VB.TextBox txtRemarks 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5280
         TabIndex        =   33
         Tag             =   "XY"
         Top             =   150
         Width           =   8175
      End
      Begin VB.TextBox txtStatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   31
         Tag             =   "XY"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblFindings 
         Alignment       =   1  'Right Justify
         Caption         =   "AUDIT FINDINGS:"
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
         Left            =   3600
         TabIndex        =   34
         Tag             =   "XY"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblRemarks 
         Alignment       =   1  'Right Justify
         Caption         =   "REMARKS:"
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
         Left            =   3600
         TabIndex        =   32
         Tag             =   "XY"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblStatus 
         Alignment       =   1  'Right Justify
         Caption         =   "STATUS:"
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
         Left            =   240
         TabIndex        =   30
         Tag             =   "XY"
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1500
      Left            =   120
      TabIndex        =   9
      Tag             =   "XY"
      Top             =   960
      Width           =   13575
      Begin VB.ComboBox cmbBranch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   17
         Tag             =   "XY"
         Top             =   480
         Width           =   5295
      End
      Begin VB.TextBox txtSTFNo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   16
         Tag             =   "XY"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtRefNo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11640
         TabIndex        =   15
         Tag             =   "XY"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtRRNo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8640
         TabIndex        =   14
         Tag             =   "XY"
         Top             =   960
         Width           =   1815
      End
      Begin VB.CheckBox chkFinal 
         Caption         =   "FINAL"
         Height          =   255
         Left            =   5640
         TabIndex        =   10
         Tag             =   "XY"
         Top             =   960
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtDateRef 
         Height          =   360
         Left            =   11640
         TabIndex        =   11
         Tag             =   "XY"
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
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
      Begin MSMask.MaskEdBox txtDateRec 
         Height          =   360
         Left            =   11640
         TabIndex        =   12
         Tag             =   "XY"
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
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
      Begin MSMask.MaskEdBox txtSTFDate 
         Height          =   360
         Left            =   3840
         TabIndex        =   13
         Tag             =   "XY"
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
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
      Begin VB.Label lblBranch 
         Caption         =   "SOURCE:"
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
         Left            =   240
         TabIndex        =   24
         Tag             =   "XY"
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblSTFNo 
         Caption         =   "STF #:"
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
         Left            =   240
         TabIndex        =   23
         Tag             =   "XY"
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblSTFDate 
         Alignment       =   1  'Right Justify
         Caption         =   "STF DATE:"
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
         Left            =   2640
         TabIndex        =   22
         Tag             =   "XY"
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblRefNo 
         Alignment       =   1  'Right Justify
         Caption         =   "REF. #:"
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
         Left            =   10440
         TabIndex        =   21
         Tag             =   "XY"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblDateRef 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE REF:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10440
         TabIndex        =   20
         Tag             =   "XY"
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblDateRec 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE REC:"
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
         Left            =   10440
         TabIndex        =   19
         Tag             =   "XY"
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblRRNo 
         Alignment       =   1  'Right Justify
         Caption         =   "RR #:"
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
         Left            =   8040
         TabIndex        =   18
         Tag             =   "XY"
         Top             =   960
         Width           =   495
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&PROCESS S.D"
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
      Image           =   "frmSIMSDiscrepancy.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   1
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
      Image           =   "frmSIMSDiscrepancy.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4620
      TabIndex        =   2
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
      Image           =   "frmSIMSDiscrepancy.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6120
      TabIndex        =   3
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
      Image           =   "frmSIMSDiscrepancy.frx":10CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   7620
      TabIndex        =   4
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
      Image           =   "frmSIMSDiscrepancy.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   9120
      TabIndex        =   5
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
      Image           =   "frmSIMSDiscrepancy.frx":1C02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1620
      TabIndex        =   6
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
      Image           =   "frmSIMSDiscrepancy.frx":219C
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame2 
      Height          =   4575
      Left            =   120
      TabIndex        =   27
      Tag             =   "XY"
      Top             =   2400
      Width           =   13575
      Begin MSFlexGridLib.MSFlexGrid flxSIMSDiscrepancy 
         Height          =   4215
         Left            =   120
         TabIndex        =   28
         Tag             =   "XY"
         Top             =   240
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   7435
         _Version        =   393216
         Cols            =   8
         ScrollBars      =   2
         FormatString    =   "|<PARTNUMBER |<DESCRIPTION |^QTY |<REMARKS |<BRANCH RECOMMENDATION |<SIM RECOMMENDATION"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   11160
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14025
      m_MDI_Height    =   9720
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "SIM DEPARTMENT"
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
      Left            =   150
      TabIndex        =   25
      Top             =   720
      Width           =   3210
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "SETTLEMENT OF DISCREPANCY"
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
      Index           =   0
      Left            =   150
      TabIndex        =   7
      Top             =   480
      Width           =   3210
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "SETTLEMENT OF DISCREPANCY"
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
      Left            =   120
      TabIndex        =   8
      Top             =   510
      Width           =   3090
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "SIM DEPARTMENT"
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
      Index           =   3
      Left            =   120
      TabIndex        =   26
      Top             =   750
      Width           =   3090
   End
   Begin VB.Menu mnuStatus 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add Row"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete Row"
      End
   End
End
Attribute VB_Name = "frmSIMSDiscrepancy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ans As String
Dim Qry As String
Dim arrBRANCH() As String
Dim i As Long, j As Long
Dim ActRow As Integer, ActCol As Integer
Dim tmpRs  As New ADODB.Recordset
Dim mainRs As New ADODB.Recordset
Dim StatSrch As Boolean
Dim strPartial As String, strTotal As String
Dim formAction As String, bCode As String
Dim PrtBln As Boolean, PrtSrch As Boolean, ProperExit As Boolean
Dim chkStock As String, ChkRec As String, globalStatus As String

Public Sub CreateTempTable()
On Error GoTo er
    LYRcon.Execute "Drop Temporary Table IF  Exists sp_mrcwarehouse.TMP_SIMSDISCREPANCYDTL;"

    DoEvents
    LYRcon.Execute " Create Temporary Table IF Not Exists sp_mrcwarehouse.TMP_SIMSDISCREPANCYDTL( " & _
        "   QTY                 INT(11),        " & _
        "   REMARKS             Varchar(40),    " & _
        "   BRECOMMENDATION     Varchar(100),   " & _
        "   SRECOMMENDATION     Varchar(100),   " & _
        "   SPAREPARTID         INT(11),        " & _
        "   SEQNO               INT(11)         " & _
        ");"

        Qry = "Insert Into sp_mrcwarehouse.TMP_SIMSDISCREPANCYDTL(QTY,REMARKS,BRECOMMENDATION,SRECOMMENDATION,SPAREPARTID,SEQNO) VALUES "
        
        For i = 1 To flxSIMSDiscrepancy.Rows - 1
            If i > 1 Then Qry = Qry & ","
            Qry = Qry & "('" & flxSIMSDiscrepancy.TextMatrix(i, 3) & "', '" & flxSIMSDiscrepancy.TextMatrix(i, 4) & "', '" & flxSIMSDiscrepancy.TextMatrix(i, 5) & "', '" & flxSIMSDiscrepancy.TextMatrix(i, 6) & "', '" & flxSIMSDiscrepancy.TextMatrix(i, 7) & "', " & i & ")"
        Next
        Call LYRcon.Execute(Qry)
Exit Sub
er:
 Call modObjTools.ErrorMessage(Err.Description, "Error->CreateTempTable()")
End Sub

'Public Sub UpdateSDiscOnline(mySTATUS As String, MyBranchCode As String, mySDiscID As Long)
'    Call ConnectOnlineBC(MyBranchCode)
'
'    If mySTATUS = "NEW" Then
'        If MyBranchCode = "MRT" Then
'            OnlineCon.Execute "UPDATE sdiscrepancy SET `STATUS` = 'VIEWED' Where sdiscrepancy.SDISCID = " & Val(mySDiscID) & " AND sdiscrepancy.`STATUS` = 'NEW';"
'        ElseIf MyBranchCode = "MCPR" Then
'            OnlineCon.Execute "UPDATE mcpdrizalonline.sdiscrepancy SET `STATUS` = 'VIEWED' Where sdiscrepancy.SDISCID = " & Val(mySDiscID) & " AND sdiscrepancy.`STATUS` = 'NEW';"
'        ElseIf MyBranchCode = "BSPT" Then
'            OnlineCon.Execute "UPDATE mrcbinostagumonline.sdiscrepancy SET `STATUS` = 'VIEWED' Where sdiscrepancy.SDISCID = " & Val(mySDiscID) & " AND sdiscrepancy.`STATUS` = 'NEW';"
'        ElseIf MyBranchCode = "MG" Then
'            OnlineCon.Execute "UPDATE mrcgensanonline.sdiscrepancy SET `STATUS` = 'VIEWED' Where sdiscrepancy.SDISCID = " & Val(mySDiscID) & " AND sdiscrepancy.`STATUS` = 'NEW';"
'        ElseIf MyBranchCode = "YMK" Then
'            OnlineCon.Execute "UPDATE mrcmarbelonline.sdiscrepancy SET `STATUS` = 'VIEWED' Where sdiscrepancy.SDISCID = " & Val(mySDiscID) & " AND sdiscrepancy.`STATUS` = 'NEW';"
'        ElseIf MyBranchCode = "MMD" Then
'            OnlineCon.Execute "UPDATE mrcmatinaonline.sdiscrepancy SET `STATUS` = 'VIEWED' Where sdiscrepancy.SDISCID = " & Val(mySDiscID) & " AND sdiscrepancy.`STATUS` = 'NEW';"
'        ElseIf MyBranchCode = "YMT" Then
'            OnlineCon.Execute "UPDATE mrcmotorhubtagumonline.sdiscrepancy SET `STATUS` = 'VIEWED' Where sdiscrepancy.SDISCID = " & Val(mySDiscID) & " AND sdiscrepancy.`STATUS` = 'NEW';"
'        ElseIf MyBranchCode = "MP" Then
'            OnlineCon.Execute "UPDATE mrcpanaboonline.sdiscrepancy SET `STATUS` = 'VIEWED' Where sdiscrepancy.SDISCID = " & Val(mySDiscID) & " AND sdiscrepancy.`STATUS` = 'NEW';"
'        ElseIf MyBranchCode = "BSP" Then
'            OnlineCon.Execute "UPDATE mrcpolomolokonline.sdiscrepancy SET `STATUS` = 'VIEWED' Where sdiscrepancy.SDISCID = " & Val(mySDiscID) & " AND sdiscrepancy.`STATUS` = 'NEW';"
'        ElseIf MyBranchCode = "BBS" Then
'            OnlineCon.Execute "UPDATE mrcsamalonline.sdiscrepancy SET `STATUS` = 'VIEWED' Where sdiscrepancy.SDISCID = " & Val(mySDiscID) & " AND sdiscrepancy.`STATUS` = 'NEW';"
'        ElseIf MyBranchCode = "BST" Then
'            OnlineCon.Execute "UPDATE mrcstotomasonline.sdiscrepancy SET `STATUS` = 'VIEWED' Where sdiscrepancy.SDISCID = " & Val(mySDiscID) & " AND sdiscrepancy.`STATUS` = 'NEW';"
'        ElseIf MyBranchCode = "MPT" Then
'            OnlineCon.Execute "UPDATE sqlmcpdepotonline.sdiscrepancy SET `STATUS` = 'VIEWED' Where sdiscrepancy.SDISCID = " & Val(mySDiscID) & " AND sdiscrepancy.`STATUS` = 'NEW';"
'        ElseIf MyBranchCode = "MST" Then
'            OnlineCon.Execute "UPDATE sqlmtrtgmonline.sdiscrepancy SET `STATUS` = 'VIEWED' Where sdiscrepancy.SDISCID = " & Val(mySDiscID) & " AND sdiscrepancy.`STATUS` = 'NEW';"
'        ElseIf MyBranchCode = "MUD" Then
'            OnlineCon.Execute "UPDATE sqlmtrulasonline.sdiscrepancy SET `STATUS` = 'VIEWED' Where sdiscrepancy.SDISCID = " & Val(mySDiscID) & " AND sdiscrepancy.`STATUS` = 'NEW';"
'        End If
'    End If
'End Sub

Function ValidateEntry() As Boolean
  ValidateEntry = False
    
    If Trim(cmbBranch.Text) = "" Then
        MsgBox "BRANCH IS REQUIRED!", vbExclamation, "Validation"
        cmbBranch.SetFocus
        Exit Function
    End If
    
    If Trim(txtSTFNO.Text) = "" Then
        MsgBox "STF # IS REQUIRED!", vbExclamation, "Validation"
        txtSTFNO.SetFocus
        Exit Function
    End If
        
    If Trim(txtSTFDate.Text) = "__/__/____" Then
        MsgBox "STF DATE IS REQUIRED!", vbExclamation, "Validation"
        txtSTFDate.SetFocus
        Exit Function
    End If
    
    If Trim(txtRRno.Text) = "" Then
        MsgBox "RECEIVING RECEIPT # IS REQUIRED!", vbExclamation, "Validation"
        txtRRno.SetFocus
        Exit Function
    End If
    
    If Trim(txtDateRec.Text) = "__/__/____" Then
        MsgBox "DATE RECEIVE IS REQUIRED!", vbExclamation, "Validation"
        txtDateRec.SetFocus
        Exit Function
    End If
    
    If Trim(txtPrepBy.Text) = "" Then
        MsgBox "PREPARED BY IS REQUIRED!", vbExclamation, "Validation"
        txtPrepBy.SetFocus
        Exit Function
    End If
    
    If Trim(txtAudBy.Text) = "" Then
        MsgBox "AUDITED BY IS REQUIRED!", vbExclamation, "Validation"
        txtPrepBy.SetFocus
        Exit Function
    End If
    
    If Trim(txtApprvby.Text) = "" Then
        MsgBox "APPROVED BY IS REQUIRED!", vbExclamation, "Validation"
        txtApprvby.SetFocus
        Exit Function
    End If
    
    ValidateEntry = True

End Function



Private Sub cmbBranch_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cmbBranch
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmbBranch-Change()")
End Sub

Private Sub cmbBranch_Click()
    bCode = arrBRANCH(cmbBranch.ListIndex, 0)
End Sub

Private Sub cmbBranch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cmbBranch_KeyPress(KeyAscii As Integer)
Dim cbolistIndex As Integer

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
   ' Call ObjKeyTrapping.Validation(cmbBranch, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        cbolistIndex = modObjTools.getListIndex(cmbBranch, cmbBranch.Text)
        
        If cbolistIndex < 0 Then
            cmbBranch.SetFocus
        Else
            cmbBranch.ListIndex = cbolistIndex
            txtSTFNO.SetFocus
        End If
        
        Call cmbBranch_Click
        
    End If
End Sub

Private Sub cmdTrans_Click(Index As Integer)
On Error GoTo er

    If Index = 0 Then 'add
        If txtStatus.Text = "NEW" Then
            Call LockData(True)
            Call SetButtonVisibility(True)
            chkFinal.Enabled = True
            chkFinal.Value = 0
            formAction = "add"
        Else
            Exit Sub
        End If
    ElseIf Index = 1 Then 'edit
        If txtStatus.Text = "FINAL" Then
            MsgBox "This Transaction is already Final", vbExclamation + vbOKOnly, "Edit"
            Exit Sub
        ElseIf txtStatus.Text = "NEW" Then
            MsgBox "Cannot Update because this Transaction is New", vbExclamation + vbOKOnly, "Edit"
            Exit Sub
        End If
        Call SetButtonVisibility(True)
        Call LockData(True)
        cmbBranch.BackColor = txtHoldGrayBackColor.BackColor
        cmbBranch.Locked = True
        If txtStatus.Text = "PENDING" Then
            chkFinal.Enabled = True
            chkFinal.Value = 0
        End If
        formAction = "edit"
    ElseIf Index = 2 Then 'view summary
        FrmSummary.mySrc = "SETTLEMENT DISCREPANCY"
        FrmSummary.Show
        Unload Me
    ElseIf Index = 3 Then
        If txtRefNo.Text = "<AUTO>" Then
            MsgBox "No Records Found", vbExclamation + vbOKOnly, "PRINT"
            Exit Sub
        End If

        Qry = "call sp_mrcwarehouse.sp_WHSDiscrepancy_getDetails('" & sysBranchCode & "'," & Val(txtSDiscID.Text) & ",'" & globalStatus & "', '" & bCode & "')"
        frmPrint.QryPrn = Qry

        frmPrint.ParamValue1 = SysBranch
        frmPrint.ParamValue2 = Format(cmbBranch.Text, "")
        frmPrint.ParamValue3 = Format(txtSTFNO.Text, "")
        frmPrint.ParamValue4 = Format(txtSTFDate.Text, "mm-dd-yyyy")
        frmPrint.ParamValue5 = Format(txtRefNo.Text, "")
        frmPrint.ParamValue6 = Format(txtDateRef.Text, "mm-dd-yyyy")
        frmPrint.ParamValue7 = Format(txtDateRec.Text, "mm-dd-yyyy")
        frmPrint.ParamValue8 = Format(txtRRno.Text, "")
        frmPrint.ParamValue9 = Format(txtStatus.Text, "")
        frmPrint.ParamValue10 = Format(txtRemarks.Text, "")
        frmPrint.ParamValue11 = Format(txtPrepBy.Text, "")
        frmPrint.ParamValue12 = Format(txtApprvby.Text, "")
        frmPrint.ParamValue13 = Format(SysAddress, "")
        frmPrint.ParamValue14 = Format(txtAudBy, "")
        frmPrint.ParamValue15 = Format(txtAuditFindings, "")
        frmPrint.MySource = "SETTLEMENT DISCREPANCY"
        frmPrint.Show
    ElseIf Index = 4 Then 'exit
        formAction = "EXIT"
        ProperExit = True
        Call clearform
        formAction = ""
        Unload Me
    ElseIf Index = 5 Then 'save
        If ValidateEntry() = False Then Exit Sub

        ans = MsgBox("Do you want to save?", vbQuestion + vbYesNo, "Save")

        If ans = vbNo Then Exit Sub

        If formAction = "add" Then
            Call CreateTempTable
            DoEvents
            Qry = "call sp_mrcwarehouse.sp_WHSDiscrepancy_Create('" & sysBranchCode & "', '" & cmbBranch.Text & "', '" & txtSTFNO.Text & "', '" & Format(txtSTFDate.Text, "YYYY-mm-dd") & "', '" & txtRRno.Text & "', '" & Format(txtDateRec.Text, "YYYY-mm-dd") & "', '" & Format(txtDateRef.Text, "YYYY-mm-dd") & "', '" & txtRemarks.Text & "', '" & txtAuditFindings.Text & "', " & UserID & ", '" & txtAudBy.Text & "', '" & txtApprvby.Text & "', " & UserID & ", " & Val(chkFinal.Value) & ", '" & bCode & "', " & Val(txtSDiscID.Text) & ")"
            With mainRs
                .Open Qry, LYRcon, adOpenKeyset, adLockReadOnly
                If .State <> 1 Then Exit Sub
                If .RecordCount > 0 Then
                    txtRefNo.Text = !RefNo
                    txtStatus.Text = !Status
                End If
                .Close
            End With
            'LYRcon.Execute (Qry)
            'txtRefNo.Text = GetFieldValue("SELECT SDISCREPANCY.REFNO FROM SDISCREPANCY WHERE SDISCREPANCY.SOURCE= '" & cmbBranch.Text & "' AND SDISCREPANCY.LASTUPDATE='" & Format(Now(), "yyyy-mm-dd hh:mm:ss") & "'")
            'txtStatus.Text = GetFieldValue("SELECT SDISCREPANCY.STATUS FROM SDISCREPANCY WHERE SDISCREPANCY.REFNO= '" & txtRefNo.Text & "'")
        Else
            Call CreateTempTable
            Qry = "call sp_mrcwarehouse.sp_WHSDiscrepancy_Update('" & sysBranchCode & "', '" & txtSTFNO.Text & "', '" & Format(txtSTFDate.Text, "YYYY-mm-dd") & "', '" & txtRRno.Text & "', '" & Format(txtDateRec.Text, "YYYY-mm-dd") & "', '" & txtRemarks.Text & "', '" & txtAuditFindings.Text & "', '" & txtAudBy.Text & "', '" & txtApprvby.Text & "', " & UserID & ", " & txtSDiscID.Text & ", " & Val(chkFinal.Value) & ", '" & bCode & "')"
            With mainRs
                .Open Qry, LYRcon, adOpenKeyset, adLockReadOnly
                If .State <> 1 Then Exit Sub
                If .RecordCount > 0 Then
                    txtStatus.Text = !Status
                End If
                .Close
            End With
            'LYRcon.Execute (Qry)
            'txtStatus.Text = GetFieldValue("SELECT SDISCREPANCY.STATUS FROM SDISCREPANCY WHERE SDISCREPANCY.REFNO= '" & txtRefNo.Text & "'")
        End If
        Call SetButtonVisibility(False)
        Call LockData(False)
        chkFinal.Enabled = False
        formAction = ""

    ElseIf Index = 6 Then 'cancel
        ans = MsgBox("Do you want to cancel?", vbQuestion + vbYesNo)
        If ans = vbNo Then Exit Sub
        Call SetButtonVisibility(False)
        Call LockData(False)
        chkFinal.Enabled = False
        formAction = ""
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdTrans-Click()")
End Sub

Private Sub flxSIMSDiscrepancy_Click()
    ActRow = flxSIMSDiscrepancy.row
    ActCol = flxSIMSDiscrepancy.Col
    
    'FlxParts Columns: 1-PartNumber, 2-Desc, 3-Qty, 4-Price, 5-Total, 6-ID, 7-Remarks
    If ActCol = 0 Or ActRow = 0 Then Exit Sub
    Select Case ActCol
'        Case 1  'PartNumber
'            If txtSTFNo.Text <> "" And txtRRNo.Text <> "" Then
'                With txtPARTNO
'                    If flxSIMSDiscrepancy.TextMatrix(ActRow, 1) <> "" Then .Tag = flxSIMSDiscrepancy.TextMatrix(ActRow, 1)
'                    .Top = Frame2.Top + flxSIMSDiscrepancy.CellTop + flxSIMSDiscrepancy.Top
'                    .Left = Frame2.Left + flxSIMSDiscrepancy.CellLeft + flxSIMSDiscrepancy.Left
'                    .Height = flxSIMSDiscrepancy.CellHeight
'                    .Width = flxSIMSDiscrepancy.CellWidth
'                    .Text = flxSIMSDiscrepancy.Text
'                    .Visible = True
'                    .SelStart = 0
'                    .SelLength = Len(.Text)
'                    .ZOrder
'                    .SetFocus
'                End With
'            Else
'                Exit Sub
'            End If
'        Case 3 'Qty.
'            If flxSIMSDiscrepancy.TextMatrix(flxSIMSDiscrepancy.Row, 1) <> "" Then
'                With txtQty
'                    If flxSIMSDiscrepancy.TextMatrix(ActRow, 3) <> "" Then .Tag = flxSIMSDiscrepancy.TextMatrix(ActRow, 3)
'                    .Top = Frame2.Top + flxSIMSDiscrepancy.CellTop + flxSIMSDiscrepancy.Top
'                    .Left = Frame2.Left + flxSIMSDiscrepancy.CellLeft + flxSIMSDiscrepancy.Left
'                    .Height = flxSIMSDiscrepancy.CellHeight
'                    .Width = flxSIMSDiscrepancy.CellWidth
'                    .Text = flxSIMSDiscrepancy.Text
'                    .Visible = True
'                    .SelStart = 0
'                    .SelLength = Len(.Text)
'                    .ZOrder
'                    .SetFocus
'                End With
'            Else
'                Exit Sub
'            End If
'        Case 4  'REMARKS
'            If flxSIMSDiscrepancy.TextMatrix(flxSIMSDiscrepancy.Row, 3) <> "" Then
'                With txtRem
'                    If flxSIMSDiscrepancy.TextMatrix(ActRow, 4) <> "" Then .Tag = flxSIMSDiscrepancy.TextMatrix(ActRow, 4)
'                    .Top = Frame2.Top + flxSIMSDiscrepancy.CellTop + flxSIMSDiscrepancy.Top
'                    .Left = Frame2.Left + flxSIMSDiscrepancy.CellLeft + flxSIMSDiscrepancy.Left
'                    .Height = flxSIMSDiscrepancy.CellHeight
'                    .Width = flxSIMSDiscrepancy.CellWidth
'                    .Text = flxSIMSDiscrepancy.Text
'                    .Visible = True
'                    .SelStart = 0
'                    .SelLength = Len(.Text)
'                    .ZOrder
'                    .SetFocus
'                End With
'            Else
'                Exit Sub
'            End If
        Case 6  'SRECOMMENDATION
            If formAction = "add" Or formAction = "edit" Then
                With txtRecom
                    If flxSIMSDiscrepancy.TextMatrix(ActRow, 6) <> "" Then .Tag = flxSIMSDiscrepancy.TextMatrix(ActRow, 6)
                    .Top = Frame2.Top + flxSIMSDiscrepancy.CellTop + flxSIMSDiscrepancy.Top
                    .Left = Frame2.Left + flxSIMSDiscrepancy.CellLeft + flxSIMSDiscrepancy.Left
                    .Height = flxSIMSDiscrepancy.CellHeight
                    .Width = flxSIMSDiscrepancy.CellWidth
                    .Text = flxSIMSDiscrepancy.Text
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .ZOrder
                    .SetFocus
                End With
            Else
                Exit Sub
            End If
       End Select
End Sub

Private Sub Form_Load()
    Call LockData(False)
    txtDateRef.Text = Format$(Now, "mm/dd/yyyy")
    txtPrepBy.Text = EmployeesName
    txtRefNo.Text = "<AUTO>"
    chkFinal.Enabled = False
    
    '===================BUTTON POSITION===================
    If cmdTrans(0).Visible = False Then
        'Align Save Button
        cmdTrans(5).Left = cmdTrans(0).Left
        cmdTrans(5).Top = cmdTrans(0).Top
        'Align Cancel Button
        cmdTrans(6).Left = cmdTrans(1).Left
        cmdTrans(6).Top = cmdTrans(1).Top
    End If
End Sub

Sub clearform()

    'Frame 1
    cmbBranch.Text = ""
    txtSTFNO.Text = ""
    txtSTFDate.Text = "__/__/____"
    txtRefNo.Text = "<AUTO>"
    txtDateRef.Text = Format(Now, "mm/dd/yyyy")
    txtDateRec.Text = "__/__/____"
    txtRRno.Text = ""
   
    'Frame 2
    flxSIMSDiscrepancy.Rows = 1
    flxSIMSDiscrepancy.Rows = 2
    
    'Frame 3
    txtStatus.Text = ""
    txtRemarks.Text = ""
    txtAuditFindings.Text = ""
    
    'Frame 4
    txtApprvby.Text = ""
    txtAudBy.Text = ""
    
End Sub

Sub LockData(DataEntry As Boolean)
    DataEntry = Not DataEntry
    'Frame 1
    cmbBranch.Locked = DataEntry
    txtSTFNO.Locked = DataEntry
    'txtSTFDate.Locked = DataEntry
    txtRefNo.Locked = DataEntry
    'txtDateRef.Locked = DataEntry
    'txtDateRec.Locked = DataEntry
    txtRRno.Locked = DataEntry
    'Frame 3
    txtStatus.Locked = DataEntry
    txtRemarks.Locked = DataEntry
    txtAuditFindings.Locked = DataEntry
    'Frame 4
    txtPrepBy.Locked = DataEntry
    txtApprvby.Locked = DataEntry
    txtAudBy.Locked = DataEntry
    
    If cmbBranch.Locked Then
        'Frame 1
        cmbBranch.BackColor = txtHoldGrayBackColor.BackColor
        txtSTFNO.BackColor = txtHoldGrayBackColor.BackColor
        txtSTFDate.BackColor = txtHoldGrayBackColor.BackColor
        txtRefNo.BackColor = txtHoldGrayBackColor.BackColor
        txtDateRef.BackColor = txtHoldGrayBackColor.BackColor
        txtDateRec.BackColor = txtHoldGrayBackColor.BackColor
        txtRRno.BackColor = txtHoldGrayBackColor.BackColor
        'Frame 3
        txtStatus.BackColor = txtHoldGrayBackColor.BackColor
        txtRemarks.BackColor = txtHoldGrayBackColor.BackColor
        txtAuditFindings.BackColor = txtHoldGrayBackColor.BackColor
        'Frame 4
        txtPrepBy.BackColor = txtHoldGrayBackColor.BackColor
        txtApprvby.BackColor = txtHoldGrayBackColor.BackColor
        txtAudBy.BackColor = txtHoldGrayBackColor.BackColor
    Else
        'Frame 1
        cmbBranch.Locked = True
        txtSTFNO.Locked = True
        'txtSTFDate.Locked = True
        txtRefNo.Locked = True
        'txtDateRec.Locked = True
        txtRRno.Locked = True
        
        'Frame 3
        txtStatus.Locked = True
        txtRemarks.BackColor = txtHoldWhiteBackColor.BackColor
        txtAuditFindings.BackColor = txtHoldWhiteBackColor.BackColor
        
        'Frame 4
        txtPrepBy.Locked = True
        txtApprvby.BackColor = txtHoldWhiteBackColor.BackColor
        txtAudBy.BackColor = txtHoldWhiteBackColor.BackColor
    End If
End Sub

Private Sub SetButtonVisibility(DataEntry As Boolean)
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->VIEW SUMM
'3 ->EXIT
'4 ->PRINT
'5 ->SAVE
'6 ->CANCEL

    'hide all button
    For i = 0 To 5
        cmdTrans(i).Visible = False
    Next i

    If DataEntry Then
        cmdTrans(5).Visible = True
        cmdTrans(6).Visible = True
    Else
        For i = 0 To 5
            cmdTrans(i).Visible = True
            cmdTrans(5).Visible = False
            cmdTrans(6).Visible = False
        Next i
    End If
End Sub

Public Sub flxChanger()
    With flxSIMSDiscrepancy
        .Cols = 8
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.15   'PARTS NUMBER
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.2  'DESCRIPTION
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.1 'QTY
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.13 'REMARKS
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.2 'BRECOMMENDATION
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.2 'SRECOMMENDATION
        .ColWidth(7) = 0 '(.Width - .ColWidth(0)) * 0.06 'SPID
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo er

    If ProperExit Then
        If mainRs.State = adStateOpen Then mainRs.Close
        Call cmdTrans_Click(4)
        ProperExit = False
    Else
        Cancel = True
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->QueryUnload()")
End Sub

Private Sub Form_Resize()
Call RXResizer1_ResizeComplete
End Sub


Private Sub RXResizer1_ResizeComplete()
Call flxChanger
End Sub

Public Sub CboDisplay()
On Error GoTo er
    'Get Data from database
    DoEvents
    Call modPMBranchPicking.sp_get_Branch(mainRs)
    If mainRs.State = 1 Then
        Call modObjTools.addToListObjectWithArrayGivenField(cmbBranch, mainRs, arrBRANCH, "BRANCHCODE", "BRANCHNAME")
    End If
    If mainRs.State <> 0 Then mainRs.Close
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub

Private Sub txtApprvBy_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtApprvby, KeyAscii, ValidChars)
    If KeyAscii = 13 Then cmbBranch.SetFocus
End Sub


Private Sub txtAudBy_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtAudBy, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtApprvby.SetFocus
End Sub

Private Sub txtAuditFindings_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtAuditFindings, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtAudBy.SetFocus
End Sub

Private Sub txtRecom_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(txtSPECMODEL, KeyAscii, ValidChars)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    
    If KeyAscii = 13 Then
        Call txtRecom_LostFocus
    End If
End Sub

Private Sub txtRecom_LostFocus()
    If txtRecom.Tag = "skiplostfocus" Then
        txtRecom.Tag = ""
        Exit Sub
    End If
    If txtRecom.Text = "" Then
        txtRecom.Visible = False
    End If
    
    flxSIMSDiscrepancy.TextMatrix(ActRow, 6) = txtRecom.Text
    txtRecom.Visible = False
End Sub



Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtRemarks, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtAuditFindings.SetFocus
End Sub

Private Sub txtRRNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtRRno, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Call txtDateRec.SetFocus
    End If
End Sub

'Private Sub txtRRNo_LostFocus()
'Dim Qry As String, tmpDateRec As String, chkExists As String
'Dim rsRec As New ADODB.Recordset
'On Error GoTo er
'    tmpDateRec = ""
'
'    If txtRRNo.Text <> "" Then
'        tmpDateRec = Format(GetFieldValue("Select spreceive.DATEREC FROM spreceive where spreceive.PACKINGNO = '" & txtRRNo.Text & "'"), "mm/dd/YYYY")
'        If tmpDateRec = "" Then
'            txtDateRec.Text = "__/__/____"
'        Else
'            txtDateRec.Text = tmpDateRec
'        End If
'
'        chkExists = GetFieldValue("Select SDISCREPANCY.REFNO FROM SDISCREPANCY where SDISCREPANCY.RRNO = '" & txtRRNo.Text & "'")
'
'        If chkExists = "" Then
'            Qry = "call sp_mrcwarehouse.sp_WHSDiscrepancy_getReceive('" & sysBranchCode & "','" & BCODE & "', '" & txtRRNo.Text & "')"
'            With rsRec
'                .Open Qry, LYRcon, adOpenStatic, adLockOptimistic
'                If .State <> 1 Then Exit Sub
'                If .RecordCount > 0 Then
'                    .MoveFirst
'                    flxSIMSDiscrepancy.Rows = .RecordCount + 1
'
'                    For i = 1 To flxSIMSDiscrepancy.Rows - 1
'                        flxSIMSDiscrepancy.TextMatrix(i, 1) = !PartNumber
'                        flxSIMSDiscrepancy.TextMatrix(i, 2) = !Description
'                        flxSIMSDiscrepancy.TextMatrix(i, 3) = !qty
'                        flxSIMSDiscrepancy.TextMatrix(i, 4) = !Remarks
'                        flxSIMSDiscrepancy.TextMatrix(i, 5) = !BRECOMMENDATION
'                        flxSIMSDiscrepancy.TextMatrix(i, 7) = !SPAREPARTID
'                        .MoveNext
'                    Next i
'                Else
'                    MsgBox "No Record Found", vbInformation, "Information"
'                    txtRRNo.SetFocus
'                    Exit Sub
'                End If
'                .Close
'            End With
'            flxSIMSDiscrepancy.SetFocus
'            flxSIMSDiscrepancy.Row = 1
'            flxSIMSDiscrepancy.Col = 6
'        Else
'            MsgBox "Cannot Process Settlement Discrepancy because Receive Receipt is already Processed", vbInformation, "Information"
'            txtRRNo.SetFocus
'            Exit Sub
'        End If
'    End If
'
'Exit Sub
'er:
'    DoEvents
'    Call modObjTools.ErrorMessage(Err.Description, "Error->txtRRNO()")
'End Sub

Public Sub ShowRecord(mySDiscID As Long, mySTATUS As String, mySTFNO As String, myBCODE As String)
On Error GoTo er
    
    Call clearform
    txtSDiscID.Text = mySDiscID 'Used for Edit Verification
    bCode = myBCODE
    globalStatus = mySTATUS 'use for printing
    Screen.MousePointer = vbHourglass
    
    If mySTATUS = "PENDING" Or mySTATUS = "NEW" Then
        chkFinal.Value = 0
    Else
        chkFinal.Value = 1
    End If
    
    
    DoEvents
    Qry = "call sp_mrcwarehouse.sp_WHSDiscrepancy_getHeader('" & sysBranchCode & "', " & mySDiscID & ",'" & mySTATUS & "','" & mySTFNO & "', '" & myBCODE & "')"
    With mainRs
        .Open Qry, LYRcon, adOpenStatic, adLockOptimistic
        
        If .RecordCount > 0 Then
            .MoveFirst
            
            If !SDTYPE = "LACKING" Then
                cmbBranch.Text = IIf(IsNull(!branchName), "", !branchName)
                txtSTFNO.Text = IIf(IsNull(!STFNO), "", !STFNO)
                txtSTFDate.Text = Format(!STFDATE, "MM/DD/YYYY")
                txtRRno.Text = IIf(IsNull(!RefNo), "", !RefNo)
                txtDateRec.Text = Format(!DATEREF, "MM/DD/YYYY")
                txtStatus.Text = IIf(IsNull(mySTATUS), "", mySTATUS)
                txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
                txtPrepBy.Text = IIf(IsNull(!PREPBY), "", !PREPBY)
            ElseIf !SDTYPE = "DAMAGE" Then
                cmbBranch.Text = IIf(IsNull(!branchName), "", !branchName)
                txtSTFNO.Text = IIf(IsNull(!RRno), "", !RRno)
                txtSTFDate.Text = Format(!STFDATE, "MM/DD/YYYY")
                txtRRno.Text = IIf(IsNull(!STFNO), "", !STFNO)
                txtDateRec.Text = Format(!DATEREC, "MM/DD/YYYY")
                txtStatus.Text = IIf(IsNull(mySTATUS), "", mySTATUS)
                txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
                txtPrepBy.Text = IIf(IsNull(!PREPBY), "", !PREPBY)
            Else
                cmbBranch.Text = IIf(IsNull(!branchName), "", !branchName)
                txtSTFNO.Text = IIf(IsNull(!STFNO), "", !STFNO)
                txtSTFDate.Text = Format(!STFDATE, "MM/DD/YYYY")
                txtRRno.Text = IIf(IsNull(!RRno), "", !RRno)
                txtDateRec.Text = Format(!DATEREC, "MM/DD/YYYY")
                txtRefNo.Text = Format(!RefNo, "MM/DD/YYYY")
                txtDateRef.Text = Format(!DATEREF, "MM/DD/YYYY")
                txtStatus.Text = IIf(IsNull(!Status), "", !Status)
                txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
                txtAuditFindings.Text = IIf(IsNull(!AUDFINDINGS), "", !AUDFINDINGS)
                txtPrepBy.Text = IIf(IsNull(!PREPBY), "", !PREPBY)
                txtAudBy.Text = IIf(IsNull(!AUDBY), "", !AUDBY)
                txtApprvby.Text = IIf(IsNull(!APPRVBY), "", !APPRVBY)
                txtWHSDiscID.Text = IIf(IsNull(!SDISCID), "", !SDISCID)
            End If
        End If
        .Close
    End With
    
    DoEvents
    Qry = "call sp_mrcwarehouse.sp_WHSDiscrepancy_getDetails('" & sysBranchCode & "', " & mySDiscID & ",'" & mySTATUS & "', '" & myBCODE & "')"
    With mainRs
        .Open Qry, LYRcon, adOpenStatic, adLockOptimistic
        
        If .RecordCount > 0 Then
            .MoveFirst
            flxSIMSDiscrepancy.Rows = .RecordCount + 1
            
            For i = 1 To flxSIMSDiscrepancy.Rows - 1
                flxSIMSDiscrepancy.TextMatrix(i, 1) = !PartNumber
                flxSIMSDiscrepancy.TextMatrix(i, 2) = !Description
                flxSIMSDiscrepancy.TextMatrix(i, 3) = !qty
                flxSIMSDiscrepancy.TextMatrix(i, 4) = !Remarks
                flxSIMSDiscrepancy.TextMatrix(i, 5) = IIf(IsNull(!BRECOMMENDATION), "", !BRECOMMENDATION)
                flxSIMSDiscrepancy.TextMatrix(i, 6) = IIf(IsNull(!SRECOMMENDATION), "", !SRECOMMENDATION)
                flxSIMSDiscrepancy.TextMatrix(i, 7) = !SPAREPARTID
                .MoveNext
            Next i
        End If
    End With
    
    If mainRs.State <> 0 Then
        Screen.MousePointer = vbDefault
        mainRs.Close
    End If
    
    'Update ang status sa <databasename>online to Viewed
    'Call UpdateSDiscOnline(mySTATUS, myBCODE, mySDiscID)
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->ShowRecord()")
End Sub

Private Sub txtSTFDate_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtSTFDate, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtRRno.SetFocus
End Sub

Private Sub txtSTFNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtSTFNO, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtSTFDate.SetFocus
End Sub

