VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPOReplStockTrans 
   Caption         =   "P.O STOCK TRANSFER & FILL-IN RATEFORM"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8370
   ScaleWidth      =   13560
   Tag             =   "XYF"
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   120
      TabIndex        =   42
      Tag             =   "XYF"
      Top             =   7560
      Width           =   13335
      Begin VB.ComboBox txtApprvby 
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
         Left            =   9000
         TabIndex        =   46
         Tag             =   "XYF"
         Top             =   240
         Width           =   4215
      End
      Begin VB.TextBox txtPrepby 
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
         Left            =   1800
         TabIndex        =   43
         Tag             =   "XYF"
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label12 
         Caption         =   "APPROVED BY:"
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
         Left            =   7320
         TabIndex        =   45
         Tag             =   "XYF"
         Top             =   280
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "PREPARED BY:"
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
         TabIndex        =   44
         Tag             =   "XYF"
         Top             =   280
         Width           =   1575
      End
   End
   Begin TabDlg.SSTab tabConfirm 
      Height          =   2055
      Left            =   5160
      TabIndex        =   35
      Top             =   2760
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3625
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "CONFIRM INTRANSIT"
      TabPicture(0)   =   "frmPOReplStockTrans.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdQuit"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdConfirm"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtChecker"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtDriver"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.TextBox txtDriver 
         Height          =   375
         Left            =   1800
         TabIndex        =   37
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtChecker 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   36
         Top             =   960
         Width           =   2655
      End
      Begin lvButton.lvButtons_H cmdConfirm 
         Height          =   375
         Left            =   1320
         TabIndex        =   38
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "CONFIRM"
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
         Image           =   "frmPOReplStockTrans.frx":001C
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdQuit 
         Height          =   375
         Left            =   2760
         TabIndex        =   39
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "CANCEL"
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
         Image           =   "frmPOReplStockTrans.frx":05B6
         cBack           =   -2147483633
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "DRIVER NAME:"
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
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "CHECKER NAME:"
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
         Left            =   120
         TabIndex        =   40
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   13560
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   13560
      TabIndex        =   30
      Text            =   "HoldBackColor"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Tag             =   "XYF"
      Top             =   6840
      Width           =   13335
      Begin VB.TextBox txtREMARKS 
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
         Left            =   1440
         TabIndex        =   18
         Tag             =   "XYF"
         Top             =   200
         Width           =   6735
      End
      Begin VB.Label Label5 
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
         Left            =   240
         TabIndex        =   17
         Tag             =   "XYF"
         Top             =   240
         Width           =   1095
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "P&rocess Repl. P.O"
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
      Image           =   "frmPOReplStockTrans.frx":0B50
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Edit Repl. P.O"
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
      Image           =   "frmPOReplStockTrans.frx":10EA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3720
      TabIndex        =   2
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "&View Summary"
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
      Image           =   "frmPOReplStockTrans.frx":1684
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   7560
      TabIndex        =   3
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "E&xit Repl. P.O"
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
      Image           =   "frmPOReplStockTrans.frx":1C1E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5640
      TabIndex        =   4
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
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
      Image           =   "frmPOReplStockTrans.frx":21B8
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   13560
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
      Image           =   "frmPOReplStockTrans.frx":2752
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   13560
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
      Image           =   "frmPOReplStockTrans.frx":2CEC
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Tag             =   "XYF"
      Top             =   360
      Width           =   13335
      Begin VB.CheckBox chkFinal 
         Caption         =   "FINAL"
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
         Left            =   3960
         TabIndex        =   34
         Tag             =   "XYF"
         Top             =   1200
         Width           =   975
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
         Left            =   1800
         TabIndex        =   33
         Tag             =   "XYF"
         Top             =   1140
         Width           =   2055
      End
      Begin VB.ComboBox cboPONO 
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
         Left            =   4800
         TabIndex        =   29
         Tag             =   "XYF"
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox cboBranch 
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
         Left            =   1800
         TabIndex        =   28
         Tag             =   "XYF"
         Top             =   300
         Width           =   4935
      End
      Begin MSMask.MaskEdBox txtDateConfirm 
         Height          =   375
         Left            =   11400
         TabIndex        =   27
         Tag             =   "XYF"
         Top             =   1140
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MSMask.MaskEdBox txtDateProc 
         Height          =   375
         Left            =   11400
         TabIndex        =   14
         Tag             =   "XYF"
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
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
      Begin VB.TextBox txtSTFNO 
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
         Left            =   11400
         TabIndex        =   13
         Tag             =   "XYF"
         Text            =   "<AUTO>"
         Top             =   300
         Width           =   1695
      End
      Begin VB.TextBox txtCATEGORY 
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
         Left            =   1800
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label6 
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
         TabIndex        =   32
         Tag             =   "XYF"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "CONFIRM INTRANSIT:"
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
         Left            =   9120
         TabIndex        =   26
         Tag             =   "XYF"
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "P.O NO.:"
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
         Left            =   3960
         TabIndex        =   25
         Tag             =   "XYF"
         Top             =   795
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE STF. PROCESS:"
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
         Left            =   8760
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   795
         Width           =   2535
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "STOCK TRANSFER #:"
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
         Left            =   9240
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "CATEGORY:"
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
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   795
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "BRANCH NAME:"
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
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4815
      Left            =   120
      TabIndex        =   15
      Tag             =   "XYF"
      Top             =   2040
      Width           =   13335
      Begin VB.TextBox txtOvrAllPercentage 
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
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "0"
         Top             =   4320
         Width           =   1575
      End
      Begin VB.TextBox txtOvrAllUnserved 
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
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "0"
         Top             =   4320
         Width           =   1575
      End
      Begin VB.TextBox txtOvrAllServe 
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
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "0"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox txtOvrAllQTY 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "0"
         Top             =   4320
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid flxPO 
         Height          =   4095
         Left            =   120
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   240
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   7223
         _Version        =   393216
         Rows            =   3
         Cols            =   13
         FixedRows       =   2
         FixedCols       =   0
         FormatString    =   "|^SKU# |<PARTNUMBER |<DESCRIPTION |^ORDER QTY |<PARTNUMBER |^SERVED |^UNSERVED |^PERCENTAGE"
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
      Begin VB.Label lblOvrAllTtl 
         Alignment       =   1  'Right Justify
         Caption         =   "OVERALL TOTAL:"
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
         TabIndex        =   20
         Top             =   4320
         Width           =   1695
      End
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   11640
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   13800
      m_MDI_Height    =   8940
   End
   Begin VB.Menu mnuSTATUS 
      Caption         =   "Popupmenu"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Row"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmPOReplStockTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Qry As String, ans As String
Dim ActRow As Integer, ActCol As Integer
Dim i As Long, j As Long
Dim formAction As String
Dim branchID As String
Dim PmanSrch        As Boolean
Dim PrtSrch         As Boolean
Dim ActSrch         As Boolean
Dim PrtBln          As Boolean
Dim BCat    As String
Dim rsRepl As New ADODB.Recordset
Dim strStatus As String, strDateDeliver As String

Sub CreateTempTableFinalTransfer()
Dim mySTFNO As String, strINTRANSIT As String, SEQNO As Integer
On Error GoTo er
    LYRcon.Execute "Drop Temporary Table IF  Exists sp_mrcwarehouse.TMP_REPLPOFINALTRANSFER_DTLS;"
    
    DoEvents
    LYRcon.Execute " Create Temporary Table IF Not Exists sp_mrcwarehouse.TMP_REPLPOFINALTRANSFER_DTLS( " & _
    "    REMARKS Varchar(20)      " & _
    "   ,SPAREPARTID Int(11)      " & _
    "   ,REFERENCE Varchar(20)      " & _
    "   ,TRANSTYPE Varchar(10)      " & _
    "   ,`OUT` Int(11)      " & _
    "   ,POREF Varchar(30)      " & _
    "   ,UPDATEDBY Int(11)      " & _
    "   ,SEQNO Int(11)      " & _
    "   ,COST   Decimal(10,2)      " & _
    "   ,RPRICE Decimal(10,2)      " & _
    ");"
    
    Qry = "Insert into sp_mrcwarehouse.TMP_REPLPOFINALTRANSFER_DTLS(REMARKS,SPAREPARTID,REFERENCE,TRANSTYPE,`OUT`,POREF,UPDATEDBY,SEQNO,COST,RPRICE) VALUES "
    If txtCATEGORY.Text <> "" Then
        mySTFNO = "STF#" & txtSTFNO.Text
    Else
        mySTFNO = "TR#" & txtSTFNO.Text
    End If
    strINTRANSIT = "INTRANSIT"
    SEQNO = 0
    For i = 2 To flxPO.Rows - 1
        If i > 2 Then Qry = Qry & ","
        SEQNO = SEQNO + 1
        Qry = Qry & "('" & strINTRANSIT & "'," & Val(flxPO.TextMatrix(i, 9)) & ",'" & mySTFNO & "', 'OUT' ," & Val(flxPO.TextMatrix(i, 6)) & ",'" & cboPONO.Text & "'," & UserID & "," & SEQNO & ",'" & flxPO.TextMatrix(i, 10) & "','" & flxPO.TextMatrix(i, 11) & "') "
    Next i
        
    Call LYRcon.Execute(Qry)
       
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CreateTempTableFinalTransfer()")
End Sub

Sub Clearform()
    flxPO.Rows = 2
    flxPO.Rows = 3
    cboBranch.Text = ""
    txtCATEGORY.Text = ""
    cboPONO.Text = ""
    txtSTFNO.Text = "<AUTO>"
    txtDateProc.Text = Format(ServerTime, "mm/dd/yyyy")
    txtStatus.Text = ""
    txtRemarks.Text = ""
    txtOvrAllPercentage.Text = 0
    txtOvrAllQTY.Text = 0
    txtOvrAllServe.Text = 0
    txtOvrAllUnserved.Text = 0
    txtDateConfirm.Text = "__/__/____"
End Sub


Public Sub ShowRecord(MyRef As String, myDateDeliver As String)
    Dim rsTrans As New ADODB.Recordset
    
    Qry = "call sp_mrcwarehouse.sp_STFReplPO_getReplShowDetails('" & MyRef & "')"
    
    strDateDeliver = ""
    
    With rsTrans
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveFirst
            flxPO.Rows = .RecordCount + 2
            cboBranch.Text = !RECBRANCH
            cboPONO.Text = IIf(IsNull(!PONO), "", !PONO)
            txtCATEGORY.Text = IIf(IsNull(!Category), "", !Category)
            txtSTFNO.Text = IIf(IsNull(!TRANSNO), !TRANSFERID, !TRANSNO)
            txtDateProc.Text = Format(!TRANSDATE, "mm/dd/yyyy")
            txtDateConfirm.Text = IIf(IsNull(!DATEFINAL), "__/__/____", Format(!DATEFINAL, "mm/dd/yyyy"))
            strDateDeliver = IIf(IsNull(!DELIVEREDDATE), "", Format(!DELIVEREDDATE, "mm/dd/yyyy"))
            txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
            txtStatus.Text = !Status
            txtPrepBy.Text = !PREPBY
            txtApprvBy.Text = !APPRVBY
            If !Status = "FINAL" Or !Status = "INTRANSIT" Then
                chkFinal.Value = 1
                chkFinal.Enabled = False
            Else
                chkFinal.Value = 0
                chkFinal.Enabled = False
            End If
            
            For i = 2 To flxPO.Rows - 1
                j = i - 1
                flxPO.TextMatrix(i, 1) = j
                flxPO.TextMatrix(i, 2) = !PartNumber
                flxPO.TextMatrix(i, 3) = !Description
                If cboPONO.Text <> "" Then
                    flxPO.TextMatrix(i, 4) = IIf(IsNull(!Order), 0, !Order)
                    flxPO.TextMatrix(i, 7) = IIf(IsNull(!UNSERVED), 0, !UNSERVED)
                End If
                flxPO.TextMatrix(i, 5) = !COUNTERPART
                flxPO.TextMatrix(i, 6) = !SERVED
                
                If !SERVED = 0 Then
                    flxPO.TextMatrix(i, 8) = "0%"
                Else
                    If Val(flxPO.TextMatrix(i, 4)) = 0 Then
                        flxPO.TextMatrix(i, 8) = "100%"
                    Else
                        flxPO.TextMatrix(i, 8) = Round(!SERVED / (Val(flxPO.TextMatrix(i, 4))) * 100) & "%"
                    End If
                End If
                flxPO.TextMatrix(i, 9) = !SPAREPARTID
                flxPO.TextMatrix(i, 10) = !Cost
                flxPO.TextMatrix(i, 11) = !STOREPRICE

                .MoveNext
            Next i
        End If
        .Close
    End With
    
    If txtStatus.Text = "FINAL" Then
        cmdTrans(1).Caption = "Confirm Intransit"
    End If
    Call ComputeTotal
End Sub

Sub CreateTempTable()
On Error GoTo er
        LYRcon.Execute "Drop  Table IF  Exists sp_mrcwarehouse.TMP_REPLSTOCKTRANSFER_DTLS;"

         DoEvents
       LYRcon.Execute " Create  Table IF Not Exists sp_mrcwarehouse.TMP_REPLSTOCKTRANSFER_DTLS( " & _
        "    SPAREPARTID Int(11)      " & _
        "   ,QTY Int(11)      " & _
        "   ,UNIT Varchar(10)      " & _
        "   ,POREF Varchar(30)      " & _
        ");"
        
        Qry = "Insert into sp_mrcwarehouse.TMP_REPLSTOCKTRANSFER_DTLS(SPAREPARTID,QTY,UNIT,POREF) VALUES "
'********************** CODE UPDATED: 06/16/17 **************************
        For i = 2 To flxPO.Rows - 1
            If i > 2 Then Qry = Qry & ","
            Qry = Qry & "(" & Val(flxPO.TextMatrix(i, 9)) & "," & Val(flxPO.TextMatrix(i, 6)) & ",'PCS', '" & cboPONO.Text & "') "
        Next i
'****************** END CODE UPDATED: 06/16/17 **************************
       
       Call LYRcon.Execute(Qry)
       
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CreateTempTable()")
End Sub


Public Sub ComputeTotal()
Dim intTotalORDQTY As Integer
Dim intTotalServed As Integer
Dim intTotalUnserved As Integer

    intTotalORDQTY = 0
    intTotalServed = 0
    intTotalUnserved = 0
    With flxPO
        For i = 2 To .Rows - 1
            intTotalORDQTY = intTotalORDQTY + Val(Replace(.TextMatrix(i, 4), ",", ""))
            intTotalServed = intTotalServed + Val(Replace(.TextMatrix(i, 6), ",", ""))
            intTotalUnserved = intTotalUnserved + Val(Replace(.TextMatrix(i, 7), ",", ""))
        Next i
    End With
    
    txtOvrAllQTY.Text = Format(intTotalORDQTY, "0")
    txtOvrAllServe.Text = Format(intTotalServed, "0")
    txtOvrAllUnserved.Text = Format(intTotalUnserved, "0")
    If Val(txtOvrAllServe.Text) = 0 And Val(txtOvrAllQTY.Text) = 0 Then
        txtOvrAllPercentage.Text = "0%"
    Else
        If Val(txtOvrAllQTY.Text) = 0 Then
            txtOvrAllPercentage.Text = "0%"
        Else
            txtOvrAllPercentage.Text = Round((Val(txtOvrAllServe.Text) / Val(txtOvrAllQTY.Text)) * 100) & "%"
        End If
    End If
End Sub


Public Sub flxColChanger()

With flxPO

    'MERGE COL
    .TextMatrix(0, 0) = " "
    .TextMatrix(1, 0) = " "
    'MERGE ROW
    .TextMatrix(0, 1) = "PURCHASE ORDER"
    .TextMatrix(0, 2) = "PURCHASE ORDER"
    .TextMatrix(0, 3) = "PURCHASE ORDER"
    .TextMatrix(0, 4) = "PURCHASE ORDER"
    .TextMatrix(0, 5) = "PURCHASE ORDER"
    .TextMatrix(0, 6) = "SERVICE RATIO"
    .TextMatrix(0, 7) = "SERVICE RATIO"
    .TextMatrix(0, 8) = "SERVICE RATIO"
    
    'MERGE COL =FALSE
    .TextMatrix(1, 1) = "SKU#"
    .TextMatrix(1, 2) = "PARTNUMBER"
    .TextMatrix(1, 3) = "DESCRIPTION"
    .TextMatrix(1, 4) = "ORDER QTY"
    .TextMatrix(1, 5) = "PARTNUMBER"
    .TextMatrix(1, 6) = "SERVED"
    .TextMatrix(1, 7) = "UNSERVED"
    .TextMatrix(1, 8) = "PERCENTAGE"

    'SET MERGING
    .MergeCells = 1
    .MergeRow(0) = True
    .MergeCol(0) = True
    .MergeCol(1) = True
    For i = 1 To 3
        .MergeCol(i) = False
    Next i
    'SET THE COLUMN COLOR GREEN
    .row = 0
    .col = 0
    .col = 1

    '.Clear
    .Cols = 13
    '.FormatString = "|^MODEL BRAND (PM) |^PARTNUMBER |^DESCRIPTION |^QOH |^STDQ |^SRQ |^QOH |^RO |^AVQ |^ALQ |<LOCATION |^ID"
    .ColWidth(0) = 200
    .ColWidth(1) = (.Width - .ColWidth(0)) * 0.05  'SKU#
    .ColWidth(2) = (.Width - .ColWidth(0)) * 0.15 'PARTNUMBER
    .ColWidth(3) = (.Width - .ColWidth(0)) * 0.23 'DESCRIPTION
    .ColWidth(4) = (.Width - .ColWidth(0)) * 0.1 'ORDER QTY
    .ColWidth(5) = (.Width - .ColWidth(0)) * 0.15 'PARTNUMBER
    .ColWidth(6) = (.Width - .ColWidth(0)) * 0.1 'SERVED
    .ColWidth(7) = (.Width - .ColWidth(0)) * 0.1 'UNSERVED
    .ColWidth(8) = (.Width - .ColWidth(0)) * 0.1 'PERCENTAGE
    .ColWidth(9) = 0 '(.Width - .ColWidth(0)) * 0.06 'SPAREPARTID
    .ColWidth(10) = 0 '(.Width - .ColWidth(0)) * 0.06 'COST
    .ColWidth(11) = 0 '(.Width - .ColWidth(0)) * 0.06 'STOREPRICE
    .ColWidth(12) = 0 '(.Width - .ColWidth(0)) * 0.06 'STOREPRICE
    
    txtOvrAllQTY.Top = .Top + .Height + 50
    txtOvrAllQTY.Left = .Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3)
    txtOvrAllQTY.Width = .ColWidth(4)
    
    txtOvrAllServe.Top = .Top + .Height + 50
    txtOvrAllServe.Left = .Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + .ColWidth(5)
    txtOvrAllServe.Width = .ColWidth(6)
    
    txtOvrAllUnserved.Top = .Top + .Height + 50
    txtOvrAllUnserved.Left = .Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + .ColWidth(5) + .ColWidth(6)
    txtOvrAllUnserved.Width = .ColWidth(7)
    
    txtOvrAllPercentage.Top = .Top + .Height + 50
    txtOvrAllPercentage.Left = .Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + .ColWidth(5) + .ColWidth(6) + .ColWidth(7)
    txtOvrAllPercentage.Width = .ColWidth(8)
    
    lblOvrAllTtl.Top = .Top + .Height + 70
    lblOvrAllTtl.Left = .Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) - 50
    lblOvrAllTtl.Width = .ColWidth(3)
End With

End Sub

Private Sub SetButtonVisibility(DataEntry As Boolean)
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->VIEW SUMM
'3 ->PRINT
'4 ->EXIT
'5 ->SAVE
'6 ->CANCEL


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

Sub LockData(DataEntry As Boolean)
    DataEntry = Not DataEntry

    cboBranch.Locked = DataEntry
    cboPONO.Locked = DataEntry
    txtCATEGORY.Locked = DataEntry
    txtRemarks.Locked = DataEntry
    txtSTFNO.Locked = DataEntry

    If cboBranch.Locked Then
        cboBranch.BackColor = txtHoldGrayBackColor.BackColor
        cboPONO.BackColor = txtHoldGrayBackColor.BackColor
        txtCATEGORY.BackColor = txtHoldGrayBackColor.BackColor
        txtRemarks.BackColor = txtHoldGrayBackColor.BackColor
        txtSTFNO.BackColor = txtHoldGrayBackColor.BackColor
        txtDateProc.BackColor = txtHoldGrayBackColor.BackColor
        txtDateConfirm.BackColor = txtHoldGrayBackColor.BackColor
        txtOvrAllQTY.BackColor = txtHoldGrayBackColor.BackColor
        txtOvrAllServe.BackColor = txtHoldGrayBackColor.BackColor
        txtOvrAllUnserved.BackColor = txtHoldGrayBackColor.BackColor
        txtOvrAllPercentage.BackColor = txtHoldGrayBackColor.BackColor
        txtStatus.BackColor = txtHoldGrayBackColor.BackColor
        txtDateConfirm.Enabled = False
        txtDateProc.Enabled = False
        txtPrepBy.BackColor = txtHoldGrayBackColor.BackColor
        txtApprvBy.BackColor = txtHoldGrayBackColor.BackColor
    Else
        cboBranch.BackColor = txtHoldWhiteBackColor.BackColor
        cboPONO.BackColor = txtHoldWhiteBackColor.BackColor
        txtRemarks.BackColor = txtHoldWhiteBackColor.BackColor
        txtApprvBy.BackColor = txtHoldWhiteBackColor.BackColor
        txtPrepBy.Locked = True
        txtCATEGORY.Locked = True
        txtSTFNO.Locked = True
    End If
End Sub


Private Sub cboBranch_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
    If PmanSrch Then
        PmanSrch = False
        Exit Sub
    End If
    With cboBranch
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            branchID = .ItemData(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                PmanSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            branchID = 0
        End If
    End With
End Sub

Private Sub cboBranch_Click()
    Call cboBranch_Change
End Sub

Private Sub cboBranch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then PmanSrch = True
End Sub

Private Sub cboBranch_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then cboPONO.SetFocus
End Sub

Private Sub cboBranch_LostFocus()
Dim Qry As String
    Qry = "SELECT BRANCHID FROM BRANCH WHERE BRANCHNAME = '" & cboBranch.Text & "' "
    branchID = GetFieldValue(Qry)

    If Val(branchID) = 0 Then Exit Sub
    
    Dim rsPOref As New Recordset
    With rsPOref
        Qry = "SELECT sporderbranch.REFERENCE, sporderbranch.BRANCHSPOID From sporderbranch Where sporderbranch.BRANCHID = '" & Val(branchID) & "' AND sporderbranch.PRINTPICK >= 1 and sporderbranch.`STATUS` = 'PICKING';"
        
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboPONO.AddItem !REFERENCE
                cboPONO.ItemData(cboPONO.ListCount - 1) = !BRANCHSPOID
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    cboPONO.SetFocus
End Sub

Private Sub cboPONO_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
    If PmanSrch Then
        PmanSrch = False
        Exit Sub
    End If
    With cboPONO
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                PmanSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
End Sub

Private Sub cboPONO_Click()

    Qry = "call sp_mrcwarehouse.sp_STFReplPO_getPODetails('" & cboPONO.Text & "')"
    With rsRepl
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenStatic, adLockReadOnly
        
        If .RecordCount > 0 Then
            .MoveFirst
            flxPO.Rows = .RecordCount + 2
            txtCATEGORY.Text = IIf(IsNull(!Category), "", !Category)
            txtStatus.Text = IIf(IsNull(!Status), "", !Status)
            For i = 2 To flxPO.Rows - 1
                j = i - 1
                flxPO.TextMatrix(i, 1) = j
                flxPO.TextMatrix(i, 2) = !PartNumber
                flxPO.TextMatrix(i, 3) = !Description
                flxPO.TextMatrix(i, 4) = !Order
                flxPO.TextMatrix(i, 5) = !COUNTERPART
                flxPO.TextMatrix(i, 6) = !SERVED
                flxPO.TextMatrix(i, 7) = !UNSERVED
                flxPO.TextMatrix(i, 8) = Round((!SERVED / !Order) * 100) & "%"
                flxPO.TextMatrix(i, 9) = !SPAREPARTID
                flxPO.TextMatrix(i, 10) = !Cost
                flxPO.TextMatrix(i, 11) = !STOREPRICE
                .MoveNext
            Next i
        End If
    End With
    Call ComputeTotal
    txtRemarks.SetFocus
End Sub

Private Sub cboPONO_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then PmanSrch = True
End Sub

Private Sub cboPONO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cboPONO_Click
End Sub

Private Sub cmdConfirm_Click()
    Dim MyRef       As String
    Dim rsConfirm   As New ADODB.Recordset
    Dim boQty       As Integer
    
        ans = MsgBox("Do you want to confirm this transaction?", vbInformation + vbYesNo, "Confirm Intransit")
        If ans = vbNo Then Exit Sub
        
        On Error GoTo LocalErr
        
        Screen.MousePointer = vbHourglass
        
        If txtDriver.Text = "" And txtChecker.Text = "" Then
            MsgBox "Please fill the Driver Name and Checker.", vbExclamation, "Confirm Intransit"
        Else
            Call CreateTempTableFinalTransfer
            Qry = "call sp_mrcwarehouse.sp_STFReplPO_FinalTransfer('" & txtDriver.Text & "','" & txtChecker.Text & "','" & cboPONO.Text & "')"
            
            With rsConfirm
                If .State = adStateOpen Then .Close
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                
                If .RecordCount > 0 Then
                    .MoveFirst
                    txtStatus.Text = !Status
                    txtDateConfirm.Text = Format(ServerTime, "mm/dd/yyyy")
                Else
                    MsgBox "UNABLE TO CONFIRM THE STOCKTRANSFER!", vbInformation + vbOKOnly, "UNSUCCESSFULL"
                    Exit Sub
                End If
                .Close
            End With
            
        End If
        MsgBox "Stock Transfer Ref#" & cboPONO.Text & " is INTRANSIT.", vbOKOnly, "Successfully Save"
        tabConfirm.Visible = False
        
        With cmdTrans()
            For i = 0 To .Count - 1
                cmdTrans(i).Enabled = True
            Next i
        End With
        
        Frame1.Enabled = True
        Frame2.Enabled = True
        Frame3.Enabled = True
        Screen.MousePointer = vbDefault
        Exit Sub
LocalErr:
    MsgBox Err.Description & vbCrLf & ". Please Contact MIS", vbCritical, "Confirm Stock Transfer" '
End Sub

Private Sub cmdQuit_Click()
    tabConfirm.Visible = False
    With cmdTrans()
        For i = 0 To .Count - 1
            cmdTrans(i).Enabled = True
        Next i
    End With
        Frame1.Enabled = True
        Frame2.Enabled = True
        Frame3.Enabled = True
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
            txtApprvBy.Clear
            Do While Not .EOF
                txtApprvBy.AddItem !FIRSTNAME & " " & !MIDDLENAME & " " & !LASTNAME
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub cmdTrans_Click(Index As Integer)
Dim strDateFinal As String
On Error GoTo er

    If Index = 0 Then 'add
        Call LockData(True)
        Call SetButtonVisibility(True)
        Call Clearform
        chkFinal.Value = 0
        chkFinal.Enabled = True
        formAction = "add"
        Qry = "SELECT concat(employees.FIRSTNAME,' ',employees.LASTNAME) as PrepBy From employees Where employees.EmpId = " & UserID & ";"
        txtPrepBy.Text = GetFieldValue(Qry)
        Call loadApprovedBy
        txtApprvBy.ListIndex = 0
        cboBranch.SetFocus
    ElseIf Index = 1 Then 'edit
        If flxPO.TextMatrix(2, 1) = "" Then
            MsgBox "No available record", vbOKOnly + vbExclamation, "Confirm Intransit"
            Exit Sub
        End If
        If InStr(txtStatus.Text, "CANCELLED") > 0 Then
            MsgBox "Unable to Confirm " & txtSTFNO.Text & " because it was already Cancelled.", vbOKOnly + vbExclamation, "Confirm Intransit"
            Exit Sub
        End If
        
        If txtStatus.Text = "INTRANSIT" Then
            MsgBox "Stock Transfer " & txtSTFNO.Text & " already Confirm", vbOKOnly + vbExclamation, "Confirm Intransit"
            Exit Sub
        End If
        
        If txtStatus.Text <> "PACKING" Then
            If txtStatus.Text <> "FINAL" Then
                MsgBox "Unable to Confirm Intransit Ref. No. " & txtSTFNO.Text & " because record is not final.", vbOKOnly + vbExclamation, "Edit"
                Exit Sub
            End If
        End If

        If chkFinal.Value = 1 Then
            tabConfirm.Visible = True
            tabConfirm.Caption = "CONFIRM INTRANSIT"
            tabConfirm.Left = (Me.Width - tabConfirm.Width) / 2
            With cmdTrans()
                For i = 0 To .Count - 1
                    cmdTrans(i).Enabled = False
                Next i
            End With
                Frame1.Enabled = False
                Frame2.Enabled = False
                Frame3.Enabled = False
        Else
            formAction = "edit"
            Call SetButtonVisibility(True)
            chkFinal.Enabled = True
            txtRemarks.Locked = False
            txtRemarks.BackColor = txtHoldWhiteBackColor.BackColor
        End If
        
    ElseIf Index = 2 Then 'view summary
        frm_OnProcess2.OwnerForm = "SP PACKING"
        FrmSummary.mySrc = "SP PACKING"
        FrmSummary.Show
        Unload Me
    ElseIf Index = 3 Then 'print
        Qry = "call sp_mrcwarehouse.sp_STFReplPO_printPreview('" & txtSTFNO.Text & "')"
        frmPrn.QryPrn = Qry
        frmPrn.ParamValue1 = SysBranch
        frmPrn.ParamValue2 = Format(SysAddress, "")
        frmPrn.ParamValue3 = txtCATEGORY.Text
        
        frmPrn.ParamValue4 = Replace(cboPONO.Text, GetFieldValue("SELECT branchcode FROM branch WHERE  BRANCHNAME = '" & cboBranch.Text & "';"), "") 'cboPONO.Text
        frmPrn.ParamValue5 = Format(txtDateProc.Text, "mm/dd/yyyy")

        If txtDateConfirm.Text = "__/__/____" Then
            frmPrn.ParamValue6 = ""
        Else
            frmPrn.ParamValue6 = IIf(IsNull(Format(txtDateConfirm.Text, "mm/dd/yyyy")), "", Format(txtDateConfirm.Text, "mm/dd/yyyy"))
        End If
        frmPrn.ParamValue7 = IIf(IsNull(strDateDeliver), "", Format(strDateDeliver, "mm/dd/yyyy"))
        frmPrn.ParamValue8 = IIf(IsNull(txtRemarks.Text), "", txtRemarks.Text)
        frmPrn.ParamValue9 = cboBranch.Text
        frmPrn.ParamValue10 = txtStatus.Text
        frmPrn.ParamValue11 = txtSTFNO.Text
        frmPrn.ParamValue12 = txtPrepBy.Text
        frmPrn.ParamValue13 = txtApprvBy.Text
        
        frmPrn.MySource = "REPL PO STOCKTRANS"
        frmPrn.Show vbModal
    ElseIf Index = 4 Then 'exit
        Unload Me
    ElseIf Index = 5 Then 'save
    If txtStatus.Text = "RESERVED" Then
        MsgBox "REPL REF #:" & cboPONO.Text & " IS NOT YET PICKED!", vbExclamation, "Validation"
        Exit Sub
    End If
    
    ans = MsgBox("Do you want to save?", vbQuestion + vbYesNo, "Save")
    
    If Trim(txtApprvBy.Text) = "" Then
        MsgBox "APPROVED BY IS REQUIRED", vbExclamation, "Validation"
        txtApprvBy.SetFocus
        Exit Sub
    End If
    
    If ans = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
            
    If formAction = "add" Then
        Call CreateTempTable
        DoEvents
        MsgBox ("The status of REF# " + cboPONO.Text + " change to Packing."), vbInformation, "PO Status"
        'check if na click ang FINAL if true procceed creating temp table
        If chkFinal Then
            Call CreateTempTableFinalTransfer
        End If
        Qry = "call sp_mrcwarehouse.sp_STFReplPO_create('" & sysBranchCode & "','" & cboBranch.Text & "','" & cboPONO.Text & "','" & txtRemarks.Text & "'," & UserID & ",'" & txtCATEGORY.Text & "'," & Val(chkFinal.Value) & ",'" & txtApprvBy.Text & "')"
        With rsRepl
            If .State = adStateOpen Then .Close
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            
            If .RecordCount > 0 Then
                MsgBox "Spareparts for Packing Successfully Saved!", vbInformation + vbOKOnly, "SAVE"
                txtSTFNO.Text = !TRANSNO
                txtStatus.Text = !Status
            End If
            .Close
        End With
    ElseIf formAction = "edit" Then
        Call CreateTempTable
        If chkFinal = 1 Then
            Call CreateTempTableFinalTransfer
        End If
        If txtCATEGORY.Text = "WHOLESALE" Then
            BCat = "WHOLESALE"
        Else
            BCat = "WAREHOUSE"
        End If
        
        DoEvents
        Qry = "call sp_mrcwarehouse.sp_STFReplPO_update('" & sysBranchCode & "','" & cboPONO.Text & "','" & txtSTFNO.Text & "','" & cboBranch.Text & "','" & txtRemarks.Text & "'," & UserID & ",'" & BCat & "'," & Val(chkFinal.Value) & ",'" & txtApprvBy.Text & "')"
        With rsRepl
            If .State = adStateOpen Then .Close
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            
            If .RecordCount > 0 Then
                If chkFinal <> 0 Then
                    MsgBox "Spareparts for Finalizing Successfully Saved!", vbInformation + vbOKOnly, "SAVE"
                    txtStatus.Text = !Status
                End If
            End If
        End With
    ElseIf formAction = "confirm" Then
        
    End If
        Call SetButtonVisibility(False)
        Call LockData(False)
        chkFinal.Enabled = False
        formAction = ""
        If txtStatus.Text = "FINAL" Then
            cmdTrans(1).Caption = "Confirm Intransit"
        End If
        Screen.MousePointer = vbDefault
    ElseIf Index = 6 Then 'cancel
        
        ans = MsgBox("Do you want to cancel?", vbQuestion + vbYesNo)
        
        If ans = vbNo Then Exit Sub
        
        Call SetButtonVisibility(False)
        Call LockData(False)
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdTrans-Click()")
End Sub


Private Sub Form_Load()
    Dim rsBranch As New ADODB.Recordset
    Dim Qry As String
    Warehouse.sBar.Panels(6).Text = Me.Caption
    ActRow = 0
    ActCol = 0
    strStatus = ""
    
    Call LockData(False)
    'chkFinal.Enabled = False
    txtSTFNO.Text = "<AUTO>"
    txtDateProc.Text = Format(ServerTime, "mm/dd/yyyy")
    txtDateConfirm.Text = "__/__/____"
    chkFinal.Value = 0
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

    With rsBranch
        Qry = "SELECT BRANCH.BRANCHNAME, BRANCH.BRANCHID From BRANCH ORDER BY BRANCH.BRANCHNAME;"
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboBranch.AddItem !branchName
                cboBranch.ItemData(cboBranch.ListCount - 1) = !branchID
                .MoveNext
            Loop
        End If
        .Close
    End With

End Sub

Private Sub RXResizer1_ResizeComplete()
Call flxColChanger
End Sub

Private Sub txtApprvBy_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub

