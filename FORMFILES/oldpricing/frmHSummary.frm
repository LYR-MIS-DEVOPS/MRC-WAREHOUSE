VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHSummary 
   Caption         =   "Summary"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   15060
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab ssTabTRec 
      Height          =   4335
      Left            =   5640
      TabIndex        =   31
      Top             =   2520
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "TREASURY PORTION"
      TabPicture(0)   =   "frmHSummary.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label9"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label8"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4"
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
      Tab(0).Control(12)=   "txtTRFPNO"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtTCVNo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtTChckNo"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtTBankAcct"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtTRecBy"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      Begin VB.TextBox txtTRecBy 
         Height          =   375
         Left            =   1515
         TabIndex        =   36
         Top             =   520
         Width           =   3375
      End
      Begin VB.TextBox txtTBankAcct 
         Height          =   375
         Left            =   1515
         TabIndex        =   35
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox txtTChckNo 
         Height          =   375
         Left            =   1515
         TabIndex        =   34
         Top             =   1990
         Width           =   3375
      End
      Begin VB.TextBox txtTCVNo 
         Height          =   375
         Left            =   1515
         TabIndex        =   33
         Top             =   2850
         Width           =   3375
      End
      Begin VB.TextBox txtTRFPNO 
         Height          =   495
         Left            =   5400
         TabIndex        =   32
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin lvButton.lvButtons_H btnTOk 
         Height          =   375
         Left            =   720
         TabIndex        =   37
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
         Image           =   "frmHSummary.frx":001C
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H btnTCancel 
         Height          =   375
         Left            =   3000
         TabIndex        =   38
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
         Image           =   "frmHSummary.frx":05B6
         cBack           =   -2147483633
      End
      Begin MSMask.MaskEdBox txtTRecDate 
         Height          =   375
         Left            =   1515
         TabIndex        =   39
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
         TabIndex        =   40
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
         TabIndex        =   41
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
      Begin VB.Label Label4 
         Caption         =   "RECEIVED BY   :"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "RECEIVED ON  :"
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   1035
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "BANK ACCT.      :"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   1650
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "CHECK NO.       :"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   2070
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "CHECK DATE    :"
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   2475
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "CV NO.              :"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "DATE DEPOSIT:"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   3330
         Width           =   1335
      End
   End
   Begin TabDlg.SSTab tabProcessing 
      Height          =   975
      Left            =   5280
      TabIndex        =   27
      Tag             =   "xyf"
      Top             =   3000
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   1720
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmHSummary.frx":0B50
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProcess"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label21"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Label Label21 
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
         Height          =   255
         Left            =   1080
         TabIndex        =   29
         Tag             =   "xyf"
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label lblProcess 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   600
         TabIndex        =   28
         Tag             =   "xyf"
         Top             =   600
         Width           =   3975
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   240
         Picture         =   "frmHSummary.frx":0B6C
         Stretch         =   -1  'True
         Tag             =   "xyf"
         Top             =   120
         Width           =   720
      End
   End
   Begin MSComCtl2.UpDown AscDsc 
      Height          =   375
      Left            =   1800
      TabIndex        =   23
      Tag             =   "xyf"
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComctlLib.ProgressBar Pbar 
      Height          =   135
      Left            =   120
      TabIndex        =   22
      Tag             =   "xy"
      Top             =   9480
      Visible         =   0   'False
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxChk 
      Height          =   4695
      Left            =   16200
      TabIndex        =   20
      Top             =   1200
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   8281
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   13680
      Top             =   0
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   7560
      TabIndex        =   12
      Tag             =   "xyf"
      Top             =   480
      Visible         =   0   'False
      Width           =   7335
      Begin VB.TextBox txtFind 
         Height          =   375
         Left            =   5040
         TabIndex        =   14
         Tag             =   "xyf"
         Top             =   240
         Width           =   2175
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   390
         Left            =   3720
         TabIndex        =   13
         Tag             =   "xyf"
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   688
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
         Image           =   "frmHSummary.frx":1564
         cBack           =   -2147483633
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSum 
      Height          =   7815
      Left            =   120
      TabIndex        =   5
      Tag             =   "xyf"
      Top             =   1320
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   13785
      _Version        =   393216
      RowHeightMin    =   300
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "E&xit Summary"
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
      Image           =   "frmHSummary.frx":1F76
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdShow 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "Sho&w Form"
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
      Image           =   "frmHSummary.frx":2988
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
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
      Image           =   "frmHSummary.frx":2F22
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdFilter 
      Height          =   390
      Left            =   3600
      TabIndex        =   3
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   688
      Caption         =   "&Filter Records"
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
      Image           =   "frmHSummary.frx":3934
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCheck 
      Height          =   390
      Left            =   14640
      TabIndex        =   21
      Tag             =   "xyf"
      Top             =   0
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   688
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
      Image           =   "frmHSummary.frx":4346
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Tag             =   "XYF"
      Top             =   480
      Visible         =   0   'False
      Width           =   7335
      Begin VB.TextBox txtYear 
         Height          =   315
         Left            =   3135
         TabIndex        =   17
         Tag             =   "XYF"
         Top             =   240
         Width           =   720
      End
      Begin VB.ComboBox cboMonth 
         Height          =   315
         Left            =   1440
         TabIndex        =   16
         Tag             =   "XYF"
         Top             =   240
         Width           =   1695
      End
      Begin MSComCtl2.UpDown UpdYear 
         Height          =   315
         Left            =   3840
         TabIndex        =   18
         Tag             =   "XYF"
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Value           =   2015
         OrigLeft        =   11310
         OrigTop         =   480
         OrigRight       =   11565
         OrigBottom      =   795
         Max             =   2020
         Min             =   2005
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "For the month of"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   300
         Width           =   1215
      End
   End
   Begin lvButton.lvButtons_H cmdSaveExcel 
      Height          =   375
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Save Excel"
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
      Image           =   "frmHSummary.frx":4D58
      cBack           =   -2147483633
   End
   Begin MSComDlg.CommonDialog dlgDEMO 
      Left            =   13200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   120
      TabIndex        =   24
      Tag             =   "xyf"
      Top             =   480
      Visible         =   0   'False
      Width           =   7335
      Begin VB.ComboBox cboBranch 
         Height          =   315
         Left            =   1680
         TabIndex        =   25
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "BRANCH NAME"
         Height          =   195
         Left            =   360
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Tag             =   "xyf"
      Top             =   480
      Visible         =   0   'False
      Width           =   7335
      Begin MSComCtl2.DTPicker Dfrom 
         Height          =   300
         Left            =   1440
         TabIndex        =   7
         Tag             =   "xyf"
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         Format          =   142016513
         CurrentDate     =   42107
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   390
         Left            =   6120
         TabIndex        =   8
         Tag             =   "xyf"
         Top             =   195
         Visible         =   0   'False
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   688
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
         Image           =   "frmHSummary.frx":52F2
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker Dto 
         Height          =   300
         Left            =   3840
         TabIndex        =   9
         Tag             =   "xyf"
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         Format          =   142016513
         CurrentDate     =   42107
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   390
         Left            =   5707
         TabIndex        =   10
         Tag             =   "xyf"
         Top             =   195
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   688
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
         Image           =   "frmHSummary.frx":5D04
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdPrev 
         Height          =   390
         Left            =   3360
         TabIndex        =   11
         Tag             =   "xyf"
         Top             =   195
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   688
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
         Image           =   "frmHSummary.frx":6716
         cBack           =   -2147483633
      End
      Begin VB.Label Label1 
         Caption         =   "Date Covered:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Tag             =   "xyf"
         Top             =   285
         Width           =   1215
      End
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   14160
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   15300
      m_MDI_Height    =   9720
   End
   Begin VB.Menu mnuSummary 
      Caption         =   "Summary"
      Visible         =   0   'False
      Begin VB.Menu mnuRFPRec 
         Caption         =   "Receiving"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmHSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSummary       As New ADODB.Recordset
Dim DailyRange      As Boolean
Dim MonthlyRange    As Boolean
Dim FixRange        As Boolean
Dim BranchRange     As Boolean
Dim ActRow          As Integer
Dim ActCol          As Integer
Dim fRow            As Integer
Dim fCol            As Integer
Dim MaxCol          As Integer
Dim ShowDate        As Integer
Public MySource     As String


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
        
            If MySource = "RFPOTHER" Then
                    'PARTNUMBER, DESCRIPTION, REMARKS, INVDATE,
                    'TRANSNO, OUT, PRICE, SUPPLIERNAME, INVTYPE,
                    'CHARGETO, ADDRESS, FRONTLINER, Branch Name, Branch Category
                    wksLineA1.Name = sysBranchCode & " SERVICE RATIO REPORT"
                    wksLineA1.Cells(1, 1).Value = "DATE"
                    wksLineA1.Cells(1, 2).Value = "RFP NO."
                    wksLineA1.Cells(1, 3).Value = "PAYEE"
                    wksLineA1.Cells(1, 4).Value = "ITEMS"
                    wksLineA1.Cells(1, 5).Value = "PREPARED BY"
                    wksLineA1.Cells(1, 6).Value = "CHECKED BY"
                    wksLineA1.Cells(1, 7).Value = "APPROVED BY"
                    wksLineA1.Cells(1, 8).Value = "STATUS"
                    wksLineA1.Cells(1, 9).Value = "STATUTORY"
                    wksLineA1.Cells(1, 10).Value = "TOTAL PAYABLE"
                    wksLineA1.Cells(1, 12).Value = "BANK ACCOUNT"
                    wksLineA1.Cells(1, 13).Value = "CHECK NO"
                                        
                    With flxSum
                        DoEvents
                       
                        tabProcessing.Visible = True
                        cmdExit.Enabled = False
                        DoEvents
                        'Pbar.Value = 0
                        'Pbar.Max = .Rows - 1
                        For i = 1 To .Rows - 1
                            For j = 1 To 13
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
                tabProcessing.Visible = False
                'cmdSaveExcel.Enabled = True
                cmdExit.Enabled = True
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error"
End Sub







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
    
    If flxSum.TextMatrix(ActRow, 8) = "FINAL" Or flxSum.TextMatrix(ActRow, 8) = "FOR APPROVAL" Or flxSum.TextMatrix(ActRow, 8) = "APPROVED WITH CHECK NO." Then
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

Private Sub mnuRFPRec_Click()
Dim TRFPRecRs As New ADODB.Recordset

If MySource = "RFPOTHER" Then
    txtTRFPNO.Text = flxSum.TextMatrix(ActRow, 2)
    If flxSum.TextMatrix(ActRow, 2) <> "" Then
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

Private Sub AscDsc_DownClick()
        flxSum.col = ActCol
        For i = 1 To flxSum.Rows - 1
            flxSum.RowSel = i
            flxSum.Sort = 4
        Next i
        AscDsc.Visible = False
End Sub

Private Sub AscDsc_UpClick()
        flxSum.col = ActCol
        For i = 1 To flxSum.Rows - 1
            flxSum.RowSel = i
            flxSum.Sort = 3
        Next i
        AscDsc.Visible = False
End Sub

Private Sub cboBranch_Change()
    
    Call cmdRefresh_Click
End Sub

Private Sub cboBranch_Click()
    Call cboBranch_Change
End Sub

Private Sub cboMonth_Click()
    If cboMonth.ListIndex = -1 Then Exit Sub
    
    'Identify source whose monthly reports is from Day 1 to Last day of the month
    If mySrc = "CLOSED ACCOUNTS FOR AGENT" Or mySrc = "Monthly Collection" Or mySrc = "Monthly Inventory" Then
        Dfrom.Value = DateSerial(Val(txtYear.Text), cboMonth.ListIndex + 1, 1)
        Dto.Value = DateAdd("M", 1, Dfrom.Value)
        Dto.Value = DateAdd("D", -1, Dto.Value)
    Else    'Use for reports whose date covered is from 7th day to 6th of the following month
        Dfrom.Value = DateSerial(Val(txtYear.Text), cboMonth.ListIndex + 1, 7)
        Dto.Value = DateSerial(Year(DateAdd("M", 1, Dfrom.Value)), Month(DateAdd("M", 1, Dfrom.Value)), 6)
    End If
    Call cmdRefresh_Click
End Sub

Private Sub cmdCheck_Click()
    Dim rsChk As New ADODB.Recordset
    Dim Qry As String, i As Integer, j As Integer
    
    If MySource = "BRANCH BACK ORDER" Then
        Qry = "SELECT backorder.REFERENCE, backorder.BRANCHSPOID, sporderbranch.BRANCHNAME, sporderbranch.branchID, backorder.SPAREPARTID " & _
              "From backorder INNER JOIN sporderbranch ON (backorder.REFERENCE = sporderbranch.REFERENCE) AND (backorder.BRANCHSPOID = sporderbranch.BRANCHSPOID) Order By backorder.TRANSDATE"
        'Display Query in Flx
        If rsChk.State = adStateOpen Then rsChk.Close
        rsChk.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If rsChk.RecordCount = 0 Then   'No records to show
            flxChk.Clear
            flxChk.Rows = 2
            flxChk.Cols = MaxCol
            Call RXResizer1_ResizeComplete
            Exit Sub
        End If
        Set flxChk.DataSource = rsChk   'With records
        'Find Branch in flxChk and display in flxsum
        Pbar.Value = 1
        Pbar.Max = flxSum.Rows
        Pbar.Visible = True
        For i = 1 To flxSum.Rows - 1
            For j = 1 To flxChk.Rows - 1
                If flxSum.TextMatrix(i, 2) = flxChk.TextMatrix(j, 1) And flxSum.TextMatrix(i, 10) = flxChk.TextMatrix(j, 2) And flxSum.TextMatrix(i, 9) = flxChk.TextMatrix(j, 5) Then
                    flxSum.TextMatrix(i, 7) = flxChk.TextMatrix(j, 3)
                    LYRcon.Execute "UPDATE backorder SET BRANCHNAME = '" & flxChk.TextMatrix(j, 3) & "', branchID = " & Val(flxChk.TextMatrix(j, 4)) & " " & _
                        "Where backorder.BRANCHSPOID = " & Val(flxSum.TextMatrix(i, 10)) & " AND backorder.SPAREPARTID = " & Val(flxSum.TextMatrix(i, 9)) & " AND backorder.REFERENCE = '" & flxSum.TextMatrix(i, 2) & "'"
                    Exit For
                End If
            Next j
            Pbar.Value = Pbar.Value + 1
            DoEvents
        Next i
        Pbar.Visible = False
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFilter_Click()
    If MySource = "DETAILED STOCKTRANSFER" Then
        MsgBox "Not yet Available", vbInformation, "FILTER " & MySource
    End If
End Sub

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxSum
        For i = fRow To .Rows - 1
            For j = fCol To .Cols - 1
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .col = 1: .row = i
                    .ColSel = .Cols - 1: .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = .Cols - 1, i + 1, i)
                    fCol = IIf(j = .Cols - 1, 1, j + 1)
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
    cmdNext.Visible = False
    cmdPrev.Visible = True
    Dto.Visible = True
    ShowDate = 2
End Sub

Private Sub cmdPrev_Click()
    Dto.Value = Dfrom.Value
    cmdNext.Left = cmdPrev.Left
    cmdNext.Visible = True
    cmdPrev.Visible = False
    Dto.Visible = False
    ShowDate = 1
End Sub

Private Sub cmdPrint_Click()
    Dim Qry As String
    
    
    If MySource = "BRANCH BACK ORDER" Then
        Qry = "SELECT DATE_FORMAT(backorder.TRANSDATE,'%c/%d/%Y'), backorder.REFERENCE, spareparts.PARTNUMBER, spareparts.DESCRIPTION, backorder.BACKORDERQTY, backorder.BOBALANCE, backorder.BRANCHNAME, backorder.SUPPORDERREF " & _
              "From backorder INNER JOIN spareparts ON (backorder.SPAREPARTID = spareparts.SPAREPARTID) " & _
              "WHERE backorder.TRANSDATE >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' AND backorder.TRANSDATE <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' ORDER BY backorder.TRANSDATE DESC"
    
        frmReport.MySource = "BRANCH BACK ORDER"
        frmReport.QryPrn = Qry
        frmReport.ParamValue1 = "LYR MOTORPARTS CORPORATION"
        frmReport.ParamValue2 = "GANTE AVENUE, TAGUM CITY"
        frmReport.Show vbModal
    
    ElseIf MySource = "BRANCH RECEIVING DISCREPANCY" Then
        
        If cboBranch = "" Then
            MsgBox "Please select branch", vbCritical, "No branch"
        Else
            Qry = "call sProc_mrcwarehouse.sProc_SumRecDis('" & Val(cboBranch.ItemData(cboBranch.ListIndex)) & "')"
            
            frmPrint.MySource = "BRANCH RECEIVING DISCREPANCY PRINT"
            frmPrint.QryPrn = Qry
            frmPrint.ParamValue1 = SysBranch
            frmPrint.ParamValue2 = cboBranch.Text
            frmPrint.ParamValue3 = SysAddress
            frmPrint.Show vbModal
        End If
        
    ElseIf MySource = "DETAILED STOCKTRANSFER" Then
        
        MsgBox "Not yet Available", vbInformation, "PRINT " & MySource
        
    ElseIf MySource = "RFPOTHER" Then
        
     Call ExtractRFP
    End If
    
End Sub

Private Sub cmdSaveExcel_Click()
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

    'Show Customer
    '1:PART NUMBER, 2:DESCRIPTION, 3:CLASSIFICATION 4:DATE SOLD, 5:SALES INV#., 6:QTY., 7:COST W/O VAT, 8:UNIT PRICE, 9:SUPPLIER, 10:TRANS. TYPE,
    '11:CUSTOMER/DEALER, 12:ADDRESS, 13:CITY/TOWN, 14:PROVINCE, 15:REGION, 16:BRANCHES, 17:CUSTOMER TYPE, 18:FRONTLINE 18:CUST.ID
                
       
        If MySource = "DETAILED STOCKTRANSFER" Then
                wksLineA1.Name = "DETAILED STOCKTRANSFER"
                wksLineA1.Cells(1, 1).Value = "DATE"
                wksLineA1.Cells(1, 2).Value = "Ref.No."
                wksLineA1.Cells(1, 3).Value = "PARTNUMBER"
                wksLineA1.Cells(1, 4).Value = "DESCRIPTION"
                wksLineA1.Cells(1, 5).Value = "TRANSFERRED TO"
                wksLineA1.Cells(1, 6).Value = "QUANTITY"
                wksLineA1.Cells(1, 7).Value = "COST W/O VAT"
                wksLineA1.Cells(1, 8).Value = "MDSE INV."
                wksLineA1.Cells(1, 9).Value = "VAT(12%)"
                wksLineA1.Cells(1, 10).Value = "TOTAL"
                wksLineA1.Cells(1, 11).Value = "SUPPLIERS NAME"
                
                With flxSum
                    'Pbar.Value = 0
                    'Pbar.Max = .Rows - 1
                    For i = 1 To .Rows - 1
                        For j = 1 To 11
                            If j = 7 And Index = 2 Then
                                wksLineA1.Cells(i + 1, j).Value = ""
                            Else
                                wksLineA1.Cells(i + 1, j).Value = .TextMatrix(i, j)
                            End If
                        Next j
                        'Pbar.Value = Pbar.Value + 1
                        DoEvents
                    Next i
                End With
                XLapp.DisplayAlerts = False
    
            ElseIf MySource = "MONTHLY MERCHANDISE INVENTORY" Then
                If InStr(GetFieldValue("SELECT pcuser.DESIGNATION From pcuser Where pcuser.EmpId ='" & UserID & "'"), "AUDIT") Then
                    
                    wksLineA1.Name = "MONTHLY MERCHANDISE INVENTORY"
                    wksLineA1.Cells(1, 1).Value = "PARTNUMBER"
                    wksLineA1.Cells(1, 2).Value = "DESCRIPTION"
                    wksLineA1.Cells(1, 3).Value = "QTY OF THE MONTH"
                    wksLineA1.Cells(1, 4).Value = "STORE PRICE"
                    wksLineA1.Cells(1, 5).Value = "DEALER PRICE"
                    wksLineA1.Cells(1, 6).Value = "COST W/O VAT"
                    wksLineA1.Cells(1, 7).Value = "SUPPLIER NAME"
                    wksLineA1.Cells(1, 8).Value = "MONTH OF"
                    
                    With flxSum
                        tabProcessing.Visible = True
                        cmdSaveExcel.Enabled = False
                        cmdExit.Enabled = False
                        'Pbar.Value = 0
                        'Pbar.Max = .Rows - 1
                        For i = 1 To .Rows - 1
                            For j = 1 To 8
                                If j = 7 And Index = 2 Then
                                    wksLineA1.Cells(i + 1, j).Value = ""
                                Else
                                    wksLineA1.Cells(i + 1, j).Value = .TextMatrix(i, j)
                                End If
                            Next j
                            'Pbar.Value = Pbar.Value + 1
                            DoEvents
                        Next i
                    End With
                    XLapp.DisplayAlerts = False
                Else
                    wksLineA1.Name = "MONTHLY MERCHANDISE INVENTORY"
                    wksLineA1.Cells(1, 1).Value = "PARTNUMBER"
                    wksLineA1.Cells(1, 2).Value = "DESCRIPTION"
                    wksLineA1.Cells(1, 3).Value = "QTY OF THE MONTH"
                    wksLineA1.Cells(1, 4).Value = "STORE PRICE"
                    wksLineA1.Cells(1, 5).Value = "DEALER PRICE"
                    wksLineA1.Cells(1, 6).Value = "COST W/O VAT"
                    wksLineA1.Cells(1, 7).Value = "SUPPLIER NAME"
                    wksLineA1.Cells(1, 8).Value = "MONTH OF"
                    
                    With flxSum
                        tabProcessing.Visible = True
                        cmdSaveExcel.Enabled = False
                        cmdExit.Enabled = False
                        'Pbar.Value = 0
                        'Pbar.Max = .Rows - 1
                        For i = 1 To .Rows - 1
                            For j = 1 To 8
                                If j = 6 Then
                                    wksLineA1.Cells(i + 1, j).Value = ""
                                Else
                                    wksLineA1.Cells(i + 1, j).Value = .TextMatrix(i, j)
                                End If
                            Next j
                            'Pbar.Value = Pbar.Value + 1
                            DoEvents
                        Next i
                    End With
                    XLapp.DisplayAlerts = False
                End If
                
            ElseIf MySource = "EXTRACTSALES" Then
                    'PARTNUMBER, DESCRIPTION, REMARKS, INVDATE,
                    'TRANSNO, OUT, PRICE, SUPPLIERNAME, INVTYPE,
                    'CHARGETO, ADDRESS, FRONTLINER, Branch Name, Branch Category
                    wksLineA1.Name = sysBranchCode & " MONTHLY SALES REPORT"
                    wksLineA1.Cells(1, 1).Value = "PARTNUMBER"
                    wksLineA1.Cells(1, 2).Value = "DESCRIPTION"
                    wksLineA1.Cells(1, 3).Value = "CLASSIFICATION"
                    wksLineA1.Cells(1, 4).Value = "INVOICE DATE"
                    wksLineA1.Cells(1, 5).Value = "INVOICE NO."
                    wksLineA1.Cells(1, 6).Value = "OUT"
                    wksLineA1.Cells(1, 7).Value = "PRICE"
                    wksLineA1.Cells(1, 8).Value = "SUPPLIER NAME"
                    wksLineA1.Cells(1, 9).Value = "INVOICE TYPE"
                    wksLineA1.Cells(1, 10).Value = "CHARGE TO/CUSTOMER"
                    wksLineA1.Cells(1, 11).Value = "ADDRESS"
                    wksLineA1.Cells(1, 12).Value = "FRONTLINER"
                    wksLineA1.Cells(1, 13).Value = "CUSTOMER TYPE"
                    wksLineA1.Cells(1, 14).Value = "BRANCH NAME"
                    wksLineA1.Cells(1, 15).Value = "BRANCH CATEGORY"
                                        
                    With flxSum
                        DoEvents
                        tabProcessing.Visible = True
                        cmdSaveExcel.Enabled = False
                        cmdExit.Enabled = False
                        DoEvents
                        'Pbar.Value = 0
                        'Pbar.Max = .Rows - 1
                        For i = 1 To .Rows - 1
                            For j = 1 To 15
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
            ElseIf MySource = "DETAILED SUPPLIER SERVICE RATIO" Then

                    wksLineA1.Name = "SUPPLIER SERVICE RATIO REPORT"
                    wksLineA1.Cells(1, 1).Value = "NO."
                    wksLineA1.Cells(1, 2).Value = "P.O DATE"
                    wksLineA1.Cells(1, 3).Value = "P.O NUMBER"
                    wksLineA1.Cells(1, 4).Value = "PARTNUMBER"
                    wksLineA1.Cells(1, 5).Value = "DESCRIPTION"
                    wksLineA1.Cells(1, 6).Value = "ORDERED QTY"
                    wksLineA1.Cells(1, 7).Value = "SERVED QTY"
                    wksLineA1.Cells(1, 8).Value = "COST"
                    wksLineA1.Cells(1, 9).Value = "CLASSIFICATION"
                    wksLineA1.Cells(1, 10).Value = "SUPPLIER"
                    
                    With flxSum
                        DoEvents
                        tabProcessing.Visible = True
                        cmdSaveExcel.Enabled = False
                        cmdExit.Enabled = False
                        DoEvents
                        For i = 1 To .Rows - 1
                            For j = 1 To 10
                                    wksLineA1.Cells(i + 1, j).Value = .TextMatrix(i, j)
                            Next j
                            DoEvents
                        Next i
                    End With
                    XLapp.DisplayAlerts = False
    Else
        MsgBox "Not yet Available", vbInformation, "SAVE EXCEL " & MySource
        Exit Sub
    End If
    
    wbkMVR.SaveAs filename:=sSaveAsFileName, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
    wbkMVR.Close SaveChanges:=False
    XLapp.Quit
 
    Set XLapp = Nothing
    Set wbkMVR = Nothing
    Set wksLineA1 = Nothing
    MsgBox "Finish Saving to Excel!", vbInformation + vbOKOnly, "SAVE EXCEL " & MySource
    tabProcessing.Visible = False
    cmdSaveExcel.Enabled = True
    cmdExit.Enabled = True
                
End Sub

Private Sub cmdRefresh_Click()
Dim Qry As String
'On Error GoTo ConError

    DoEvents
    tabProcessing.Caption = MySource
    tabProcessing.Visible = True
    DoEvents
    If MySource = "PRODUCTS" Then
        Qry = "SELECT products.BRAND, products.MODEL, products.DESCRIPTION, products.CLASSIFICATION, products.PRODTYPE, products.PRODINIT, products.ACCTITLE, products.PRODID " & _
            "From products Order By products.BRAND, products.MODEL"
    
    ElseIf MySource = "BRANCH BACK ORDER" Then
        Qry = "SELECT DATE_FORMAT(backorder.TRANSDATE,'%c/%d/%Y'), backorder.REFERENCE, spareparts.PARTNUMBER, spareparts.DESCRIPTION, backorder.BACKORDERQTY, backorder.BOBALANCE, backorder.BRANCHNAME, backorder.SUPPORDERREF, backorder.SPAREPARTID, backorder.BRANCHSPOID, backorder.BRANCHID, backorder.SUPPOSTATUS " & _
              "From backorder INNER JOIN spareparts ON (backorder.SPAREPARTID = spareparts.SPAREPARTID) "
        
        Qry = "SELECT spareparts.PARTNUMBER, spareparts.DESCRIPTION, DATE_FORMAT(spareparts.LASTUPDATE,'%m/%d/%Y') as DATE_FORMATTED, Format(spareparts.STOREPRICE,2), spareparts.REMARKS FROM spareparts WHERE spareparts.LASTUPDATE >= '" & Format(Dfrom.Value, "yyyy-mm-dd h:mm:ss") & "' And spareparts.LASTUPDATE <= '" & Format(Dto.Value, "yyyy-mm-dd 23:59:59") & "' ORDER BY spareparts.PARTNUMBER"
'        MsgBox Qry & _
              "WHERE backorder.TRANSDATE >= '" & Format(dFrom.Value, "YYYY-MM-DD") & "' AND backorder.TRANSDATE <= '" & Format(dTo.Value, "YYYY-MM-DD") & "' ORDER BY backorder.TRANSDATE DESC"
    
    ElseIf MySource = "CLOSED ACCOUNTS FOR AGENT" Then
        Qry = "SELECT CONCAT(agents.LASTNAME, ', ', agents.FIRSTNAME) AS AGENT, agentorders.RANK, DATE_FORMAT(unitordered.DATECLOSED,'%c/%d/%Y'), unitordered.ACCOUNTNO, CONCAT(customer.LASTNAME, ', ', customer.FIRSTNAME) AS customername, CONCAT(customer.ADDRESS, ' ', customer.TOWN) AS custAddress, concat(productline.BRAND, ' ', productline.DESCRIPTION, ' ', unitordered.MODEL) AS UnitDesc " & _
              "From agingprev INNER JOIN unitordered ON (agingprev.ORDERID = unitordered.ORDERID) INNER JOIN agentorders ON (agingprev.ORDERID = agentorders.ORDERID) INNER JOIN agents ON (agentorders.AGENTID = agents.AGENTID) INNER JOIN customer ON (unitordered.CUSTOMERID = customer.CUSTOMERID) INNER JOIN productline ON (unitordered.PRODID = productline.PRODID) " & _
              "Where (agentorders.RANK = 'USM' OR agentorders.RANK = 'USS') AND unitordered.DATECLOSED >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' AND unitordered.DATECLOSED < '" & Format(Dto.Value, "YYYY-MM-DD") & "' AND agingprev.REMARKS = 'CLOSED' AND agingprev.`STATUS` = 'GOOD' Order By agentorders.RANK, AGENT"
    
    ElseIf MySource = UCase("Branch Receiving Discrepancy") Then
        'On Error GoTo RecDisError
        Qry = "call sProc_mrcwarehouse.sProc_SumRecDis('" & Val(cboBranch.ItemData(cboBranch.ListIndex)) & "')"
    
    ElseIf MySource = "PRICING" Then
        
         Qry = "SELECT spareparts.PARTNUMBER,  spareparts.DESCRIPTION,  Format(spareparts.STOREPRICE,2),  DATE_FORMAT(spareparts.PRICEUPDATE,'%m/%d/%Y'), FORMAT(SUBSTR(spareparts.PRICEHIST, 12, 10),2) As PREVCOST," & _
        "DATE_FORMAT(SUBSTR(spareparts.PRICEHIST, 1, 10), '%m/%d/%Y') AS PREVUPDATE " & " From spareparts WHERE spareparts.PRICEUPDATE >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And spareparts.PRICEUPDATE <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' ORDER BY spareparts.PARTNUMBER"
    
    ElseIf MySource = "PARTNUMBER" Then
'        format(SUBSTR(spareparts.LASTUPDATE, 1, 10),'%m,%d%y')
    ElseIf MySource = "DETAILED STOCKTRANSFER" Then
        Qry = "CALL sp_mrcwarehouse.sp_RptDtlStockTrans('" & SysBranchID & "','" & Format(Dfrom.Value, "yyyy-mm-dd") & "' ,'" & Format(Dto.Value, "yyyy-mm-dd") & "');"
    
    ElseIf MySource = "MONTHLY MERCHANDISE INVENTORY" Then
'        Qry = "SELECT spareparts.PARTNUMBER, spareparts.DESCRIPTION, mnthlyinv.QTYONHAND, " & _
'                "format(mnthlyinv.STOREPRICE, '0000.00') as Price, mnthlyinv.DLRPRICE, date_format(mnthlyinv.MONTHEND, '%M, %Y'), mnthlyinv.COST, mnthlyinv.SPAREPARTID " & _
'                "From spareparts INNER JOIN mnthlyinv ON (spareparts.SPAREPARTID = mnthlyinv.SPAREPARTID) " & _
'                "Where  monthname(mnthlyinv.MONTHEND) = '" & cboMonth.Text & "' and year(mnthlyinv.MONTHEND) = " & Val(txtYear.Text) & ";"
        Qry = "Select partnumber, DESCRIPTION, mInv.QTYONHAND, format(mInv.STOREPRICE, 2) as Price, " & _
                    "FORMAT(mInv.DLRPRICE, 2) AS DLPRICE, format(mInv.COST,2) as spcost, suppliername, date_format(mInv.MONTHEND, '%M, %Y'), mInv.SPAREPARTID from mnthlyinv as mInv " & _
                        "Left Join " & _
                            "(Select sp.sparepartid, sp.partnumber, sp.description, sp.SUPPLIERCODE from spareparts as sp) as spTab on mInv.SPAREPARTID = sptab.SPAREPARTID " & _
                        "Left Join " & _
                            "(select suppliercode, suppliername from supplier) as sup on sptab.suppliercode = sup.suppliercode " & _
                    "Where  monthname(mInv.MONTHEND) = '" & cboMonth.Text & "' and year(mInv.MONTHEND) = " & Val(txtYear.Text) & " GROUP BY SPAREPARTID;"
    ElseIf MySource = "DETAILED SUPPLIER SERVICE RATIO" Then
        Qry = "CALL sp_mrcwarehouse.sp_Dtld_Supplier_Service_Ratio('" & Format(Dfrom.Value, "yyyy-mm-dd") & "' ,'" & Format(Dto.Value, "yyyy-mm-dd") & "');"
    ElseIf MySource = "RFPOTHER" Then
'        If cboPrepBy.Text = "" Then
            If dCode = 1 Then
                Qry = "SELECT a.RFPDATE,a.RFPNO, a.PAYEE,COUNT(b.RFPID) as ITEMS," & _
                      "emp.PREPBY,a.CHECKEDBY,a.APPROVEDBY,a.RFPSTATUS,a.RFPCLASS, FORMAT(b.NETPAY,2), a.PrintCopy, treasury.BANKACCNT, a.CHECKNO, a.RFPID " & _
                      "FROM mrcwarehouse.rfp_main as a INNER JOIN mrcwarehouse.rfp_dtls as b ON a.RFPID = b.RFPID " & _
                      "LEFT JOIN mrcwarehouse.rfp_treasury as treasury ON treasury.RFPNO = a.RFPNO " & _
                      "INNER JOIN (Select CONCAT(FIRSTNAME,' ',MIDDLENAME,' ',LASTNAME) as PREPBY, EMPID From mrcwarehouse.Employees) as emp ON a.UPDATEDBY = emp.EMPID " & _
                      "WHERE a.RFPDATE = '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' " & _
                      "AND a.RFPCLASS = 'OTHER' GROUP BY a.RFPNO ORDER BY RFPDATE DESC;"
            Else
                Qry = "SELECT a.RFPDATE,a.RFPNO, a.PAYEE,COUNT(b.RFPID) as ITEMS," & _
                      "emp.PREPBY,a.CHECKEDBY,a.APPROVEDBY,a.RFPSTATUS,a.RFPCLASS, FORMAT(b.NETPAY,2), a.PrintCopy, treasury.BANKACCNT, a.CHECKNO, a.RFPID " & _
                      "FROM mrcwarehouse.rfp_main as a INNER JOIN mrcwarehouse.rfp_dtls as b ON a.RFPID = b.RFPID " & _
                      "LEFT JOIN mrcwarehouse.rfp_treasury as treasury ON treasury.RFPNO = a.RFPNO " & _
                      "INNER JOIN (Select CONCAT(FIRSTNAME,' ',MIDDLENAME,' ',LASTNAME) as PREPBY, EMPID From mrcwarehouse.Employees) as emp ON a.UPDATEDBY = emp.EMPID " & _
                      "WHERE a.RFPDATE >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' " & _
                      "AND a.RFPDATE <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' AND a.RFPCLASS = 'OTHER' GROUP BY a.RFPNO ORDER BY RFPDATE DESC;"
            End If
    End If
    
    
    'Display Query in Flx
    If rsSummary.State = adStateOpen Then rsSummary.Close
    rsSummary.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    
    If rsSummary.RecordCount = 0 Then   'No records to show
        flxSum.Clear
        flxSum.Rows = 2
        flxSum.Cols = MaxCol
        DoEvents
        tabProcessing.Visible = False
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxSum.DataSource = rsSummary   'With records
    DoEvents
    tabProcessing.Visible = False
    ActRow = 1
    Call RXResizer1_ResizeComplete
    
Exit Sub
'ConError:
 '   MsgBox Err.Description & Err.Number
End Sub

Private Sub cmdShow_Click()
    If MySource = "PRODUCTS" Then
        Call frmproductEnt.ShowRecord(Val(flxSum.TextMatrix(ActRow, 8)))
        frmproductEnt.Show vbModal
    ElseIf MySource = "RFPOTHER" Then
    
        frmRFPOther.RFPID = flxSum.TextMatrix(ActRow, 14)
        frmRFPOther.tmrGetRecord.Enabled = True
        
        DoEvents
        frm_OnProcess2.Show vbModal
        Unload Me
        
    End If
End Sub

Private Sub Dfrom_Change()
    If ShowDate = 1 Then Dto.Value = Dfrom.Value
End Sub

Private Sub flxSum_Click()
    ActRow = flxSum.row
    ActCol = flxSum.col
    If ActCol <> 3 Then AscDsc.Visible = False
        If MySource = "BRANCH BACK ORDER" Then
            If ActRow = 1 Then
                Select Case ActCol
                    Case 3
                        
                        With AscDsc
                            .Visible = True
                             'fScrol.Left = flxSP.ColPos(6) + 80
                            .Height = flxSum.CellHeight
                            .Top = flxSum.Top
                            .Left = flxSum.ColWidth(4) + 2500
                        End With
                End Select
            End If
        End If
    
End Sub



Private Sub flxSum_EnterCell()
    ActRow = flxSum.row
End Sub

Private Sub flxSum_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        If MySource = "BRANCH BACK ORDER" Then
            mnuCancel.Caption = "CANCEL " & MySource
            Call PopupMenu(mnuSummary)
        ElseIf MySource = "RFPOTHER" Then
            If flxSum.TextMatrix(ActRow, 2) <> "" Then
                Call PopupMenu(mnuSummary)
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    
    Me.Caption = MySource & " SUMMARY"
    Warehouse.sBar.Panels(6).Text = Me.Caption
    If MySource = "MONTHLY MERCHANDISE INVENTORY" Or MySource = "EXTRACTSALES" Then
    
        cmdShow.Visible = False
        cmdPrint.Visible = False
        cmdFilter.Visible = False
        cmdSaveExcel.Visible = True
        cmdSaveExcel.Left = 0
        cmdExit.Left = cmdSaveExcel.Width
    End If
    tabProcessing.Caption = MySource

    Timer1.Enabled = True
    
    If MySource = "RFPOTHER" Then
        cmdPrint.Caption = "EXTRACT"
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     Warehouse.sBar.Panels(6).Text = ""
End Sub

Private Sub mnuCancel_Click()

    If MySource = "BRANCH BACK ORDER" Then
        With flxSum
            If .TextMatrix(ActRow, 2) = "" Then
                MsgBox "No record found", vbExclamation, "CANCEL"
                Exit Sub
            ElseIf .TextMatrix(ActRow, 12) = "CANCELLED" Or .TextMatrix(ActRow, 12) = "FINAL" Then
                MsgBox "This BBO is already " & .TextMatrix(ActRow, 12), vbCritical, "CANCEL"
                Exit Sub
            Else
                ans = MsgBox("Are you sure to cancel BBO Ref#: " & .TextMatrix(ActRow, 2) & " (SP#:" & .TextMatrix(ActRow, 3) & ")", vbYesNo, "CANCEL BRANCH BACK ORDER")
                If ans = vbNo Then
                    Exit Sub
                End If
                    
                            'Qry = "call sProc_CancelBBO('" & Val(Replace(.TextMatrix(ActRow, 9), ",", "")) & "','" & Val(Replace(.TextMatrix(ActRow, 11), ",", "")) & "','" & .TextMatrix(ActRow, 2) & "')"
                            LYRcon.Execute "UPDATE backorder SET SUPPOSTATUS = 'CANCELLED', BOBALANCE = 0 WHERE backorder.SPAREPARTID = '" & Val(Replace(.TextMatrix(ActRow, 9), ",", "")) & "' AND backorder.REFERENCE = '" & .TextMatrix(ActRow, 2) & "' AND backorder.BRANCHID = '" & Val(Replace(.TextMatrix(ActRow, 11), ",", "")) & "';"
                            'LYRcon.Execute Qry
                            Exit Sub
            End If
        End With
    End If
    
End Sub

Private Sub RXResizer1_ResizeComplete()
    With flxSum
        If MySource = "PRODUCTS" Then
            .FormatString = "   |^BRAND|<MODEL|<DESCRIPTION|<CLASSIFICATION|<PRODUCT TYPE|<PROD. INITIAL|<ACCOUNT TITLE"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.12
            .ColWidth(2) = (.Width - 300) * 0.12
            .ColWidth(3) = (.Width - 300) * 0.2
            .ColWidth(4) = (.Width - 300) * 0.15
            .ColWidth(5) = (.Width - 300) * 0.12
            .ColWidth(6) = (.Width - 300) * 0.08
            .ColWidth(7) = (.Width - 300) * 0.18
            .ColWidth(8) = 0
        ElseIf MySource = "BRANCH BACK ORDER" Then
           .FormatString = " |^DATE|^REF.NO.|<PARTNUMBER|<DESCRIPTION|^BBO QTY.|^BBO BAL.|<BRANCH |<P.O. REF#. |^|^|^ |<P.O. STATUS"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.07
            .ColWidth(2) = (.Width - 300) * 0.07
            .ColWidth(3) = (.Width - 300) * 0.18
            .ColWidth(4) = (.Width - 300) * 0.2
            .ColWidth(5) = (.Width - 300) * 0.05
            .ColWidth(6) = (.Width - 300) * 0.05
            .ColWidth(7) = (.Width - 300) * 0.1
            .ColWidth(8) = (.Width - 300) * 0.1
            .ColWidth(9) = 0                        'SPID
            .ColWidth(10) = 0                       'BSPOID
            .ColWidth(11) = 0                       'BID
            .ColWidth(12) = (.Width - 300) * 0.1
            .row = 1
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) = "" Then
                    Exit Sub
                End If
                If .TextMatrix(i, 12) = "CANCELLED" And Val(.TextMatrix(i, 6)) = 0 Then
                    .col = 2: .row = i: .ColSel = 12: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
                    .col = 2
                End If
            Next i
        ElseIf MySource = "CLOSED ACCOUNTS FOR AGENT" Then
            .FormatString = "  |<NAME OF AGENT|^RANK|^DATE CLOSED|^ACCOUNTNO|<CUSTOMER NAME|<ADDRESS|<UNIT DESCRIPTION"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.15
            .ColWidth(2) = (.Width - 300) * 0.08
            .ColWidth(3) = (.Width - 300) * 0.08
            .ColWidth(4) = (.Width - 300) * 0.1
            .ColWidth(5) = (.Width - 300) * 0.18
            .ColWidth(6) = (.Width - 300) * 0.18
            .ColWidth(7) = (.Width - 300) * 0.18
        'PARTNUMBER, DESCRIPTION, PACKINGNO, DATEREC, QTY, GOOD, DAMAGE, LACKING, EXCESS, WRONGPN, REMARKS, SPAREPARTID
        ElseIf MySource = "BRANCH RECEIVING DISCREPANCY" Then
            .FormatString = "  |<PARTNUMBER|<DESCRIPTION|^PACKING NO.|^DATE RECEIVE|^TR. QTY|^REC. QTY |^DAMAGE|^LACKING|^EXCESS|^WRONG SP|<REMARKS"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.13
            .ColWidth(2) = (.Width - 300) * 0.3
            .ColWidth(3) = (.Width - 300) * 0.06
            .ColWidth(4) = (.Width - 300) * 0.07
            .ColWidth(5) = (.Width - 300) * 0.05
            .ColWidth(6) = (.Width - 300) * 0.05
            .ColWidth(7) = (.Width - 300) * 0.05
            .ColWidth(8) = (.Width - 300) * 0.05
            .ColWidth(9) = (.Width - 300) * 0.05
            .ColWidth(10) = (.Width - 300) * 0.06
            .ColWidth(11) = (.Width - 300) * 0.1
            .ColWidth(12) = 0
        
        ElseIf MySource = "PRICING" Then
            .FormatString = "  |<PARTNUMBER|^DESCRIPTION|^STOREPRICE|^DATE UPDATED|<PREV-PRICE|<PREV-UPDATE"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.18
            .ColWidth(2) = (.Width - 300) * 0.27
            .ColWidth(3) = (.Width - 300) * 0.13
            .ColWidth(4) = (.Width - 300) * 0.15
            .ColWidth(5) = (.Width - 300) * 0.13
            .ColWidth(6) = (.Width - 300) * 0.13
            .ColWidth(7) = 0
            .ColWidth(8) = 0
            
        ElseIf MySource = "PARTNUMBER" Then
            .FormatString = "  |<PARTNUMBER|^DESCRIPTION|^DATE CREATED|^STOREPRICE|<REMARKS"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.18
            .ColWidth(2) = (.Width - 300) * 0.32
            .ColWidth(3) = (.Width - 300) * 0.13
            .ColWidth(4) = (.Width - 300) * 0.15
            .ColWidth(5) = (.Width - 300) * 0.2
        ElseIf MySource = "DETAILED STOCKTRANSFER" Then
            .FormatString = "  |<DATE|^REF.NO.|<PARTNUMBER|<DESCRIPTION|<TRANSFERRED TO|^QTY|>COST W/O VAT|>MSDE|>VAT|>TOTAL|<SUPPLIERS NAME"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.05
            .ColWidth(2) = (.Width - 300) * 0.05
            .ColWidth(3) = (.Width - 300) * 0.13
            .ColWidth(4) = (.Width - 300) * 0.25
            .ColWidth(5) = (.Width - 300) * 0.13
            .ColWidth(6) = (.Width - 300) * 0.05
            .ColWidth(7) = (.Width - 300) * 0.08
            .ColWidth(8) = (.Width - 300) * 0.05
            .ColWidth(9) = (.Width - 300) * 0.05
            .ColWidth(10) = (.Width - 300) * 0.05
            .ColWidth(11) = (.Width - 300) * 0.2
        ElseIf MySource = "MONTHLY MERCHANDISE INVENTORY" Then
            .FormatString = " |<PARTNUMBER|<DESCRIPTION|^QTY ON HAND|>RETAIL PRICE|^DEALER PRICE|^COST W/O VAT|<SUPPLIER NAME|^CUTOFF DATE"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.2     'Spareparts
            .ColWidth(2) = (.Width - 300) * 0.3     'description
            .ColWidth(3) = (.Width - 300) * 0.08    'qty
            .ColWidth(4) = (.Width - 300) * 0.08    'price
            .ColWidth(5) = 0                        'dealer price
            .ColWidth(6) = 0                        'cost
            .ColWidth(7) = (.Width - 300) * 0.08    'suppliername
            .ColWidth(8) = (.Width - 300) * 0.08    'cutoff date
            '.ColWidth(8) = 0                        'cost
            .ColWidth(9) = 0                        'spid
        ElseIf MySource = "DETAILED SUPPLIER SERVICE RATIO" Then
            .FormatString = "  |^NO |<P.O DATE |^P.O NUMBER |<PARTNUMBER |<DESCRIPTION |^ORDERED QTY |^SERVED QTY |>COST |<CLASSIFICATION |<SUPPLIERS NAME"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.03        'NO
            .ColWidth(2) = (.Width - 300) * 0.06        'DATE
            .ColWidth(3) = (.Width - 300) * 0.11        'PO#
            .ColWidth(4) = (.Width - 300) * 0.15        'PART#
            .ColWidth(5) = (.Width - 300) * 0.25        'DESC
            .ColWidth(6) = (.Width - 300) * 0.07        'ORD QTY
            .ColWidth(7) = (.Width - 300) * 0.07        'SERV QTY
            .ColWidth(8) = (.Width - 300) * 0.07        'COST
            .ColWidth(9) = (.Width - 300) * 0.13        'CLASS
            .ColWidth(10) = (.Width - 300) * 0.2        'SUPPLIER
        ElseIf MySource = "RFPOTHER" Then

          '                  0    1      2        3          4            5           6            7           8        9              10           11
            .FormatString = "  |^DATE|^RFP NO.|<PAYEE  |^ITEMS |<PREPARED BY |<CHECKED BY |<APPROVED BY |^STATUS |^STATUTORY |>TOTAL PAYABLE |^PRINT COPY  |<BANK ACCNT |^CHECK NO"
            .ColWidth(0) = 200
            .ColWidth(1) = (.Width - 300) * 0.05 'DATE
            .ColWidth(2) = (.Width - 300) * 0.05  'RFP NO.
            .ColWidth(3) = (.Width - 300) * 0.14   'SOLDTO
            .ColWidth(4) = (.Width - 300) * 0.04  'ITEMS
            .ColWidth(5) = (.Width - 300) * 0.1  'PREPARED BY
            .ColWidth(6) = (.Width - 300) * 0.1 'CHECKED BY
            .ColWidth(7) = (.Width - 300) * 0.1 'APPROVED BY
            .ColWidth(8) = (.Width - 300) * 0.14 'STATUS
            .ColWidth(9) = (.Width - 300) * 0.08 'STATUTORY
            .ColWidth(10) = (.Width - 300) * 0.1 'TOTAL PAYABLE
            .ColWidth(11) = (.Width - 300) * 0.07  'Print Copy.
            .ColWidth(12) = (.Width - 300) * 0.14 'BANK ACCT 0 'RFP ID.
            .ColWidth(13) = (.Width - 300) * 0.14 'BANK ACCT
            .ColWidth(14) = (.Width - 300) * 0 'CHECK NO
            With flxSum
                For i = 1 To .Rows - 1
                    If Val(.TextMatrix(i, 8)) = 0 Then
                        .col = 1: .row = i: .ColSel = 13: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    Else
                        .col = 1: .row = i: .ColSel = 13: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = vbWhite
                    End If
                    
                    If .TextMatrix(i, 8) = "CANCELLED" Then
                        .col = 1: .row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
                    End If
                Next i
            End With
        End If
    End With
End Sub


Private Sub Timer1_Timer()
    Dim i As Integer
    
    If MySource = "PRODUCTS" Then
        FixRange = True
        DailyRange = False
        MonthlyRange = False
        MaxCol = 8
    ElseIf MySource = "BRANCH BACK ORDER" Then
        FixRange = False
        DailyRange = True
        MonthlyRange = False
        MaxCol = 13
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto.Value = Date
        ShowDate = 2
        cmdCheck.Visible = True
    ElseIf MySource = "CLOSED ACCOUNTS FOR AGENT" Then
        FixRange = False
        DailyRange = False
        MonthlyRange = True
        MaxCol = 7
    ElseIf MySource = "BRANCH RECEIVING DISCREPANCY" Then
        FixRange = False
        DailyRange = False
        MonthlyRange = False
        BranchRange = True
        MaxCol = 12
    ElseIf MySource = "PRICING" Then
        FixRange = False
        DailyRange = True
        MonthlyRange = False
        MaxCol = 7
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto.Value = Date
        ShowDate = 2
        cmdShow.Visible = False
        cmdPrint.Visible = False
        cmdFilter.Visible = False
        cmdExit.Left = cmdShow.Left
    ElseIf MySource = "PARTNUMBER" Then
        FixRange = False
        DailyRange = True
        MonthlyRange = False
        MaxCol = 6
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto.Value = Date
        ShowDate = 2
        cmdShow.Visible = False
        cmdPrint.Visible = False
        cmdFilter.Visible = False
        cmdExit.Left = cmdShow.Left
    ElseIf MySource = "DETAILED STOCKTRANSFER" Then
        FixRange = False
        DailyRange = True
        MonthlyRange = False
        MaxCol = 12
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto.Value = Date
        cmdShow.Visible = False
        cmdSaveExcel.Visible = True
    ElseIf MySource = "MONTHLY MERCHANDISE INVENTORY" Then
        FixRange = False
        DailyRange = False
        MonthlyRange = True
        MaxCol = 8
    ElseIf MySource = "DETAILED SUPPLIER SERVICE RATIO" Then
        FixRange = False
        DailyRange = True
        MonthlyRange = False
        MaxCol = 12
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto.Value = Date
        cmdShow.Visible = False
        cmdSaveExcel.Visible = True
    ElseIf MySource = "RFPOTHER" Then
        FixRange = False
        DailyRange = True
        MonthlyRange = False
        cmdFilter.Enabled = False
        TCol = 14
        Dto.Value = Date
        Dfrom.Value = DateAdd("D", -30, Date)
        ShowDate = 2
    End If
    
        
    If DailyRange Then
        Frame1.Visible = True
        If MySource = "DETAILED STOCKTRANSFER" Then
            Frame2.Visible = True
        ElseIf MySource = "RFPOTHER" Then
            Frame2.Visible = True
        Else
            Frame2.Visible = False
        End If
        Frame3.Visible = False
        Frame4.Visible = False
    ElseIf MonthlyRange Then
        Frame1.Visible = False
        Frame2.Visible = True
        Frame3.Visible = True
        Frame4.Visible = False
        For i = 1 To 12
            cboMonth.AddItem MonthName(i)
        Next i
        'UpdYear.Value = Year(Date)
        txtYear.Text = Year(Date)
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto.Value = DateAdd("d", -1, DateAdd("m", 1, Dfrom.Value))
        cboMonth.ListIndex = Month(Dfrom.Value) - 1
    ElseIf FixRange Then
        Frame1.Visible = False
        Frame2.Visible = True
        Frame3.Visible = False
        Frame4.Visible = False
    ElseIf BranchRange Then
        Frame1.Visible = False
        Frame2.Visible = True
        Frame3.Visible = False
        Frame4.Visible = True
        Qry = "SELECT branchID , BRANCHNAME From Branch WHERE CATEGORY = 'MRC' AND SYSTEMBRANCH <> 1;"
            With cboBranch
                rsAnyUse.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                    If rsAnyUse.RecordCount > 0 Then
                        rsAnyUse.MoveFirst
                        Do While Not rsAnyUse.EOF
                            .AddItem rsAnyUse!branchName
                            .ItemData(.ListCount - 1) = rsAnyUse!branchID
                            rsAnyUse.MoveNext
                        Loop
                    End If
                rsAnyUse.Close
            End With
    End If
    
    If BranchRange = False Then
        DoEvents
        Call cmdRefresh_Click
    End If
    Timer1.Enabled = False
End Sub



Private Sub txtFind_Change()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub UpdYear_DownClick()
    If MySource = "MONTHLY MERCHANDISE INVENTORY" Then
        txtYear.Text = Val(txtYear.Text) - 1
    End If
End Sub

Private Sub UpdYear_UpClick()
    If MySource = "MONTHLY MERCHANDISE INVENTORY" Then
        txtYear.Text = Val(txtYear.Text) + 1
    End If

End Sub
