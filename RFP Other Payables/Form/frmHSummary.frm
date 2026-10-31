VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmHSummary 
   Caption         =   "Summary"
   ClientHeight    =   8532
   ClientLeft      =   -528
   ClientTop       =   -1728
   ClientWidth     =   15156
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8532
   ScaleWidth      =   15156
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboPrepBy 
      BackColor       =   &H00FFFFFF&
      Height          =   288
      ItemData        =   "frmHSummary.frx":0000
      Left            =   1440
      List            =   "frmHSummary.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   46
      Tag             =   "xyf"
      Top             =   8040
      Visible         =   0   'False
      Width           =   2772
   End
   Begin VB.Frame framStatus 
      Height          =   735
      Left            =   6840
      TabIndex        =   42
      Tag             =   "xyf"
      Top             =   720
      Visible         =   0   'False
      Width           =   2655
      Begin VB.ComboBox cboStat 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmHSummary.frx":0004
         Left            =   840
         List            =   "frmHSummary.frx":0011
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblStat 
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.TextBox txt 
      Height          =   285
      Left            =   2880
      TabIndex        =   41
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtdr 
      Height          =   285
      Left            =   1440
      TabIndex        =   40
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxUpdateDr 
      Height          =   2535
      Left            =   15240
      TabIndex        =   39
      Tag             =   "XYF"
      Top             =   2760
      Width           =   4095
      _ExtentX        =   7218
      _ExtentY        =   4466
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2292
      Left            =   5400
      TabIndex        =   25
      Top             =   4200
      Visible         =   0   'False
      Width           =   5892
      _ExtentX        =   10393
      _ExtentY        =   4043
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmHSummary.frx":002A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdprosfil"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdvisble"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "DTPicker2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "DTPicker1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cbofilter"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.ComboBox cbofilter 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1800
         TabIndex        =   26
         Tag             =   "XYF"
         Text            =   " "
         Top             =   1320
         Width           =   3735
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   1800
         TabIndex        =   27
         Tag             =   "xyf"
         Top             =   840
         Width           =   1815
         _ExtentX        =   3196
         _ExtentY        =   529
         _Version        =   393216
         Format          =   96468993
         CurrentDate     =   42473
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   300
         Left            =   3720
         TabIndex        =   28
         Tag             =   "xyf"
         Top             =   840
         Width           =   1815
         _ExtentX        =   3196
         _ExtentY        =   529
         _Version        =   393216
         Format          =   89849857
         CurrentDate     =   42473
      End
      Begin lvButton.lvButtons_H cmdvisble 
         Height          =   375
         Left            =   3840
         TabIndex        =   29
         Top             =   1800
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   656
         Caption         =   "&Exit"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         Image           =   "frmHSummary.frx":0046
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdprosfil 
         Height          =   375
         Left            =   2160
         TabIndex        =   30
         Top             =   1800
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   656
         Caption         =   "&FILTER"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         Image           =   "frmHSummary.frx":05E0
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdExcelBudgetAlo 
         Height          =   375
         Index           =   0
         Left            =   2280
         TabIndex        =   36
         Tag             =   "XYF"
         Top             =   -2520
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   656
         Caption         =   "&Export to Excel"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         Image           =   "frmHSummary.frx":5DD2
         cBack           =   -2147483633
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Please fill up all necessary data needed for filtering"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   34
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "SELECT CATEGORY"
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
         Left            =   1320
         TabIndex        =   33
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Date Covered:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   32
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   31
         Top             =   1320
         Width           =   1335
      End
   End
   Begin lvButton.lvButtons_H cmdExcel 
      Height          =   375
      Index           =   2
      Left            =   9600
      TabIndex        =   37
      Tag             =   "XYF"
      Top             =   960
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   656
      Caption         =   "&Export to Excel"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      Image           =   "frmHSummary.frx":B5C4
      cBack           =   -2147483633
   End
   Begin TabDlg.SSTab tabwait 
      Height          =   1092
      Left            =   6000
      TabIndex        =   22
      Top             =   3000
      Visible         =   0   'False
      Width           =   5292
      _ExtentX        =   9335
      _ExtentY        =   1926
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmHSummary.frx":10DB6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Processing Data, Please be patient"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   23
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label6 
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
         TabIndex        =   24
         Top             =   120
         Width           =   3255
      End
      Begin VB.Image Image1 
         Height          =   576
         Left            =   240
         Picture         =   "frmHSummary.frx":10DD2
         Top             =   120
         Width           =   576
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8160
      Top             =   0
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   656
      Caption         =   "E&xit Summary"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      Image           =   "frmHSummary.frx":1178F
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdShow 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   656
      Caption         =   "Sho&w Form"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      Image           =   "frmHSummary.frx":121A1
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   656
      Caption         =   "&Print Preview"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      Image           =   "frmHSummary.frx":1273B
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdFilter 
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   656
      Caption         =   "&Filter Records"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      Image           =   "frmHSummary.frx":1314D
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Tag             =   "xyf"
      Top             =   720
      Visible         =   0   'False
      Width           =   7335
      Begin MSComCtl2.DTPicker Dfrom 
         Height          =   300
         Left            =   1440
         TabIndex        =   7
         Tag             =   "xyf"
         Top             =   240
         Width           =   1815
         _ExtentX        =   3196
         _ExtentY        =   529
         _Version        =   393216
         Format          =   97320961
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
         _ExtentY        =   699
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         Image           =   "frmHSummary.frx":13B5F
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker Dto 
         Height          =   300
         Left            =   3840
         TabIndex        =   9
         Tag             =   "xyf"
         Top             =   240
         Width           =   1815
         _ExtentX        =   3196
         _ExtentY        =   529
         _Version        =   393216
         Format          =   97320961
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
         _ExtentY        =   699
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         Image           =   "frmHSummary.frx":14571
         ImgSize         =   24
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
         _ExtentY        =   699
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         Image           =   "frmHSummary.frx":14F83
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date Covered:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Tag             =   "xyf"
         Top             =   285
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Tag             =   "XYF"
      Top             =   720
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
         _ExtentX        =   445
         _ExtentY        =   550
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSum 
      Height          =   6372
      Left            =   120
      TabIndex        =   5
      Tag             =   "xyf"
      Top             =   1560
      Width           =   14892
      _ExtentX        =   26268
      _ExtentY        =   11240
      _Version        =   393216
      RowHeightMin    =   300
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   7560
      TabIndex        =   12
      Tag             =   "xyf"
      Top             =   720
      Width           =   7455
      Begin VB.TextBox txtFind 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
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
         _ExtentY        =   699
         Caption         =   "Find Next"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         Image           =   "frmHSummary.frx":15995
         cBack           =   -2147483633
      End
   End
   Begin MSComDlg.CommonDialog DlgDemo 
      Left            =   8760
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblPrepBy 
      Caption         =   "Prepared By:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   45
      Tag             =   "xyf"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Label txtTTLamount 
      Height          =   372
      Left            =   10920
      TabIndex        =   38
      Top             =   8040
      Width           =   3732
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      Caption         =   "Form SUMMARY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   9360
      TabIndex        =   20
      Top             =   360
      Width           =   5655
   End
   Begin VB.Label txtTOTAL 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      TabIndex        =   35
      Tag             =   "xyf"
      Top             =   9240
      Width           =   3375
   End
   Begin VB.Label lblHeading 
      Caption         =   "LYRGCOM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   21
      Top             =   480
      Width           =   7095
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   7680
      Top             =   0
      _ExtentX        =   677
      _ExtentY        =   677
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   15372
      m_MDI_Height    =   9072
   End
   Begin VB.Menu mnuBtn 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuUpdateDev 
         Caption         =   "Update Delivery"
      End
      Begin VB.Menu mnuConShipto 
         Caption         =   "Confirm Branch Allocation"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDisSet 
         Caption         =   "Discrepancy Settlement"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuConfirm 
         Caption         =   "Confirm Delivery"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu menuCM 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuCm 
         Caption         =   "CM"
      End
   End
   Begin VB.Menu MenuCancel 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuCancel 
         Caption         =   "cancel transaction"
      End
   End
End
Attribute VB_Name = "frmHSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSummary As New ADODB.Recordset
Dim rsFilteredProducts As New ADODB.Recordset
Dim DailyRange As Boolean, MonthlyRange As Boolean, FixRange As Boolean
Dim Qry As String, rsSupp As New ADODB.Recordset
Dim ActRow As Integer, fRow As Integer, fCol As Integer, TCol As Integer, ShowDate As Integer
Public MySource As String, DateStart As String, DateEnd As String
Dim ComboSrch As Boolean, TotalPO As Currency


Public Sub RefreshRecord(dCode As Integer)
    Dim i As Integer
    
    'Call RXResizer1_ResizeComplete
    
    If MySource = "PRODUCTS" Then
        Qry = "SELECT productline.PRODCODE,productline.MODEL,productline.MODELGROUP,productline.DESCRIPTION, " & _
              "productline.PRODTYPE,productline.PRODINIT,productline.ACCTITLE,productline.BRAND,productline.CLASSIFICATION, " & _
              "productline.ProdID From  productline where productline.main_prodid = 0 Order By  productline.model"
    ElseIf MySource = "RFP" Then
        'Qry = "SELECT rfp.RFPDATE, rfp.RFPNO,supplier.SUPPLIERNAME,rfp.CHECKNO,rfp.CHECKAMNT,rfp.DUEDATE,rfp.PREPBY,rfp.REVBY,rfp.APPBY,RFID FROM rfp INNER JOIN supplier ON (rfp.SUPPLIERID = supplier.SUPPLIERID) WHERE rfp.RFPDATE >= '" & Format(DFROM.Value, "YYYY-MM-DD") & "' AND rfp.RFPDATE <= '" & Format(DTO.Value, "YYYY-MM-DD") & "' AND rfp.RTYPE ='COD' ORDER BY rfp.RFPDATE"
        Qry = "CALL `sp_rfpcodsum`( '" & Format(Dfrom.Value, "YYYY-MM-DD") & "', '" & Format(Dto.Value, "YYYY-MM-DD") & "');"
    
    ElseIf MySource = "CBD" Then
        Qry = "SELECT rfp.RFPDATE, rfp.RFPNO,supplier.SUPPLIERNAME,rfp.status,rfp.CHECKAMNT,rfp.DUEDATE,rfp.PREPBY,rfp.REVBY,rfp.APPBY,rfp.RFID,rfp.PrintCopy FROM rfp INNER JOIN supplier ON (rfp.SUPPLIERID = supplier.SUPPLIERID) WHERE rfp.RFPDATE >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' AND rfp.RFPDATE <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' " & _
        "AND rfp.RTYPE ='CBD'AND rfp.POTYPE ='AP' ORDER BY rfp.RFPDATE"
        
    ElseIf MySource = "MCCBD" Then
           Qry = "SELECT rfp.RFPDATE, rfp.RFPNO,supplier.SUPPLIERNAME,rfp.status,rfp.CHECKAMNT,rfp.DUEDATE,rfp.PREPBY,rfp.REVBY,rfp.APPBY,rfp.RFID,rfp.PrintCopy FROM rfp INNER JOIN supplier ON (rfp.SUPPLIERID = supplier.SUPPLIERID) WHERE rfp.RFPDATE >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' AND rfp.RFPDATE <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' " & _
        "AND rfp.RTYPE ='CBD' AND rfp.POTYPE ='MC'ORDER BY rfp.RFPDATE"
    
    ElseIf MySource = "MOTORCYCLERFP" Then
           Qry = "CALL `sp_rfpcodsumMOTOR`( '" & Format(Dfrom.Value, "YYYY-MM-DD") & "', '" & Format(Dto.Value, "YYYY-MM-DD") & "');"
    
    ElseIf MySource = "BACK ORDER" Then
        Qry = "SELECT DATE_FORMAT(backorder.TRANSDATE,'%c/%d/%Y'), backorder.REFERENCE, spareparts.PARTNUMBER, spareparts.DESCRIPTION, backorder.BACKORDERQTY, backorder.BOBALANCE, backorder.BRANCHNAME, backorder.SUPPORDERREF " & _
              "From backorder INNER JOIN spareparts ON (backorder.SPAREPARTID = spareparts.SPAREPARTID) " & _
              "WHERE backorder.TRANSDATE >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' AND backorder.TRANSDATE <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' ORDER BY backorder.TRANSDATE DESC"
'    ElseIf MySource = "CHECKPO" Then
'        Qry = "call lyrpurchase.`sp_Checkpo`( '" & Format(Dfrom.Value, "YYYY-MM-DD") & "', '" & Format(Dto.Value, "YYYY-MM-DD") & "');"
    ElseIf MySource = "CHART OF ACCOUNTS" Then
        Qry = "SELECT chartaccts.ACCTNO,chartaccts.ACCTTITLE,chartaccts.ACCTTYPE,chartaccts.GRPACCT, " & _
            "chartaccts.ACCTRPT,chartaccts.NBAL,chartaccts.REMARKS,chartaccts.`STATUS`,chartaccts.ACCTID " & _
            "From  chartaccts"
    ElseIf MySource = "INST-ACCT" Then
        If dCode = 1 Then
            Qry = "SELECT " & _
            "DATE_FORMAT(SalesReport.INVDATE,'%c/%d/%Y'), " & _
            "SalesReport.ACCOUNTNO, " & _
            "SalesReport.CUSTOMER, " & _
            "SalesReport.ADDRESS, " & _
            "SalesReport.UNITDESC, " & _
            "SalesReport.COD, " & _
            "SalesReport.LCP, " & _
            "SalesReport.INVAMOUNT, " & _
            "SalesReport.INVTYPE, " & _
            "SalesReport.TERMS, " & _
            "SalesReport.SALESAGENT, " & _
            "SalesReport.COMMISSION,  " & _
            "SalesReport.DATERELEASED, " & _
            "SalesReport.PCVNO, " & _
            "SalesReport.REBCOMM, " & _
            "unitordered.MGRREMARKS " & _
            "From  " & _
            "SalesReport " & _
            "INNER JOIN unitordered ON (SalesReport.ACCOUNTNO = unitordered.ACCOUNTNO) " & _
            "Where " & _
            "SalesReport.INVDATE = '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' " & _
            "AND salesreport.INVDATE >= '2015-01-01' " & _
            "AND  salesreport.INVDATE <= '2015-04-01' AND  unitordered.MGRREMARKS LIKE '%INSTITUTIONAL%' " & _
            "Order By  salesreport.INVDATE,  salesreport.INVTYPE,  SalesReport.ACCOUNTNO "
        Else
            Qry = "SELECT " & _
            "DATE_FORMAT(SalesReport.INVDATE,'%c/%d/%Y'),  " & _
            "SalesReport.ACCOUNTNO,  " & _
            "SalesReport.CUSTOMER, " & _
            "SalesReport.ADDRESS, " & _
            "SalesReport.UNITDESC, " & _
            "SalesReport.COD,  " & _
            "SalesReport.LCP,  " & _
            "SalesReport.INVAMOUNT, " & _
            "SalesReport.INVTYPE, " & _
            "SalesReport.TERMS, " & _
            "SalesReport.SALESAGENT, " & _
            "SalesReport.COMMISSION, " & _
            "SalesReport.DATERELEASED, " & _
            "SalesReport.PCVNO, " & _
            "SalesReport.REBCOMM, " & _
            "unitordered.MGRREMARKS " & _
            "From  SalesReport  INNER JOIN unitordered ON (SalesReport.ACCOUNTNO = unitordered.ACCOUNTNO) " & _
            "Where SalesReport.INVDATE >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' " & _
            "AND salesreport.INVDATE >= '2015-01-01' AND  salesreport.INVDATE <= '2015-04-01' " & _
            "AND SalesReport.INVDATE <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' " & _
            "AND  unitordered.MGRREMARKS LIKE '%INSTITUTIONAL%' " & _
            "Order By  salesreport.INVDATE,  salesreport.INVTYPE,  SalesReport.ACCOUNTNO "
        End If
        
'        If MySource = "Receiving Receipt" Then
'        If dCode = 1 Then
'            Qry = "SELECT DATE_FORMAT(invreceiving.RRDATE,'%c/%d/%Y'), invreceiving.RRNO, invreceiving.RECFROM, invreceiving.DOCREFNO, DATE_FORMAT(invreceiving.DOCREFDATE,'%c/%d/%Y'), invreceiving.CHECKBY, invreceiving.REMARKS, invreceiving.PRNCOPY, invreceiving.RRECID " & _
'                  "From invreceiving Where invreceiving.RRDATE = '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' Order By invreceiving.RRDATE DESC, invreceiving.RRno;"
'        Else
'            Qry = "SELECT DATE_FORMAT(invreceiving.RRDATE,'%c/%d/%Y'), invreceiving.RRNO, invreceiving.RECFROM, invreceiving.DOCREFNO, DATE_FORMAT(invreceiving.DOCREFDATE,'%c/%d/%Y'), invreceiving.CHECKBY, invreceiving.REMARKS, invreceiving.PRNCOPY, invreceiving.RRECID " & _
'                  "From invreceiving Where invreceiving.RRDATE >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' AND invreceiving.RRDATE <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' Order By invreceiving.RRDATE DESC, invreceiving.RRno;"
'        End If

    
    ElseIf MySource = "BRANCH" Then
 
        Qry = "SELECT branch.BRANCHNAME,branch.BRANCHCODE,branch.CGROUP,branch.MANAGER,branch.ADDRESS,branch.CONTACTNOS,branch.FAXNO,branch.SPARK,branch.BRANCHTYPE,branch.TIN,branch.DATESTART,branch.REMARKS,branch.EMAIL,branch.PRODUCTS,branch.`AREA`,branch.MAIN,branch.DSNNAME,branch.BRANCHID " & _
              "From branch " & _
              "Order By  branch.BRANCHNAME"
    ElseIf MySource = "CHART OF ACCOUNTS" Then
        Qry = "SELECT chartaccts.ACCTID,chartaccts.ACCTNO,chartaccts.ACCTTITLE,chartaccts.ACCTTYPE,chartaccts.GRPACCT, " & _
            "chartaccts.ACCTRPT,chartaccts.NBAL,chartaccts.REMARKS,chartaccts.`STATUS` " & _
            "From chartaccts"
    
    ElseIf MySource = "AGENT" Then
        If dCode = 1 Then
            Qry = "SELECT CONCAT(agents.LASTNAME, ', ', agents.FIRSTNAME) AS AGENT,agentorders.RANK,DATE_FORMAT(unitordered.DATECLOSED,'%c/%d/%Y'),unitordered.ACCOUNTNO,CONCAT(customer.LASTNAME, ', ', customer.FIRSTNAME) AS customername,CONCAT(customer.ADDRESS, ' ', customer.TOWN) AS custAddress,concat(productline.BRAND, ' ', productline.DESCRIPTION, ' ', unitordered.MODEL) AS UnitDesc,agingprev.REMARKS " & _
                "From agingprev INNER JOIN unitordered ON (agingprev.ORDERID = unitordered.ORDERID)INNER JOIN agentorders ON (agingprev.ORDERID = agentorders.ORDERID) INNER JOIN agents ON (agentorders.AGENTID = agents.AGENTID) INNER JOIN customer ON (unitordered.CUSTOMERID = customer.CUSTOMERID) INNER JOIN productline ON (unitordered.PRODID = productline.PRODID) " & _
                "Where (agentorders.RANK = 'USM' OR agentorders.RANK = 'USS') AND unitordered.DATECLOSED = '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' AND unitordered.DATECLOSED >= '2015-01-01' AND unitordered.DATECLOSED < '2015-02-01' AND agingprev.REMARKS = 'CLOSED' AND agingprev.`STATUS` = 'GOOD' " & _
                "Order By agentorders.RANK,AGENT"
        Else
            Qry = "SELECT CONCAT(agents.LASTNAME, ', ', agents.FIRSTNAME) AS AGENT,agentorders.RANK,DATE_FORMAT(unitordered.DATECLOSED,'%c/%d/%Y'),unitordered.ACCOUNTNO,CONCAT(customer.LASTNAME, ', ', customer.FIRSTNAME) AS customername,CONCAT(customer.ADDRESS, ' ', customer.TOWN) AS custAddress,concat(productline.BRAND, ' ', productline.DESCRIPTION, ' ', unitordered.MODEL) AS UnitDesc,agingprev.REMARKS " & _
                "From agingprev INNER JOIN unitordered ON (agingprev.ORDERID = unitordered.ORDERID)INNER JOIN agentorders ON (agingprev.ORDERID = agentorders.ORDERID) INNER JOIN agents ON (agentorders.AGENTID = agents.AGENTID) INNER JOIN customer ON (unitordered.CUSTOMERID = customer.CUSTOMERID) INNER JOIN productline ON (unitordered.PRODID = productline.PRODID) " & _
                "Where (agentorders.RANK = 'USM' OR agentorders.RANK = 'USS') AND unitordered.DATECLOSED >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' AND unitordered.DATECLOSED >= '2015-01-01' AND unitordered.DATECLOSED < '2015-02-01' AND unitordered.DATECLOSED <= '" & Format(Dto.Value, "YYYY-MM-DD") & "'AND agingprev.REMARKS = 'CLOSED' AND agingprev.`STATUS` = 'GOOD' " & _
                "Order By agentorders.RANK,AGENT"
        End If

    ElseIf MySource = "SUPPLIER" Then
        Qry = "SELECT * From  supplier Order By  supplier.SUPPLIERNAME"
    ElseIf MySource = "FINDPO" Then
        Qry = "SELECT * From posbrnref "
        
    ElseIf MySource = "PRCM" Then
            Qry = "CALL sp_cmpSum( '" & Format(Dfrom.Value, "YYYY-MM-DD") & "', '" & Format(Dto.Value, "YYYY-MM-DD") & "');"
    ElseIf MySource = "PRDM" Then
            Qry = "CALL sp_dmpSum( '" & Format(Dfrom.Value, "YYYY-MM-DD") & "', '" & Format(Dto.Value, "YYYY-MM-DD") & "');"
    ElseIf MySource = "PURCHASE ORDER" Then
        If dCode = 1 Then
           Qry = "CALL sp_POsummary( '" & Format(Dfrom.Value, "YYYY-MM-DD") & "');"
        Else
            Qry = "CALL sp_POSUMMARY2( '" & Format(Dfrom.Value, "YYYY-MM-DD") & "', '" & Format(Dto.Value, "YYYY-MM-DD") & "');"
        End If
        
    ElseIf MySource = "PR" Then
        If dCode = 1 Then
           Qry = "CALL sp_PRsummary( '" & Format(Dfrom.Value, "YYYY-MM-DD") & "');"
        Else
            Qry = "CALL sp_PR2SUMMARY2app( '" & Format(Dfrom.Value, "YYYY-MM-DD") & "', '" & Format(Dto.Value, "YYYY-MM-DD") & "');"
        End If
        
    ElseIf MySource = "PRMC" Then
        If dCode = 1 Then
           Qry = "CALL sp_PR2summary( '" & Format(Dfrom.Value, "YYYY-MM-DD") & "','MC');"
        Else
            Qry = "CALL sp_PR2SUMMARY2( '" & Format(Dfrom.Value, "YYYY-MM-DD") & "', '" & Format(Dto.Value, "YYYY-MM-DD") & "','MC');"
        End If
'----------------------------07/10/2019- ARIAS --------------------
    ElseIf MySource = "PURCHASE ORDER2" Then
        If dCode = 1 Then
            Qry = "SELECT DATE_FORMAT(posupplier.PODATE,'%c/%d/%Y'), posupplier.PONUMBER, posupplier.TERMS, posupplier.`STATUS`, posupplier.SUPPLIER, posupplier.AUTHDEALER, posupplier.REMARKS, posupplier.PREPBY, posupplier.PURCHASEID From posupplier " & _
                  "Where posupplier.PODATE = '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' and posupplier.`TYPE` ='MC' Order By posupplier.PODATE Desc"
        Else
'            Qry = "SELECT DATE_FORMAT(posupplier.PODATE,'%c/%d/%Y'), posupplier.PONUMBER, posupplier.TERMS, posupplier.`STATUS`, posupplier.SUPPLIER, posupplier.AUTHDEALER, posupplier.REMARKS, posupplier.PREPBY, posupplier.PURCHASEID From posupplier " & _
'                  "Where posupplier.PODATE >= '" & Format(DFROM.Value, "YYYY-MM-DD") & "' AND posupplier.PODATE <= '" & Format(DTO.Value, "YYYY-MM-DD") & "' and (posupplier.category like '%MOTORCYCLE%' OR posupplier.category like '%THREE-WHEELED%' OR posupplier.category LIKE '%HELMET%') Order By posupplier.PODATE Desc"
        '  InputBox Qry, , Qry
        Qry = "CALL sp_POSUMMARY2MC( '" & Format(Dfrom.Value, "YYYY-MM-DD") & "', '" & Format(Dto.Value, "YYYY-MM-DD") & "');"
        End If
    ElseIf MySource = "POBILLING" Then
        Qry = "CALL `sp_BillingSummary`( '" & Format(Dfrom.Value, "YYYY-MM-DD") & "', '" & Format(Dto.Value, "YYYY-MM-DD") & "');"
    ElseIf MySource = "POBILLINGMC" Then
        Qry = "CALL `sp_BillingMCSummary`( '" & Format(Dfrom.Value, "YYYY-MM-DD") & "', '" & Format(Dto.Value, "YYYY-MM-DD") & "');"
    ElseIf MySource = "Price List" Then
'----------------------------08/15/2018 ARIAS ---------------------
'        Qry = "SELECT DATE_FORMAT(EFFECTDATE,'%c/%d/%Y'), REFNO, PRODTYPE, BRAND, DESCRIPTION, ACTIVEPRICE, FINAL, PRICELISTID " & _
'              "From PRICELIST ORDER BY EFFECTDATE DESC, PRODTYPE ASC, BRAND ASC;"
        Qry = "CALL `sp_pricelistshow`();"
'------------------------------ARIAS------------------------------------
    ElseIf MySource = "Employees" Then
        Qry = "SELECT * From employees Order By EMPLOYEES.LASTNAME"
'------------------------------1/09/2018--------------------------------
    ElseIf MySource = "BRANCH ALLOCATION" Then
        Qry = "CALL sp_allocation_summary( '" & Format(Dfrom.Value, "YYYY-MM-DD") & "', '" & Format(Dto.Value, "YYYY-MM-DD") & "');"
'------------------------------ARIAS---------------------
    
    ElseIf MySource = "SPCRR" Then
        'Qry = "CALL sp_spcrr_summary( '" & Format(Dfrom.Value, "YYYY-MM-DD") & "', '" & Format(Dto.Value, "YYYY-MM-DD") & "', '" & cboStat.Text & "');"
            
        If cboStat.Text = "ALL" Then
            Qry = "SELECT DATECREATED, SPCRRNO, RRNO, RRDATE, BRANCH, SUPPLIER,PREPAREDBY, TERMS, INVSTATUS, INVOICETTLPAY, SPCRRID FROM " & _
                  "spcrrmain WHERE DATECREATED >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' AND DATECREATED <= '" & Format(Dto.Value, "YYYY-MM-DD") & "';"
        Else
             Qry = "SELECT DATECREATED, SPCRRNO, RRNO, RRDATE, BRANCH, SUPPLIER,PREPAREDBY, TERMS, INVSTATUS, INVOICETTLPAY, SPCRRID FROM " & _
                  "spcrrmain WHERE DATECREATED >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' AND DATECREATED <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' AND INVSTATUS = '" & cboStat.Text & "';"
        End If

    ElseIf MySource = "FA" Then
        
        If dCode = 1 Then
            Qry = "SELECT a.RFPDATE,a.RFPNO, b.SOLDTO,COUNT(b.RFPID) as ITEMS," & _
                  "a.PREPBY,a.REVBY,a.APPBY,a.`STATUS`, a.POTYPE,FORMAT(SUM(b.POAMNT),2) as NETPAY, a.PrintCopy, a.RFID " & _
                  "FROM rfp as a INNER JOIN rfpdetails as b ON a.RFID = b.RFPID WHERE a.RFPDATE = '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' " & _
                  "AND a.RTYPE = 'GSSD' AND a.POTYPE = 'FA' GROUP BY a.RFPNO ORDER BY RFPDATE ASC;"
        Else
            Qry = "SELECT a.RFPDATE,a.RFPNO, b.SOLDTO,COUNT(b.RFPID) as ITEMS," & _
                  "a.PREPBY,a.REVBY,a.APPBY,a.`STATUS`, a.POTYPE,FORMAT(SUM(b.POAMNT),2) as NETPAY, a.PrintCopy, a.RFID " & _
                  "FROM rfp as a INNER JOIN rfpdetails as b ON a.RFID = b.RFPID WHERE a.RFPDATE >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' " & _
                  "AND a.RFPDATE <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' AND a.RTYPE = 'GSSD' AND a.POTYPE = 'FA' GROUP BY a.RFPNO ORDER BY RFPDATE ASC;"
        End If

    ElseIf MySource = "MP" Then
        
        If dCode = 1 Then
            Qry = "SELECT a.RFPDATE,a.RFPNO, b.SOLDTO,COUNT(b.RFPID) as ITEMS," & _
                  "a.PREPBY,a.REVBY,a.APPBY,a.`STATUS`, a.POTYPE,FORMAT(SUM(b.POAMNT),2) as NETPAY, a.PrintCopy, a.RFID " & _
                  "FROM rfp as a INNER JOIN rfpdetails as b ON a.RFID = b.RFPID WHERE a.RFPDATE = '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' " & _
                  "AND a.RTYPE = 'GSSD' AND a.POTYPE = 'MP' GROUP BY a.RFPNO ORDER BY RFPDATE ASC;"
        Else
            Qry = "SELECT a.RFPDATE,a.RFPNO, b.SOLDTO,COUNT(b.RFPID) as ITEMS," & _
                  "a.PREPBY,a.REVBY,a.APPBY,a.`STATUS`, a.POTYPE,FORMAT(SUM(b.POAMNT),2) as NETPAY, a.PrintCopy, a.RFID " & _
                  "FROM rfp as a INNER JOIN rfpdetails as b ON a.RFID = b.RFPID WHERE a.RFPDATE >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' " & _
                  "AND a.RFPDATE <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' AND a.RTYPE = 'GSSD' AND a.POTYPE = 'MP' GROUP BY a.RFPNO ORDER BY RFPDATE ASC;"
        End If
        
    ElseIf MySource = "RFPOTHER" Then
        
        If cboPrepBy.Text = "" Then
            If dCode = 1 Then
                Qry = "SELECT a.RFPDATE,a.RFPNO, b.SOLDTO,COUNT(b.RFPID) as ITEMS," & _
                      "a.PREPBY,a.REVBY,a.APPBY,a.`STATUS`,a.POTYPE, FORMAT(b.NETPAY,2), a.PrintCopy, a.RFID " & _
                      "FROM rfp as a INNER JOIN rfpdetails as b ON a.RFID = b.RFPID WHERE a.RFPDATE = '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' " & _
                      "AND a.RTYPE = 'OTHER' GROUP BY a.RFPNO ORDER BY RFPDATE DESC;"
            Else
                Qry = "SELECT a.RFPDATE,a.RFPNO, b.SOLDTO,COUNT(b.RFPID) as ITEMS," & _
                      "a.PREPBY,a.REVBY,a.APPBY,a.`STATUS`,a.POTYPE, FORMAT(b.NETPAY,2), a.PrintCopy, a.RFID " & _
                      "FROM rfp as a INNER JOIN rfpdetails as b ON a.RFID = b.RFPID WHERE a.RFPDATE >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' " & _
                      "AND a.RFPDATE <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' AND a.RTYPE = 'OTHER' GROUP BY a.RFPNO ORDER BY RFPDATE DESC;"
            End If
            
        Else
        
            If dCode = 1 Then
                Qry = "SELECT a.RFPDATE,a.RFPNO, b.SOLDTO,COUNT(b.RFPID) as ITEMS," & _
                      "a.PREPBY,a.REVBY,a.APPBY,a.`STATUS`,a.POTYPE, FORMAT(b.NETPAY,2), a.PrintCopy, a.RFID " & _
                      "FROM rfp as a INNER JOIN rfpdetails as b ON a.RFID = b.RFPID WHERE a.RFPDATE = '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' " & _
                      "AND a.RTYPE = 'OTHER' AND a.PREPBY = '" & cboPrepBy.Text & "' GROUP BY a.RFPNO ORDER BY RFPDATE DESC;"
            Else
                Qry = "SELECT a.RFPDATE,a.RFPNO, b.SOLDTO,COUNT(b.RFPID) as ITEMS," & _
                      "a.PREPBY,a.REVBY,a.APPBY,a.`STATUS`,a.POTYPE, FORMAT(b.NETPAY,2), a.PrintCopy, a.RFID " & _
                      "FROM rfp as a INNER JOIN rfpdetails as b ON a.RFID = b.RFPID WHERE a.RFPDATE >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' " & _
                      "AND a.RFPDATE <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' AND a.RTYPE = 'OTHER' AND a.PREPBY = '" & cboPrepBy.Text & "' GROUP BY a.RFPNO ORDER BY RFPDATE DESC;"
            End If
        End If
        
    End If
    
   'Display Query in Flx
    If rsSummary.State = adStateOpen Then rsSummary.Close
    If MySource = "POBILLING" Then
        tabwait.Visible = True
        DoEvents
        Set rsSummary = LYRcon.Execute(Qry)
        'tabwait.Visible = False
    '------------------------------1/13/2018---------------------
    ElseIf MySource = "CHECKPO" Then
        tabwait.Visible = True
        DoEvents
        rsSummary.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        'tabwait.Visible = False
     ElseIf MySource = "BRANCH ALLOCATION" Then
        tabwait.Visible = True
        DoEvents
        rsSummary.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        'tabwait.Visible = False
    '------------------------------ARIAS---------------------
    Else
        tabwait.Visible = True
        DoEvents
        rsSummary.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        'tabwait.Visible = False
    End If
    If rsSummary.RecordCount = 0 Then   'No records to show
        flxSum.clear
        flxSum.Rows = 2
        flxSum.Cols = TCol
        Call RXResizer1_ResizeComplete
        tabwait.Visible = False
        Exit Sub
    End If
    Set flxSum.DataSource = rsSummary   'With records
    ActRow = 1
    flxSum.Visible = True
    Call RXResizer1_ResizeComplete
    If MySource = "PR" Or MySource = "PRMC" Then
            TotalPO = 0
            For i = 1 To flxSum.Rows - 1
                TotalPO = Format(Val(TotalPO) + Val(Replace(flxSum.TextMatrix(i, 10), ",", "")), "#,##0.00")
            Next i
            txtTotal.Caption = ""
            txtTotal.Caption = "OVER ALL TOTAL" & ": " & Format(TotalPO, "#,##0.00")
        '------------------------------06/19/2019---------------------
     ElseIf MySource = "POBILLING" Then
            TotalPO = 0
            For i = 1 To flxSum.Rows - 1
                TotalPO = Format(Val(TotalPO) + Val(Replace(flxSum.TextMatrix(i, 11), ",", "")), "#,##0.00")
            Next i
            txtTotal.Caption = ""
            txtTotal.Caption = "OVER ALL TOTAL" & ": " & Format(TotalPO, "#,##0.00")
'    ElseIf MySource = "POBILLINGMC" Then
'            TotalPO = 0
'            For i = 1 To flxSum.Rows - 1
'                TotalPO = Format(Val(TotalPO) + Val(Replace(flxSum.TextMatrix(i, 11), ",", "")), "#,##0.00")
'            Next i
'            txtTotal.Caption = ""
'            txtTotal.Caption = "OVER ALL TOTAL" & ": " & Format(TotalPO, "#,##0.00")
        '------------------------------1/09/2018---------------------
    ElseIf MySource = "CHECKPO" Then
            cmdShow.Enabled = False
            TotalPO = 0
            For i = 1 To flxSum.Rows - 1
                TotalPO = Format(Val(TotalPO) + Val(Replace(flxSum.TextMatrix(i, 4), ",", "")), "#,##0.00")
            Next i
            txtTotal.Caption = ""
            txtTotal.Caption = "TOTAL QTY" & ": " & Format(TotalPO, "#,##0")
    ElseIf MySource = "BRANCH ALLOCATION" Then
            cmdShow.Enabled = False
            TotalPO = 0
            For i = 1 To flxSum.Rows - 1
                TotalPO = Format(Val(TotalPO) + Val(Replace(flxSum.TextMatrix(i, 6), ",", "")), "#,##0.00")
            Next i
            txtTotal.Caption = ""
            txtTotal.Caption = "OVER ALL TOTAL" & ": " & Format(TotalPO, "#,##0.00")
        '------------------------------ARIAS---07/03/2019------------------
       ElseIf MySource = "POBILLING" Then
        For i = 1 To flxSum.Rows - 1
            If flxSum.TextMatrix(i, 8) = flxSum.TextMatrix(i, 9) Then
                With flxSum
                    .COL = 0:
                    .Row = i:
                    .ColSel = 13:
                    .RowSel = i:
                    .FillStyle = flexFillRepeat:
                    .CellBackColor = RGB(255, 255, 155)
                    .COL = 2
                End With
            End If
        Next i
'      ElseIf MySource = "POBILLINGMC" Then
'        For i = 1 To flxSum.Rows - 1
'            If flxSum.TextMatrix(i, 8) = flxSum.TextMatrix(i, 9) Then
'                With flxSum
'                    .COL = 0:
'                    .Row = i:
'                    .ColSel = 14:
'                    .RowSel = i:
'                    .FillStyle = flexFillRepeat:
'                    .CellBackColor = RGB(255, 255, 155)
'                    .COL = 2
'                End With
'            End If
'        Next i
     ElseIf MySource = "RFP" Or MySource = "MOTORCYCLERFP" Then
        For i = 1 To flxSum.Rows - 1
            If flxSum.TextMatrix(i, 5) = flxSum.TextMatrix(i, 6) Then
                With flxSum
                    .COL = 0:
                    .Row = i:
                    .ColSel = 15:
                    .RowSel = i:
                    .FillStyle = flexFillRepeat:
                    .CellBackColor = RGB(255, 255, 155)
                    .COL = 2
                End With
            End If
            
            If flxSum.TextMatrix(i, 15) = "CANCELLED" Then
                With flxSum
                    .COL = 0:
                    .Row = i:
                    .ColSel = 15:
                    .RowSel = i:
                    .FillStyle = flexFillRepeat:
                    .CellBackColor = &H8080FF
                    .COL = 2
                End With
            End If
            
        Next i
      TotalPO = 0
            For i = 1 To flxSum.Rows - 1
                TotalPO = Format(Val(TotalPO) + Val(Replace(flxSum.TextMatrix(i, 9), ",", "")), "#,##0.00")
            Next i
            txtTotal.Caption = ""
            txtTotal.Caption = "NET OVER ALL TOTAL" & ": " & Format(TotalPO, "#,##0.00")
    End If   '------------------------------ARIAS---07/03/2019------------------
    If MySource = "PR" Then
        Call checkPR
    ElseIf MySource = "PRMC" Then
        Call checkPRMC
    ElseIf MySource = "PURCHASE ORDER" Or MySource = "PURCHASE ORDER2" Then
      Call checkPO
    ElseIf MySource = "CBD" Then
        Call checkPODRinvoice
        
    ElseIf MySource = "MCCBD" Then
        
        For i = 1 To flxSum.Rows - 1
            If flxSum.TextMatrix(i, 4) = "CANCELLED" Then
                With flxSum
                    .COL = 0:
                    .Row = i:
                    .ColSel = 10:
                    .RowSel = i:
                    .FillStyle = flexFillRepeat:
                    .CellBackColor = &H8080FF
                    .COL = 2
                End With
            End If
        Next i
        
    ElseIf MySource = "PRCM" Then
        Call checkCMID
        
    ElseIf MySource = "FA" Or MySource = "MP" Or MySource = "RFPOTHER" Then
        With flxSum
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 11)) = 0 Then
                    .COL = 1: .Row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                Else
                    .COL = 1: .Row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = vbWhite
                End If
                
                If .TextMatrix(i, 8) = "CANCELLED" Then
                    .COL = 1: .Row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
                End If
            Next i
        End With
    End If
    
    tabwait.Visible = False
    
End Sub


Private Sub checkCMID()
  Dim i As Integer, Qry As String, cmNO As Variant
     With flxSum
        For i = 1 To .Rows - 1
          Qry = "SELECT cmpreturn.CMNO FROM cmpreturn WHERE cmpreturn.CMID = '" & .TextMatrix(i, 1) & "'"
'           InputBox Qry, , Qry
            cmNO = GetFieldValue(Qry)
            If cmNO = "" Then cmNO = 0
            If cmNO > 0 Then
                .COL = 0:
                .Row = i:
                .ColSel = 11:
                .RowSel = i:
                .FillStyle = flexFillRepeat:
                .CellBackColor = RGB(192, 192, 192)
                .COL = 2
            End If
        Next i
    End With
End Sub
 
Private Sub cbofilter_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    
    If ComboSrch Then
        ComboSrch = False
        Exit Sub
    End If
    With cbofilter
        strPartial = .Text
        i = SendMessage(.hWnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                ComboSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
 
    End With
    
End Sub

'Private Sub cbofilter_Click()
'    Dim Qry As String
'    Dim rsPONo As New ADODB.Recordset
'    Qry = "SELECT purchasebill.BILLTOBRN,purchasebill.STBRNCODE,DATE_FORMAT(purchasebill.DRDATE, '%b %d, %Y') AS DRDATE,purchasebill.DRNO," & _
'    "purchasebill.INVNO,DATE_FORMAT(purchasebill.DATEREC, '%b %d, %Y') AS DATEREC, FORMAT(purchasebill.TOTALAMT, 2) AS TOTALAMT,purchasebill.PBILLID " & _
'    "From lyrpurchase.purchasebill purchasebill WHERE  purchasebill.RECFROM ='" & cboSupplier.Text & "' ORDER BY purchasebill.BILLTOBRN"
'    InputBox Qry, , Qry
'    If rsPONo.State = adStateOpen Then rsPONo.Close
'    Set rsPONo = LYRcon.Execute(Qry)
'    If rsPONo.RecordCount = 0 Then
'        flxSum.clear
'        flxSum.Rows = 2
'        Exit Sub
'    End If
'    Set flxSum.DataSource = rsPONo
'    Call RXResizer1_ResizeComplete
'End Sub
'
'Private Sub cbofilter_GotFocus()
'    Dim Tmp
'
'   Const CB_SHOWDROPDOWN = &H14F
'   Tmp = SendMessage(cbofilter.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
'End Sub

Private Sub cbofilter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then ComboSrch = True
    If KeyCode = vbKeyBack Then ComboSrch = True
End Sub

Private Sub cbofilter_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdFind.SetFocus
End Sub

'Private Sub cbofilter_LostFocus()
'          Call cbofilter_Click
'End Sub

Private Sub cboSupplier_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    
    If ComboSrch Then
        ComboSrch = False
        Exit Sub
    End If
    With cbofilter
        strPartial = .Text
        i = SendMessage(.hWnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                ComboSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
 
    End With
    

End Sub

Private Sub cboSupplier_Click()
    Dim Qry As String
    Dim rsPONo As New ADODB.Recordset
    Qry = "SELECT purchasebill.BILLTOBRN,purchasebill.STBRNCODE,DATE_FORMAT(purchasebill.DRDATE, '%b %d, %Y') AS DRDATE,purchasebill.DRNO," & _
    "purchasebill.INVNO,DATE_FORMAT(purchasebill.DATEREC, '%b %d, %Y') AS DATEREC, FORMAT(purchasebill.TOTALAMT, 2) AS TOTALAMT,purchasebill.PBILLID " & _
    "From lyrpurchase.purchasebill purchasebill WHERE  purchasebill.RECFROM ='" & cboSupplier.Text & "' ORDER BY purchasebill.BILLTOBRN"
    If rsPONo.State = adStateOpen Then rsPONo.Close
    Set rsPONo = LYRcon.Execute(Qry)
    If rsPONo.RecordCount = 0 Then
        flxSum.clear
        flxSum.Rows = 2
        Exit Sub
    End If
    Set flxSum.DataSource = rsPONo
    Call RXResizer1_ResizeComplete
End Sub

Private Sub cboSupplier_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboSupplier.hWnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then ComboSrch = True
    If KeyCode = vbKeyBack Then ComboSrch = True
End Sub

Private Sub cboSupplier_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdFind.SetFocus
End Sub

Private Sub cboSupplier_LostFocus()
    Call cboSupplier_Click
End Sub
'--------------------------------ADD BY : ARIAS 06/19/2019-------------------
Private Sub cmdExcel2_Click()
    If MySource = "POBILLING" Then
          Call ExportExCelPOBILLING
    ElseIf MySource = "POBILLINGMC" Then
          Call ExportExCelPOBILLINGMC
    ElseIf MySource = "PURCHASE ORDER" Then
          Call ExportExCelPOBILLINGAPP
    ElseIf MySource = "RFP" Or MySource = "MOTORCYCLERFP" Then
          Call ExportExCelRFP
    End If
End Sub
Private Sub ExportExCelPOBILLINGMC()
''--------------------------------ARIAS-------------------
    Dim XLapp As Excel.Application, rRng As Range, rRng2 As Range
    Dim wbkMVR As Excel.Workbook
    Dim wksLineA1 As Excel.Worksheet, TotalRow As Integer '  Dim sSaveAsFileName
    Dim i As Integer, C As Integer, ActRow As Integer, NEXTROW As Integer, ROWOFFLX As Integer, Mark As Boolean


    Ans = MsgBox("Are you sure you want to Export file to Excel?", vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub

    With dlgDEMO
        .CancelError = True
        .InitDir = mstrLastDir
        .Flags = cdlOFNHideReadOnly
        .filename = ""
        .Filter = "Microsoft Excel (*.xls)|*.xls|All Files(*.*)|*.*"
        .ShowOpen
        sSaveAsFileName = .filename
    End With
    Set XLapp = CreateObject("Excel.Application")
    XLapp.Visible = True
    XLapp.Workbooks.Add
    Set wbkMVR = XLapp.ActiveWorkbook
    Set wksLineA1 = wbkMVR.Sheets(1)

    wksLineA1.Name = "SUM"
    wksLineA1.Cells(1, 1).Value = "SUPPLIERS PO SUMMARY"
    wksLineA1.Cells(2, 1).Value = "MONTH OF : " & Dfrom.Value & " to " & Dto.Value
    Set rRng = wksLineA1.Range("A1:A2") 'set range
    rRng.Font.Bold = True
    rRng.Font.Size = 10
    rRng.Font.Name = "TAHOMA"

    wksLineA1.Cells(4, 1).Value = "DRDATE"
    wksLineA1.Cells(4, 2).Value = "DRNUMBER"
    wksLineA1.Cells(4, 3).Value = "INVNUM"
    wksLineA1.Cells(4, 4).Value = "PONUM"
    wksLineA1.Cells(4, 5).Value = "BILLTO"
    wksLineA1.Cells(4, 6).Value = "SHIPTO"
    wksLineA1.Cells(4, 7).Value = "BILLING STAT"
    wksLineA1.Cells(4, 8).Value = "TTLPO"
    wksLineA1.Cells(4, 9).Value = "TTLINV"
    wksLineA1.Cells(4, 10).Value = "TTLREC"
    wksLineA1.Cells(4, 11).Value = "TTLDR AMNT"
    wksLineA1.Cells(4, 12).Value = "REC FROM"
    wksLineA1.Cells(4, 13).Value = "COD RFPNO"
    wksLineA1.Cells(4, 14).Value = "CBD RFPNO"
    
    Set rRng = wksLineA1.Range("A4:N4")
    rRng.Borders.LineStyle = xlNone
    rRng.Borders.LineStyle = xlContinuous
    rRng.HorizontalAlignment = xlVAlignCenter
    rRng.Font.Bold = True

    wksLineA1.Range(wksLineA1.Cells(4, 1), wksLineA1.Cells(4, 14)).Font.Color = vbWhite 'change font color
    wksLineA1.Range(wksLineA1.Cells(4, 1), wksLineA1.Cells(4, 14)).Interior.Color = RGB(0, 176, 80) 'make colorgreen
    wksLineA1.Range(wksLineA1.Cells(4, 1), wksLineA1.Cells(4, 14)).VerticalAlignment = xlVAlignCenter

    wksLineA1.Range(wksLineA1.Cells(4, 1), wksLineA1.Cells(4, 1)).ColumnWidth = 11.43 'DRDATE
    wksLineA1.Range(wksLineA1.Cells(4, 2), wksLineA1.Cells(4, 2)).ColumnWidth = 12#   'DRNUMBER
    wksLineA1.Range(wksLineA1.Cells(4, 3), wksLineA1.Cells(4, 3)).ColumnWidth = 11.27 'INVNUM
    wksLineA1.Range(wksLineA1.Cells(4, 4), wksLineA1.Cells(4, 4)).ColumnWidth = 9.29  'PONUM
    wksLineA1.Range(wksLineA1.Cells(4, 5), wksLineA1.Cells(4, 5)).ColumnWidth = 8.29  'BILLTO
    wksLineA1.Range(wksLineA1.Cells(4, 6), wksLineA1.Cells(4, 6)).ColumnWidth = 9.57  'SHIPTO
    wksLineA1.Range(wksLineA1.Cells(4, 7), wksLineA1.Cells(4, 7)).ColumnWidth = 17.43 'BILLING STAT
    wksLineA1.Range(wksLineA1.Cells(4, 8), wksLineA1.Cells(4, 8)).ColumnWidth = 10.14 'TTLPO
    wksLineA1.Range(wksLineA1.Cells(4, 9), wksLineA1.Cells(4, 9)).ColumnWidth = 10.14 'TTLINV
    wksLineA1.Range(wksLineA1.Cells(4, 10), wksLineA1.Cells(4, 10)).ColumnWidth = 10.14 'TTLREC
    wksLineA1.Range(wksLineA1.Cells(4, 11), wksLineA1.Cells(4, 11)).ColumnWidth = 15.57 'TTLDR AMNT
    wksLineA1.Range(wksLineA1.Cells(4, 12), wksLineA1.Cells(4, 12)).ColumnWidth = 25.29 'REC FROM
    wksLineA1.Range(wksLineA1.Cells(4, 13), wksLineA1.Cells(4, 13)).ColumnWidth = 13.14 'COD RFPNO
    wksLineA1.Range(wksLineA1.Cells(4, 14), wksLineA1.Cells(4, 14)).ColumnWidth = 13.14 'CBD RFPNO
    
    
    NEXTROW = 4
    Mark = False
    Dim TOTALAMNT As Currency
    With flxSum
       For i = 1 To .Cols - 1
            If Not (i = 13 Or i = 15 Or i = 17) Then
                If (i = 14) Then
                    ROWOFFLX = i - 1
                ElseIf (i = 17) Then
                    ROWOFFLX = i - 2
                Else
                    ROWOFFLX = i
                End If
                For C = 1 To .Rows - 1
                    wksLineA1.Cells(C + NEXTROW, ROWOFFLX) = .TextMatrix(C, i)
                If i = 11 Then
                TOTALAMNT = Val(Replace(TOTALAMNT, ",", "")) + Val(Replace(wksLineA1.Cells(C + NEXTROW, i), ",", ""))
                End If
            Next C
            If i = 10 Then wksLineA1.Cells(C + 4, 10).Value = "TOTAL"
            If i = 11 Then wksLineA1.Cells(C + 4, 11).Value = Format(TOTALAMNT, "#,###.00")
            End If
        Next i
        TotalRow = C
    End With
    Set rRng = wksLineA1.Range(wksLineA1.Cells(5, 1), wksLineA1.Cells(TotalRow + 3, 14))
    wksLineA1.Range(wksLineA1.Cells(TotalRow + 5, 9), wksLineA1.Cells(TotalRow + 3, 9)).Font.Bold = True
    wksLineA1.Range(wksLineA1.Cells(TotalRow + 5, 10), wksLineA1.Cells(TotalRow + 3, 10)).Font.Bold = True
    rRng.Font.Name = "TAHOMA"
    rRng.Borders.LineStyle = xlNone
    rRng.Borders.LineStyle = xlContinuous
    rRng.Font.Size = 10

    XLapp.DisplayAlerts = False
    wbkMVR.SaveAs filename:=sSaveAsFileName, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
'    wbkMVR.Close SaveChanges:=False
'    XLapp.Quit

    Set XLapp = Nothing
    Set wbkMVR = Nothing
    Set wksLineA1 = Nothing
    MsgBox "Finish Saving to Excel!", vbInformation + vbOKOnly, "SUM"
    
End Sub
Private Sub ExportExCelPOBILLINGAPP()
''--------------------------------ARIAS-------------------
    Dim XLapp As Excel.Application, rRng As Range, rRng2 As Range
    Dim wbkMVR As Excel.Workbook
    Dim wksLineA1 As Excel.Worksheet, TotalRow As Integer '  Dim sSaveAsFileName
    Dim i As Integer, C As Integer, ActRow As Integer, NEXTROW As Integer, ROWOFFLX As Integer, Mark As Boolean


    Ans = MsgBox("Are you sure you want to Export file to Excel?", vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub

    With dlgDEMO
        .CancelError = True
        .InitDir = mstrLastDir
        .Flags = cdlOFNHideReadOnly
        .filename = ""
        .Filter = "Microsoft Excel (*.xls)|*.xls|All Files(*.*)|*.*"
        .ShowOpen
        sSaveAsFileName = .filename
    End With
    Set XLapp = CreateObject("Excel.Application")
    XLapp.Visible = True
    XLapp.Workbooks.Add
    Set wbkMVR = XLapp.ActiveWorkbook
    Set wksLineA1 = wbkMVR.Sheets(1)

    wksLineA1.Name = "SUM"
    wksLineA1.Cells(1, 1).Value = "SUPPLIERS PO SUMMARY"
    wksLineA1.Cells(2, 1).Value = "MONTH OF : " & Dfrom.Value & " to " & Dto.Value
    Set rRng = wksLineA1.Range("A1:A2") 'set range
    rRng.Font.Bold = True
    rRng.Font.Size = 10
    rRng.Font.Name = "TAHOMA"

    wksLineA1.Cells(4, 1).Value = "PODATE"
    wksLineA1.Cells(4, 2).Value = "POREF"
    wksLineA1.Cells(4, 3).Value = "TERMS"
    wksLineA1.Cells(4, 4).Value = "STATUS"
    wksLineA1.Cells(4, 5).Value = "SUPPLIER NAME"
    wksLineA1.Cells(4, 6).Value = "REMARKS"
    wksLineA1.Cells(4, 7).Value = "TOTAL ITEM"
    wksLineA1.Cells(4, 8).Value = "TTLPOCOST"
    wksLineA1.Cells(4, 9).Value = "TTLESS"
    wksLineA1.Cells(4, 10).Value = "TOTALPOAMOUNT"
    wksLineA1.Cells(4, 11).Value = "PREPAREDBY"

    Set rRng = wksLineA1.Range("A4:K4")
    rRng.Borders.LineStyle = xlNone
    rRng.Borders.LineStyle = xlContinuous
    rRng.HorizontalAlignment = xlVAlignCenter
    rRng.Font.Bold = True

    wksLineA1.Range(wksLineA1.Cells(4, 1), wksLineA1.Cells(4, 11)).Font.Color = vbWhite 'change font color
    wksLineA1.Range(wksLineA1.Cells(4, 1), wksLineA1.Cells(4, 11)).Interior.Color = RGB(0, 176, 80) 'make colorgreen
    wksLineA1.Range(wksLineA1.Cells(4, 1), wksLineA1.Cells(4, 11)).VerticalAlignment = xlVAlignCenter

    wksLineA1.Range(wksLineA1.Cells(4, 1), wksLineA1.Cells(4, 1)).ColumnWidth = 11.43 'PODATE
    wksLineA1.Range(wksLineA1.Cells(4, 2), wksLineA1.Cells(4, 2)).ColumnWidth = 10.47 'POREF
    wksLineA1.Range(wksLineA1.Cells(4, 3), wksLineA1.Cells(4, 3)).ColumnWidth = 10.47 'TERMS
    wksLineA1.Range(wksLineA1.Cells(4, 4), wksLineA1.Cells(4, 4)).ColumnWidth = 10.43 'STATUS
    wksLineA1.Range(wksLineA1.Cells(4, 5), wksLineA1.Cells(4, 5)).ColumnWidth = 30.43 'SUPPLIER NAME
    wksLineA1.Range(wksLineA1.Cells(4, 6), wksLineA1.Cells(4, 6)).ColumnWidth = 22.57 'REMARKS
    wksLineA1.Range(wksLineA1.Cells(4, 7), wksLineA1.Cells(4, 7)).ColumnWidth = 10.47 'TOTAL ITEM
    wksLineA1.Range(wksLineA1.Cells(4, 8), wksLineA1.Cells(4, 8)).ColumnWidth = 14.29 'TTLPOCOST
    wksLineA1.Range(wksLineA1.Cells(4, 9), wksLineA1.Cells(4, 9)).ColumnWidth = 10.41 'TTLESS
    wksLineA1.Range(wksLineA1.Cells(4, 10), wksLineA1.Cells(4, 10)).ColumnWidth = 18.29 'TOTALPOAMOUNT
    wksLineA1.Range(wksLineA1.Cells(4, 11), wksLineA1.Cells(4, 11)).ColumnWidth = 19.71 'PREPAREDBY

    NEXTROW = 4
    Mark = False
    Dim TOTALAMNT As Currency
    With flxSum
       For i = 1 To .Cols - 1
            If Not (i = 12) Then
                If (i = 14) Then
                    ROWOFFLX = i - 1
                ElseIf (i = 15) Then
                    ROWOFFLX = i - 2
                Else
                    ROWOFFLX = i
                End If
                For C = 1 To .Rows - 1
                    wksLineA1.Cells(C + NEXTROW, ROWOFFLX) = .TextMatrix(C, i)
                If i = 10 Then
                TOTALAMNT = Val(Replace(TOTALAMNT, ",", "")) + Val(Replace(wksLineA1.Cells(C + NEXTROW, i), ",", ""))
                End If
            Next C
            If i = 9 Then wksLineA1.Cells(C + 4, 9).Value = "TOTAL"
            If i = 10 Then wksLineA1.Cells(C + 4, 10).Value = Format(TOTALAMNT, "#,###.00")
            End If
        Next i
        TotalRow = C
    End With
    Set rRng = wksLineA1.Range(wksLineA1.Cells(5, 1), wksLineA1.Cells(TotalRow + 3, 11))
    wksLineA1.Range(wksLineA1.Cells(TotalRow + 5, 9), wksLineA1.Cells(TotalRow + 3, 9)).Font.Bold = True
    wksLineA1.Range(wksLineA1.Cells(TotalRow + 5, 10), wksLineA1.Cells(TotalRow + 3, 10)).Font.Bold = True
    rRng.Font.Name = "TAHOMA"
    rRng.Borders.LineStyle = xlNone
    rRng.Borders.LineStyle = xlContinuous
    rRng.Font.Size = 10

    XLapp.DisplayAlerts = False
    wbkMVR.SaveAs filename:=sSaveAsFileName, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
'    wbkMVR.Close SaveChanges:=False
'    XLapp.Quit

    Set XLapp = Nothing
    Set wbkMVR = Nothing
    Set wksLineA1 = Nothing
    MsgBox "Finish Saving to Excel!", vbInformation + vbOKOnly, "SUM"
    
End Sub

Private Sub ExportExCelPOBILLING()
'--------------------------------ARIAS-------------------
    Dim XLapp As Excel.Application, rRng As Range, rRng2 As Range
    Dim wbkMVR As Excel.Workbook
    Dim wksLineA1 As Excel.Worksheet, TotalRow As Integer '  Dim sSaveAsFileName
    Dim i As Integer, C As Integer, ActRow As Integer, NEXTROW As Integer, ROWOFFLX As Integer, Mark As Boolean
   
    
    Ans = MsgBox("Are you sure you want to Export file to Excel?", vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    
    With dlgDEMO
        .CancelError = True
        .InitDir = mstrLastDir
        .Flags = cdlOFNHideReadOnly
        .filename = ""
        .Filter = "Microsoft Excel (*.xls)|*.xls|All Files(*.*)|*.*"
        .ShowOpen
        sSaveAsFileName = .filename
    End With
    Set XLapp = CreateObject("Excel.Application")
    XLapp.Visible = True
    XLapp.Workbooks.Add
    Set wbkMVR = XLapp.ActiveWorkbook
    Set wksLineA1 = wbkMVR.Sheets(1)
    
    wksLineA1.Name = "SUM"
    wksLineA1.Cells(1, 1).Value = "SUPPLIERS INVOICE SUMMARY"
    wksLineA1.Cells(2, 1).Value = "MONTH OF : " & Dfrom.Value & " to " & Dto.Value
    Set rRng = wksLineA1.Range("A1:A2") 'set range
    rRng.Font.Bold = True
    rRng.Font.Size = 10
    rRng.Font.Name = "TAHOMA"
   
    wksLineA1.Cells(4, 1).Value = "DRDATE"
    wksLineA1.Cells(4, 2).Value = "DRNUMBER"
    wksLineA1.Cells(4, 3).Value = "INVNUM"
    wksLineA1.Cells(4, 4).Value = "P.O NUM"
    wksLineA1.Cells(4, 5).Value = "BILL TO"
    wksLineA1.Cells(4, 6).Value = "SHIP TO"
    wksLineA1.Cells(4, 7).Value = "BILLING STAT"
    wksLineA1.Cells(4, 8).Value = "TTLPO"
    wksLineA1.Cells(4, 9).Value = "TTLINV "
    wksLineA1.Cells(4, 10).Value = "TTLREC "
    wksLineA1.Cells(4, 11).Value = "TTL DR AMNT"
    wksLineA1.Cells(4, 12).Value = "CODRFP"
    wksLineA1.Cells(4, 13).Value = "CBDRFP"

    Set rRng = wksLineA1.Range("A4:M4")
    rRng.Borders.LineStyle = xlNone
    rRng.Borders.LineStyle = xlContinuous
    rRng.HorizontalAlignment = xlVAlignCenter
    rRng.Font.Bold = True
    
    wksLineA1.Range(wksLineA1.Cells(4, 1), wksLineA1.Cells(4, 13)).Font.Color = vbWhite 'change font color
    wksLineA1.Range(wksLineA1.Cells(4, 1), wksLineA1.Cells(4, 13)).Interior.Color = RGB(0, 176, 80) 'make colorgreen
    wksLineA1.Range(wksLineA1.Cells(4, 1), wksLineA1.Cells(4, 13)).VerticalAlignment = xlVAlignCenter
    
    wksLineA1.Range(wksLineA1.Cells(4, 1), wksLineA1.Cells(4, 4)).ColumnWidth = 10.43 'columwitd
    wksLineA1.Range(wksLineA1.Cells(4, 5), wksLineA1.Cells(4, 6)).ColumnWidth = 10.57 'columwitd
    wksLineA1.Range(wksLineA1.Cells(4, 7), wksLineA1.Cells(4, 7)).ColumnWidth = 17.71 'columwitd
    wksLineA1.Range(wksLineA1.Cells(4, 8), wksLineA1.Cells(4, 10)).ColumnWidth = 8.57 'columwitd
    wksLineA1.Range(wksLineA1.Cells(4, 11), wksLineA1.Cells(4, 11)).ColumnWidth = 13.71 'columwitd
    wksLineA1.Range(wksLineA1.Cells(4, 12), wksLineA1.Cells(4, 13)).ColumnWidth = 10#    'columwitd
    
    NEXTROW = 4
    Mark = False
    With flxSum
       For i = 1 To .Cols - 1
            If Not (i = 12 Or i = 14 Or i = 16 Or i = 17) Then
                If (i = 13) Then
                    ROWOFFLX = i - 1
                ElseIf (i = 15) Then
                    ROWOFFLX = i - 2
                Else
                    ROWOFFLX = i
                End If
                For C = 1 To .Rows - 1
                    wksLineA1.Cells(C + NEXTROW, ROWOFFLX) = .TextMatrix(C, i)
                 Next C
             End If
        Next i
        TotalRow = C
    End With
    Set rRng = wksLineA1.Range(wksLineA1.Cells(5, 1), wksLineA1.Cells(TotalRow + 3, 13))
    rRng.Font.Name = "TAHOMA"
    rRng.Borders.LineStyle = xlNone
    rRng.Borders.LineStyle = xlContinuous
    rRng.Font.Size = 10
    rRng.HorizontalAlignment = xlVAlignCenter

    XLapp.DisplayAlerts = False
    wbkMVR.SaveAs filename:=sSaveAsFileName, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
'    wbkMVR.Close SaveChanges:=False
'    XLapp.Quit

    Set XLapp = Nothing
    Set wbkMVR = Nothing
    Set wksLineA1 = Nothing
    MsgBox "Finish Saving to Excel!", vbInformation + vbOKOnly, "SUM"
    
End Sub
Private Sub ExportExCelRFP()
'--------------------------------ARIAS-------------------
    Dim XLapp As Excel.Application, rRng As Range, rRng2 As Range
    Dim wbkMVR As Excel.Workbook
    Dim wksLineA1 As Excel.Worksheet, TotalRow As Integer '  Dim sSaveAsFileName
    Dim i As Integer, C As Integer, ActRow As Integer, NEXTROW As Integer, ROWOFFLX As Integer, Mark As Boolean
   
    
    Ans = MsgBox("Are you sure you want to Export file to Excel?", vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    
    With dlgDEMO
        .CancelError = True
        .InitDir = mstrLastDir
        .Flags = cdlOFNHideReadOnly
        .filename = ""
        .Filter = "Microsoft Excel (*.xls)|*.xls|All Files(*.*)|*.*"
        .ShowOpen
        sSaveAsFileName = .filename
    End With
    Set XLapp = CreateObject("Excel.Application")
    XLapp.Visible = True
    XLapp.Workbooks.Add
    Set wbkMVR = XLapp.ActiveWorkbook
    Set wksLineA1 = wbkMVR.Sheets(1)
    
    wksLineA1.Name = "SUM"
    wksLineA1.Cells(1, 1).Value = "REQUEST FOR PAYMENT SUMMARY"
    wksLineA1.Cells(2, 1).Value = "MONTH OF : " & Dfrom.Value & " to " & Dto.Value
    Set rRng = wksLineA1.Range("A1:A2") 'set range
    rRng.Font.Bold = True
    rRng.Font.Size = 10
    rRng.Font.Name = "TAHOMA"
   
    wksLineA1.Cells(4, 1).Value = "DATE"
    wksLineA1.Cells(4, 2).Value = "RFPNO"
    wksLineA1.Cells(4, 3).Value = "SOLDTO"
    wksLineA1.Cells(4, 4).Value = "SUPPLIER NAME"
    wksLineA1.Cells(4, 5).Value = "SI# RECIEVED"
    wksLineA1.Cells(4, 6).Value = "DR RECIEVED"
    wksLineA1.Cells(4, 7).Value = "TTL AMOUNTT"
    wksLineA1.Cells(4, 8).Value = "TTL DEDUCTION"
    wksLineA1.Cells(4, 9).Value = "NET PAYABLE"
    wksLineA1.Cells(4, 10).Value = "DUEDATE"
    wksLineA1.Cells(4, 11).Value = "PREPARED BY"
    wksLineA1.Cells(4, 12).Value = "RECIEVED BY"
    wksLineA1.Cells(4, 13).Value = "APPROVED BY"
   
    Set rRng = wksLineA1.Range("A4:M4")
    rRng.Borders.LineStyle = xlNone
    rRng.Borders.LineStyle = xlContinuous
    rRng.HorizontalAlignment = xlVAlignCenter
    rRng.Font.Bold = True
    
    wksLineA1.Range(wksLineA1.Cells(4, 1), wksLineA1.Cells(4, 13)).Font.Color = vbWhite 'change font color
    wksLineA1.Range(wksLineA1.Cells(4, 1), wksLineA1.Cells(4, 13)).Interior.Color = RGB(0, 176, 80) 'make colorgreen
    wksLineA1.Range(wksLineA1.Cells(4, 1), wksLineA1.Cells(4, 13)).VerticalAlignment = xlVAlignCenter
    wksLineA1.Range(wksLineA1.Cells(4, 5), wksLineA1.Cells(4, 5)).WrapText = True
    wksLineA1.Range(wksLineA1.Cells(4, 6), wksLineA1.Cells(4, 6)).WrapText = True
    
    wksLineA1.Range(wksLineA1.Cells(4, 6), wksLineA1.Cells(4, 2)).ColumnWidth = 10.4  'columwitd
    wksLineA1.Range(wksLineA1.Cells(4, 3), wksLineA1.Cells(4, 3)).ColumnWidth = 11.29 'columwitd
    wksLineA1.Range(wksLineA1.Cells(4, 4), wksLineA1.Cells(4, 4)).ColumnWidth = 30.09 'columwitd
    wksLineA1.Range(wksLineA1.Cells(4, 5), wksLineA1.Cells(4, 6)).ColumnWidth = 10.57 'columwitd
    wksLineA1.Range(wksLineA1.Cells(4, 7), wksLineA1.Cells(4, 8)).ColumnWidth = 13.14 'columwitd
    wksLineA1.Range(wksLineA1.Cells(4, 9), wksLineA1.Cells(4, 9)).ColumnWidth = 14.89 'columwitd
    wksLineA1.Range(wksLineA1.Cells(4, 10), wksLineA1.Cells(4, 10)).ColumnWidth = 10.71 'columwitd
    wksLineA1.Range(wksLineA1.Cells(4, 11), wksLineA1.Cells(4, 13)).ColumnWidth = 21.71    'columwitd
    NEXTROW = 4
    
    Mark = False
    With flxSum
       For i = 1 To .Cols - 1
            If Not (i = 14) Then
            If (i = 14) Then
                    ROWOFFLX = i - 1
               Else
                    ROWOFFLX = i
                End If
                For C = 1 To .Rows - 1
                    wksLineA1.Cells(C + NEXTROW, ROWOFFLX) = .TextMatrix(C, i)
                 Next C
             End If
        Next i
        TotalRow = C
    End With
    Set rRng = wksLineA1.Range(wksLineA1.Cells(5, 1), wksLineA1.Cells(TotalRow + 3, 13))
    rRng.Font.Name = "TAHOMA"
    rRng.Borders.LineStyle = xlNone
    rRng.Borders.LineStyle = xlContinuous
    rRng.Font.Size = 10
    rRng.HorizontalAlignment = xlVAlignCenter
    
    XLapp.DisplayAlerts = False
    wbkMVR.SaveAs filename:=sSaveAsFileName, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
'    wbkMVR.Close SaveChanges:=False
'    XLapp.Quit

    Set XLapp = Nothing
    Set wbkMVR = Nothing
    Set wksLineA1 = Nothing
    MsgBox "Finish Saving to Excel!", vbInformation + vbOKOnly, "SUM"
    
End Sub


Private Sub cmdExcel_Click(Index As Integer)
     Call cmdExcel2_Click
End Sub

Private Sub cmdExit_Click()
        FRMMAIN.MenuSupplier.Visible = True
'        FRMMAIN.MenuCompany.Visible = True
'        FRMMAIN.MenuCustomer.Visible = True
'        FRMMAIN.MenuBanking.Visible = True
        FRMMAIN.MenuReports.Visible = True
        FRMMAIN.MenuReports.Visible = True
        FRMMAIN.MenuList.Visible = True
        FRMMAIN.MenuEntry.Visible = True
        FRMMAIN.MenuTools.Visible = True
        FRMMAIN.MenuQuit.Visible = True

    Unload Me
    FRMMAIN.tbMain.Visible = True
End Sub
Private Sub cmdfilter_Click()
    SSTab1.Visible = True
End Sub

Private Sub cmdFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then ComboSrch = True
    If KeyCode = vbKeyBack Then ComboSrch = True
End Sub

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    flxSum.HighLight = flexHighlightAlways
    
    With flxSum
        For i = fRow To .Rows - 1
            For j = fCol To TCol - 1
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .COL = 1: .Row = i
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
    Dto.Visible = True
    cmdRefresh.Visible = True
    Dto.Value = Dfrom.Value
End Sub

Private Sub cmdPrev_Click()
    cmdNext.Tag = "1"
    Dto.Value = Dfrom.Value
    cmdNext.Left = cmdPrev.Left
    cmdNext.Visible = True
    cmdPrev.Visible = False
    Dto.Visible = False
    cmdRefresh.Visible = False
    ShowDate = 1
    Call RefreshRecord(1)
    
End Sub

Private Sub cmdPrint_Click()
    'Dim Qry As String
    Dim dcov As String
    
    If MySource = "INST-ACCT" Then
    'dCov = "Date Covered: " & Format(Dfrom.Value, "Mmm. dd, yyyy") & IIf(Val(cmdNext.Tag) = 2, " to " & Format(Dto.Value, "Mmm. dd, yyyy"), "")
    If Val(cmdNext.Tag) = 1 Then
    
    Qry = "SELECT " & _
            "DATE_FORMAT(SalesReport.INVDATE,'%c/%d/%Y'), " & _
            "SalesReport.ACCOUNTNO, " & _
            "SalesReport.CUSTOMER, " & _
            "SalesReport.ADDRESS, " & _
            "SalesReport.UNITDESC, " & _
            "SalesReport.COD, " & _
            "SalesReport.LCP, " & _
            "SalesReport.INVAMOUNT, " & _
            "SalesReport.INVTYPE, " & _
            "SalesReport.TERMS, " & _
            "SalesReport.SALESAGENT, " & _
            "SalesReport.COMMISSION,  " & _
            "SalesReport.DATERELEASED, " & _
            "SalesReport.PCVNO, " & _
            "SalesReport.REBCOMM, " & _
            "unitordered.MGRREMARKS " & _
          "From  " & _
            "SalesReport " & _
            "INNER JOIN unitordered ON (SalesReport.ACCOUNTNO = unitordered.ACCOUNTNO) " & _
            "Where " & _
            "SalesReport.INVDATE = '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' " & _
            "AND salesreport.INVDATE >= '2015-01-01' " & _
            "AND  salesreport.INVDATE <= '2015-04-01' AND  unitordered.MGRREMARKS LIKE '%INSTITUTIONAL%' " & _
            "Order By  salesreport.INVDATE,  salesreport.INVTYPE,  SalesReport.ACCOUNTNO "
            dcov = "Date Covered: " & Format(Dfrom.Value, "Mmm. dd, yyyy")
        Else
            Qry = "SELECT " & _
            "DATE_FORMAT(SalesReport.INVDATE,'%c/%d/%Y'),  " & _
            "SalesReport.ACCOUNTNO,  " & _
            "SalesReport.CUSTOMER, " & _
            "SalesReport.ADDRESS, " & _
            "SalesReport.UNITDESC, " & _
            "SalesReport.COD,  " & _
            "SalesReport.LCP,  " & _
            "SalesReport.INVAMOUNT, " & _
            "SalesReport.INVTYPE, " & _
            "SalesReport.TERMS, " & _
            "SalesReport.SALESAGENT, " & _
            "SalesReport.COMMISSION, " & _
            "SalesReport.DATERELEASED, " & _
            "SalesReport.PCVNO, " & _
            "SalesReport.REBCOMM, " & _
            "unitordered.MGRREMARKS " & _
            "From  SalesReport  INNER JOIN unitordered ON (SalesReport.ACCOUNTNO = unitordered.ACCOUNTNO) " & _
            "Where SalesReport.INVDATE >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' " & _
            "AND salesreport.INVDATE >= '2015-01-01' AND  salesreport.INVDATE <= '2015-04-01' " & _
            "AND SalesReport.INVDATE <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' " & _
            "AND  unitordered.MGRREMARKS LIKE '%INSTITUTIONAL%' " & _
            "Order By  salesreport.INVDATE,  salesreport.INVTYPE,  SalesReport.ACCOUNTNO "
            dcov = "Date Covered: " & Format(Dfrom.Value, "mm/dd/yyyy") & " to " & Format(Dto.Value, "mm/dd/yyyy")
        End If
        frmReport.QryPrn = Qry
        frmReport.ParamValue1 = dcov
        frmReport.MySource = "INST-ACCT"
        frmReport.Show vbModal
    ElseIf MySource = "BRANCH" Then
'        Qry = "SELECT branch.BRANCHNAME,branch.MANAGER,branch.ADDRESS,branch.CONTACTNOS,branch.MCDEALERCODE,branch.CVDEALERCODE,branch.VATREGTIN " & _
'        "From branch"
        Qry = "SELECT branch.BRANCHNAME,branch.BRANCHCODE,branch.CGROUP,branch.MANAGER,branch.ADDRESS,branch.DSNNAME,branch.CONTACTNOS FROM  branch " & _
        "Order By branch.BRANCHNAME"
        frmReport.QryPrn = Qry
        frmReport.MySource = "BRANCH"
        'frmReport.ParamValue1 = "LYR TAGUM"
        frmReport.Show vbModal
    ElseIf MySource = "CHART OF ACCOUNTS" Then
        Qry = "SELECT chartaccts.ACCTNO,chartaccts.ACCTTITLE,chartaccts.ACCTTYPE,chartaccts.GRPACCT,chartaccts.ACCTRPT,chartaccts.NBAL," & _
        "chartaccts.REMARKS, chartaccts.`STATUS`,chartaccts.ACCTID " & _
        "From  chartaccts"
        frmReport.QryPrn = Qry
        frmReport.MySource = "CHART OF ACCOUNTS"
        frmReport.Show vbModal
    ElseIf MySource = "PRODUCTS" Then
'        Qry = "SELECT products.BRAND, products.MODEL, products.DESCRIPTION, products.CLASSIFICATION, products.PRODTYPE, products.PRODINIT, products.ACCTITLE, products.PRODID " & _
'        "From products Order By products.BRAND, products.MODEL"
'      Qry = "SELECT   products.BRAND,  products.MODEL,  products.DESCRIPTION,  products.CLASSIFICATION,  products.PRODTYPE, " & _
'        "products.PRODINIT,  products.ACCTITLE,  supplier.SUPPLIERNAME,  products.PRODID " & _
'        "From  products " & _
'        "INNER JOIN supplier ON (products.SUPPLIERID = supplier.SUPPLIERID)" & _
'        "Order By  products.BRAND,  products.MODEL "
        Qry = "SELECT PRODUCTS.BRAND,PRODUCTS.MODEL,products.MODELGROUP,PRODUCTS.DESCRIPTION,PRODUCTS.CLASSIFICATION,PRODUCTS.PRODTYPE,PRODUCTS.PRODINIT, " & _
            "PRODUCTS.ACCTITLE,supplier.SUPPLIERNAME " & _
            "From  PRODUCTS  INNER JOIN supplier ON (PRODUCTS.SUPPLIERID = supplier.SUPPLIERID) ORDER BY products.MODELGROUP DESC"
    
            frmReport.QryPrn = Qry
            frmReport.MySource = "PRODUCTS"
            frmReport.Show vbModal
            
    ElseIf MySource = "SPCRR" Then
            Qry = "CALL sp_spcrr_summary( '" & Format(Dfrom.Value, "YYYY-MM-DD") & "', '" & Format(Dto.Value, "YYYY-MM-DD") & "', '" & cboStat.Text & "');"
            frmReport.QryPrn = Qry
            frmReport.MySource = "SPCRR"
            frmReport.Show vbModal
            
    ElseIf MySource = "FA" Then
            Qry = "SELECT a.RFPDATE,a.RFPNO, b.SOLDTO,COUNT(b.RFPID) as ITEMS," & _
                  "a.PREPBY,a.REVBY,a.APPBY,a.`STATUS`, a.POTYPE,SUM(b.POAMNT) as NETPAY " & _
                  "FROM rfp as a INNER JOIN rfpdetails as b ON a.RFID = b.RFPID WHERE a.RFPDATE >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' " & _
                  "AND a.RFPDATE <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' AND a.RTYPE = 'GSSD' AND a.POTYPE = 'FA' GROUP BY a.RFPNO ORDER BY RFPDATE ASC;"
            frmReport.QryPrn = Qry
            frmReport.MySource = "RFPGSSDSUM"
            frmReport.Show vbModal
            
    ElseIf MySource = "MP" Then
            Qry = "SELECT a.RFPDATE,a.RFPNO, b.SOLDTO,COUNT(b.RFPID) as ITEMS," & _
                  "a.PREPBY,a.REVBY,a.APPBY,a.`STATUS`, a.POTYPE,SUM(b.POAMNT) as NETPAY " & _
                  "FROM rfp as a INNER JOIN rfpdetails as b ON a.RFID = b.RFPID WHERE a.RFPDATE >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' " & _
                  "AND a.RFPDATE <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' AND a.RTYPE = 'GSSD' AND a.POTYPE = 'MP' GROUP BY a.RFPNO ORDER BY RFPDATE ASC;"
            frmReport.QryPrn = Qry
            frmReport.MySource = "RFPGSSDSUM"
            frmReport.Show vbModal
            
    ElseIf MySource = "RFP" Then
            Qry = "CALL `sp_showprintpreviewSumRFP`('" & Format(Dfrom.Value, "YYYY-MM-DD") & "', '" & Format(Dto.Value, "YYYY-MM-DD") & "');"
            frmReport.ParamValue1 = Dfrom.Value
            frmReport.ParamValue2 = Dto.Value
            frmReport.QryPrn = Qry
            frmReport.MySource = "APPLIANCERFP"
            frmReport.Show vbModal
            
    ElseIf MySource = "MOTORCYCLERFP" Then
            Qry = "CALL `sp_showprintpreviewSumMC`('" & Format(Dfrom.Value, "YYYY-MM-DD") & "', '" & Format(Dto.Value, "YYYY-MM-DD") & "');"
            frmReport.ParamValue1 = Dfrom.Value
            frmReport.ParamValue2 = Dto.Value
            frmReport.QryPrn = Qry
            frmReport.MySource = "MOTORCYCLERFP"
            frmReport.Show vbModal
            
    ElseIf MySource = "AGENT" Then
    'dCov = "Date Covered: " & Format(Dfrom.Value, "Mmm. dd, yyyy") & IIf(Val(cmdNext.Tag) = 2, " to " & Format(Dto.Value, "Mmm. dd, yyyy"), "")
        If Val(cmdNext.Tag) = 1 Then
            Qry = "SELECT " & _
                "CONCAT(agents.LASTNAME, ', ', agents.FIRSTNAME) AS AGENT, " & _
                "agentorders.RANK,DATE_FORMAT(unitordered.DATECLOSED,'%c/%d/%Y'),unitordered.ACCOUNTNO,CONCAT(customer.LASTNAME, ', ', customer.FIRSTNAME) AS customername,CONCAT(customer.ADDRESS, ' ', customer.TOWN) AS custAddress,concat(productline.BRAND, ' ', productline.DESCRIPTION, ' ', unitordered.MODEL) AS UnitDesc,agingprev.REMARKS " & _
              "From " & _
                "agingprev INNER JOIN unitordered ON (agingprev.ORDERID = unitordered.ORDERID) " & _
                "INNER JOIN agentorders ON (agingprev.ORDERID = agentorders.ORDERID) " & _
                "INNER JOIN agents ON (agentorders.AGENTID = agents.AGENTID) " & _
                "INNER JOIN customer ON (unitordered.CUSTOMERID = customer.CUSTOMERID) " & _
                "INNER JOIN productline ON (unitordered.PRODID = productline.PRODID) " & _
              "Where " & _
                "(agentorders.RANK = 'USM' OR agentorders.RANK = 'USS') " & _
                "AND unitordered.DATECLOSED = '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' " & _
                "AND unitordered.DATECLOSED >= '2015-01-01' " & _
                "AND unitordered.DATECLOSED < '2015-02-01' " & _
                "AND agingprev.REMARKS = 'CLOSED' " & _
                "AND agingprev.`STATUS` = 'GOOD' " & _
              "Order By " & _
                "agentorders.RANK, " & _
                "AGENT"
                
                dcov = "Date Covered: " & Format(Dfrom.Value, "Mmm. dd, yyyy")
        Else
                Qry = "SELECT " & _
                    "CONCAT(agents.LASTNAME, ', ', agents.FIRSTNAME) AS AGENT, " & _
                    "agentorders.RANK,DATE_FORMAT(unitordered.DATECLOSED,'%c/%d/%Y'), " & _
                    "unitordered.ACCOUNTNO,CONCAT(customer.LASTNAME, ', ', customer.FIRSTNAME) AS customername, " & _
                    "CONCAT(customer.ADDRESS, ' ', customer.TOWN) AS custAddress, " & _
                    "concat(productline.BRAND, ' ', productline.DESCRIPTION, ' ', unitordered.MODEL) AS UnitDesc, " & _
                    "agingprev.REMARKS " & _
                  "From " & _
                    "agingprev INNER JOIN unitordered ON (agingprev.ORDERID = unitordered.ORDERID) " & _
                    "INNER JOIN agentorders ON (agingprev.ORDERID = agentorders.ORDERID) " & _
                    "INNER JOIN agents ON (agentorders.AGENTID = agents.AGENTID) " & _
                    "INNER JOIN customer ON (unitordered.CUSTOMERID = customer.CUSTOMERID) " & _
                    "INNER JOIN productline ON (unitordered.PRODID = productline.PRODID) " & _
                  "Where " & _
                    "(agentorders.RANK = 'USM' OR agentorders.RANK = 'USS') " & _
                    "AND unitordered.DATECLOSED >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' " & _
                    "AND unitordered.DATECLOSED >= '2015-01-01' " & _
                    "AND unitordered.DATECLOSED < '2015-02-01' " & _
                    "AND unitordered.DATECLOSED <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' " & _
                    "AND agingprev.REMARKS = 'CLOSED' " & _
                    "AND agingprev.`STATUS` = 'GOOD' " & _
                  "Order By " & _
                    "agentorders.RANK, " & _
                    "AGENT"
                dcov = "Date Covered: " & Format(Dfrom.Value, "mm/dd/yyyy") & " to " & Format(Dto.Value, "mm/dd/yyyy")

        End If
        

        frmReport.QryPrn = Qry
        frmReport.ParamValue1 = dcov
        frmReport.MySource = "AGENT"
        frmReport.Show vbModal

    ElseIf MySource = "SUPPLIER" Then
        Qry = "SELECT supplier_0.SUPPLIERID, supplier_0.SUPPLIERNAME, supplier_0.SUPPLIERCODE, supplier_0.ADDRESS, supplier_0.CONTACTNO," & _
        "supplier_0.FAXNO, supplier_0.CONTACTPERSON, supplier_0.CONPERPOSITION, supplier_0.MOBILENO, supplier_0.EMAIL, supplier_0.UPDATEDBY," & _
        "supplier_0.LASTUPDATE, supplier_0.TELNO, supplier_0.TINNUMBER, supplier_0.INPUTTAX, supplier_0.TERMS, supplier_0.BANK, supplier_0.REMARKS," & _
        "supplier_0.BANKACCT, supplier_0.WEB, supplier_0.STATUTORY, supplier_0.BILLTYPE, supplier_0.SUPTYPE  FROM lyrpurchase.supplier supplier_0"
'        Qry = "SELECT  supplier.SUPPLIERNAME,supplier.ADDRESS,supplier.CONTACTPERSON,supplier.TELNO,supplier.MOBILENO,supplier.FAXNO," & _
'        "supplier.EMAIL,supplier.SUPPLIERCODE,supplier.TINNUMBER,supplier.TERMS,supplier.BANK,supplier.REMARKS, supplier.BANKACCT,supplier.INPUTTAX, & _
            "supplier.WEB,supplier.SUPPLIERID FROM supplier"
        frmReport.QryPrn = Qry
        frmReport.MySource = "SUPPLIER"
        frmReport.Show vbModal
    ElseIf MySource = "BACK ORDER" Then
        Qry = "SELECT DATE_FORMAT(backorder.TRANSDATE,'%c/%d/%Y'), backorder.REFERENCE, spareparts.PARTNUMBER, spareparts.DESCRIPTION, backorder.BACKORDERQTY, backorder.BOBALANCE, backorder.BRANCHNAME, backorder.SUPPORDERREF " & _
              "From backorder INNER JOIN spareparts ON (backorder.SPAREPARTID = spareparts.SPAREPARTID) " & _
              "WHERE backorder.TRANSDATE >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' AND backorder.TRANSDATE <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' ORDER BY backorder.TRANSDATE DESC"
        frmReport.QryPrn = Qry
        frmReport.MySource = "BACK ORDER"
        frmReport.Show vbModal
        
    ElseIf MySource = "RFPOTHER" Then
        
        If cboPrepBy.Text = "" Then
        
            Qry = "SELECT a.RFPDATE,a.RFPNO, b.SOLDTO,s.SupplierName, a.TYPEXP, " & _
                  "a.Remarks,b.NETPAY,a.PREPBY,a.REVBY,a.NOTBY, a.`Status` " & _
                  "FROM rfp as a INNER JOIN rfpdetails as b ON a.RFID = b.RFPID " & _
                  "INNER JOIN supplier as s ON a.supplierid = s.supplierid " & _
                  "WHERE a.RFPDATE >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' AND a.RFPDATE <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' " & _
                  "AND a.RTYPE = 'OTHER' GROUP BY a.RFPNO ORDER BY a.RFPDATE DESC"
        
        Else
                 
            Qry = "SELECT a.RFPDATE,a.RFPNO, b.SOLDTO,s.SupplierName, a.TYPEXP, " & _
                  "a.Remarks,b.NETPAY,a.PREPBY,a.REVBY,a.NOTBY, a.`Status` " & _
                  "FROM rfp as a INNER JOIN rfpdetails as b ON a.RFID = b.RFPID " & _
                  "INNER JOIN supplier as s ON a.supplierid = s.supplierid " & _
                  "WHERE a.RFPDATE >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' AND a.RFPDATE <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' " & _
                  "AND a.PREPBY = '" & cboPrepBy.Text & "' AND a.RTYPE = 'OTHER' GROUP BY a.RFPNO ORDER BY a.RFPDATE DESC"
        End If
        
        frmReport.QryPrn = Qry
        'InputBox Qry, Qry, Qry
        frmReport.MySource = "RFPOTHERSUM"
        frmReport.Show vbModal
    
    Else
    
        MsgBox "Not yet available", vbInformation, "No print out"
        Screen.MousePointer = vbDefault
        Exit Sub
        
    End If
End Sub

Private Sub cmdprosfil_Click()
    Dim Qry As String, i As Integer
    Dim rsSummary As New ADODB.Recordset

    If cbofilter.Text = "" Then
        MsgBox "Please select suppliername", vbCritical + vbOKOnly, "System Message"
        Exit Sub
    End If
    
    If MySource = "RFP" Then 'OK
        Qry = "CALL `sp_filterRfpAppSum`( '" & Format(Dfrom.Value, "YYYY-MM-DD") & "', '" & Format(Dto.Value, "YYYY-MM-DD") & "', '" & cbofilter.Text & "');"
    ElseIf MySource = "MOTORCYCLERFP" Then
        Qry = "CALL `sp_filterRfpMcSum`( '" & Format(Dfrom.Value, "YYYY-MM-DD") & "', '" & Format(Dto.Value, "YYYY-MM-DD") & "', '" & cbofilter.Text & "');"
    ElseIf MySource = "POBILLING" Then
        Qry = "CALL `sp_filterBillingSummary`('" & Format(Dfrom.Value, "YYYY-MM-DD") & "', '" & Format(Dto.Value, "YYYY-MM-DD") & "', '" & cbofilter.Text & "');"
    ElseIf MySource = "POBILLINGMC" Then
        Qry = "CALL `sp_filterBillingMcSummary`('" & Format(Dfrom.Value, "YYYY-MM-DD") & "', '" & Format(Dto.Value, "YYYY-MM-DD") & "', '" & cbofilter.Text & "');"
    ElseIf MySource = "PURCHASE ORDER" Then
        Qry = "CALL `sp_filterPOsum`( '" & Format(Dfrom.Value, "YYYY-MM-DD") & "', '" & Format(Dto.Value, "YYYY-MM-DD") & "', '" & cbofilter.Text & "');"
    ElseIf MySource = "BRANCH ALLOCATION" Then
'        Qry = "CALL `sp_filterPOsum`( '" & Format(DFROM.Value, "YYYY-MM-DD") & "', '" & Format(DTO.Value, "YYYY-MM-DD") & "', '" & cbofilter.Text & "');"
    ElseIf MySource = "SPCRR" Then
        If cboStat.Text = "ALL" Then
            Qry = "SELECT DATECREATED, SPCRRNO, RRNO, RRDATE, BRANCH, SUPPLIER,PREPAREDBY, TERMS, INVSTATUS, INVOICETTLPAY, SPCRRID FROM " & _
                  "spcrrmain WHERE DATECREATED >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' AND DATECREATED <= '" & Format(Dto.Value, "YYYY-MM-DD") & "';"
        Else
             Qry = "SELECT DATECREATED, SPCRRNO, RRNO, RRDATE, BRANCH, SUPPLIER,PREPAREDBY, TERMS, INVSTATUS, INVOICETTLPAY, SPCRRID FROM " & _
                  "spcrrmain WHERE DATECREATED >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' AND DATECREATED <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' AND INVSTATUS = '" & cboStat.Text & "';"
        End If
    End If
    
    rsSummary.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxSum.DataSource = rsSummary   'With records
    ActRow = 1
    flxSum.Visible = True
   Call RXResizer1_ResizeComplete
    If MySource = "RFP" Or MySource = "MOTORCYCLERFP" Then
        For i = 1 To flxSum.Rows - 1
            If flxSum.TextMatrix(i, 5) = flxSum.TextMatrix(i, 6) Then
                With flxSum
                    .COL = 0:
                    .Row = i:
                    .ColSel = 13:
                    .RowSel = i:
                    .FillStyle = flexFillRepeat:
                    .CellBackColor = RGB(255, 255, 155)
                    .COL = 2
                End With
            End If
        Next i
      TotalPO = 0
            For i = 1 To flxSum.Rows - 1
                TOTALAMNT = Format(Val(TOTALAMNT) + Val(Replace(flxSum.TextMatrix(i, 9), ",", "")), "#,##0.00")
            Next i
            txtTTLamount.Caption = ""
            txtTTLamount.Caption = "NET OVER ALL TOTAL" & ": " & Format(TOTALAMNT, "#,##0.00")
    SSTab1.Visible = False
    
    ElseIf MySource = "POBILLING" Or MySource = "POBILLINGMC" Then
        For i = 1 To flxSum.Rows - 1
            TOTALAMNT = Format(Val(TOTALAMNT) + Val(Replace(flxSum.TextMatrix(i, 11), ",", "")), "#,##0.00")
        Next i
        txtTTLamount.Caption = ""
        txtTTLamount.Caption = "OVER ALL TOTAL" & ": " & Format(TOTALAMNT, "#,##0.00")
    SSTab1.Visible = False
    
    ElseIf MySource = "BRANCH ALLOCATION" Then
      For i = 1 To flxSum.Rows - 1
            TOTALAMNT = Format(Val(TOTALAMNT) + Val(Replace(flxSum.TextMatrix(i, 9), ",", "")), "#,##0.00")
        Next i
        txtTTLamount.Caption = ""
        txtTTLamount.Caption = "OVER ALL TOTAL" & ": " & Format(TOTALAMNT, "#,##0.00")
    SSTab1.Visible = False
    
    ElseIf MySource = "PURCHASE ORDER" Or MySource = "PR" Or MySource = "PRMC" Then
      For i = 1 To flxSum.Rows - 1
            TOTALAMNT = Format(Val(TOTALAMNT) + Val(Replace(flxSum.TextMatrix(i, 10), ",", "")), "#,##0.00")
        Next i
        txtTTLamount.Caption = ""
        txtTTLamount.Caption = "OVER ALL TOTAL" & ": " & Format(TOTALAMNT, "#,##0.00")
    SSTab1.Visible = False
    
    Dfrom.Value = DTPicker1.Value
    Dto.Value = DTPicker2.Value
    Call cmdvisble_Click
  End If
End Sub

Private Sub cmdRefresh_Click()
    Call RefreshRecord(2)
End Sub
Private Sub cmdShow_Click()

    With flxSum
        If .Rows = 2 And .TextMatrix(1, 1) = "" Then
            MsgBox "There is no record to view!", vbInformation + vbOKOnly, "Confirm"
            Exit Sub
        End If
        If MySource = "PRODUCTS" Then
            Call Product.ShowRecord(Val(flxSum.TextMatrix(ActRow, 10)))
            Unload Me
            Product.Show vbModal
        ElseIf MySource = "BRANCH" Then
            Call Branch.ShowRecord(Val(flxSum.TextMatrix(ActRow, 18)))
            Branch.Show vbModal
        ElseIf MySource = "SUPPLIER" Then
            Call frmSupplier.ShowRecord(Val(flxSum.TextMatrix(ActRow, 1)))
            frmSupplier.Show
        ElseIf MySource = "CHART OF ACCOUNTS" Then
            Call ChartOfAcct.ShowRecord(Val(flxSum.TextMatrix(ActRow, 9)))
            ChartOfAcct.Show 1
        ElseIf MySource = "PURCHASE ORDER" Then
            Call frmPO.ShowRecord(Val(flxSum.TextMatrix(ActRow, 12)))
            frmPO.Show
            Unload Me
        '------------------------------1/09/2018---------------------
        ElseIf MySource = "BRANCH ALLOCATION" Then
            Call frmPO.ShowRecord(Val(flxSum.TextMatrix(ActRow, 9)))
            frmPO.Show
            Unload Me
        ElseIf MySource = "PRCM" Then
            Call frmPReturn.ShowRecord(Val(flxSum.TextMatrix(ActRow, 11)))
            frmPReturn.Show
            Unload Me
         ElseIf MySource = "PRDM" Then
            Call frmDebitMemo.ShowRecord(Val(flxSum.TextMatrix(ActRow, 13)))
            frmDebitMemo.Show
            Unload Me
        '------------------------------ARIAS ------------------------
        ElseIf MySource = "PR" Then
            frmPR.SecID = "1"
            frmPR.HIS = False
            Call frmPR.ShowRecord(Val(flxSum.TextMatrix(ActRow, 12)))
            frmPR.Show
            Unload Me
        ElseIf MySource = "PRMC" Then
            FrmPR2.SecID = "1"
            FrmPR2.HIS = False
            Call FrmPR2.ShowRecord(Val(flxSum.TextMatrix(ActRow, 12)))
            FrmPR2.Show
            Unload Me
        ElseIf MySource = "PURCHASE ORDER2" Then
            Call FrmPO2.ShowRecord(Val(flxSum.TextMatrix(ActRow, 12)))
            FrmPO2.Show
            Unload Me
        ElseIf MySource = "POBILLING" Then
            Call frmBillingApp.ShowRecPO(flxSum.TextMatrix(ActRow, 13))
            frmBillingApp.Show
            Unload Me
        ElseIf MySource = "POBILLINGMC" Then
            Call frmBillingMotor.ShowRecPO(flxSum.TextMatrix(ActRow, 13))
            frmBillingMotor.Show
            Unload Me
        ElseIf MySource = "CHECKPO" Then
            If .TextMatrix(ActRow, 3) = "" And .TextMatrix(ActRow, 5) = "" Then
                MsgBox "This P.O is not available for viewing", vbOKOnly + vbExclamation, "Access Denied"
                Exit Sub
            End If
            Call frmBilling.ShowRecPO(flxSum.TextMatrix(ActRow, 5))
            frmBilling.Show
            Unload Me
        ElseIf MySource = "RFP" Or MySource = "MOTORCYCLERFP" Then
            Call frmRFP.ShowRecord(flxSum.TextMatrix(ActRow, 14))
            frmRFP.Show
            Unload Me
        ElseIf MySource = "CBD" Then
            Call frmrfpCBD.ShowRecord(flxSum.TextMatrix(ActRow, 10))
            frmrfpCBD.Show
            Unload Me
         ElseIf MySource = "MCCBD" Then
            Call frmrfpCBD.ShowRecord(flxSum.TextMatrix(ActRow, 10))
            frmrfpCBD.Show
            Unload Me
        ElseIf MySource = "Price List" Then
            Call PrcPriceList.ShowRecord(.TextMatrix(ActRow, 9))
            Unload Me
        ElseIf MySource = "Employees" Then
            Call Employees.ShowRecord(Val(flxSum.TextMatrix(ActRow, 1)))
            Unload Me
        ElseIf MySource = "SPCRR" Then
            Call frmSPCRR.ShowRecord(flxSum.TextMatrix(ActRow, 2))
            Unload Me
        ElseIf MySource = "FA" Or MySource = "MP" Then
            
            frmRFPGSSD.RFPID = flxSum.TextMatrix(ActRow, 12)
            frmRFPGSSD.POType = flxSum.TextMatrix(ActRow, 9)
            frmRFPGSSD.tmrGetRecord.Enabled = True
            frmRFPGSSD.POType = MySource
            
            DoEvents
            Call Load(frm_OnProcess)
            frm_OnProcess.Show vbModal
            Unload Me
        
        ElseIf MySource = "RFPOTHER" Then
        
            frmRFPOther.RFPID = flxSum.TextMatrix(ActRow, 12)
            frmRFPOther.tmrGetRecord.Enabled = True
            
            DoEvents
            frm_OnProcess.Show vbModal
            Unload Me
            
        End If
        
    End With
 
End Sub

Private Sub Command1_Click()
Call checkPR
    
End Sub

 

Private Sub cmdvisble_Click()
    cbofilter.Text = ""
    SSTab1.Visible = False
End Sub

Private Sub Dfrom_CloseUp()
    DateStart = Dfrom.Value
    If cmdNext.Tag = "1" Then
        Call RefreshRecord(1)
    End If
End Sub


Private Sub Dto_CloseUp()
    DateEnd = Dto.Value
    If Dto.Value < Dfrom.Value Then
        MsgBox "Please use a date not less than " & Dfrom.Value, vbExclamation + vbOKOnly, "Validation Error"
        Dto.Value = Dfrom.Value
    End If
End Sub

Private Sub flxBranch_EnterCell()
'    With flxBranch
'        FromFlx = True
'        txtShipTo.Text = .TextMatrix(.Row, 1)
'        lblTypeShipTo.Caption = .TextMatrix(.Row, 2)
'        lblBranchShipToCode.Caption = .TextMatrix(.Row, 3)
'        lbldsnCon.Caption = .TextMatrix(.Row, 4)
'
'    End With
End Sub

Private Sub flxBranch_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyEscape Then
'        flxBranch.Visible = False
'        txtShipTo.SetFocus
'    End If
End Sub

Private Sub flxBranch_KeyPress(KeyAscii As Integer)
'    flxBranch.Visible = False
'    cmdPONo.SetFocus
End Sub
Private Sub flxBranch_LostFocus()
      flxBranch.Visible = False
End Sub

Private Sub flxSum_Click()
    ActRow = flxSum.Row
    ActCol = flxSum.COL
    flxSum.HighLight = flexHighlightWithFocus
    
    If ActCol = 0 Or ActRow = 0 Then Exit Sub
    If ViewOnly = True Then
        Exit Sub
    End If
    
    If MySource = "POBILLINGMC" Then
    If ActCol = 2 Then
               With txt
                  .Tag = flxSum.TextMatrix(ActRow, ActCol)
                   .Top = flxSum.CellTop + flxSum.Top
                   .Left = flxSum.CellLeft + flxSum.Left
                   .Width = flxSum.CellWidth
                   .Text = flxSum.Text
                   .Visible = True
                   .SelStart = 0
                   .SelLength = Len(.Text)
                   .ZOrder
                   .SetFocus
               End With
           End If
         If ActCol = 6 Then
               With txtdr
                  .Tag = flxSum.TextMatrix(ActRow, ActCol)
                   .Top = flxSum.CellTop + flxSum.Top
                   .Left = flxSum.CellLeft + flxSum.Left
                   .Width = flxSum.CellWidth
                   .Text = flxSum.Text
                   .Visible = True
                   .SelStart = 0
                   .SelLength = Len(.Text)
                   .ZOrder
                   .SetFocus
               End With
           End If
End If
End Sub
   
Private Sub flxSum_EnterCell()
    ActRow = flxSum.Row
End Sub

Private Sub flxSum_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If MySource = "POBILLINGMC" Or MySource = "POBILLING" Then
        If Button = vbRightButton Then
            ActRow = flxSum.Row
            Call PopupMenu(mnuBtn)
        End If
'     ElseIf MySource = "RFP" Or MySource = "MOTORCYCLERFP" Then
'        If Button = vbRightButton Then
'            ActRow = flxSum.Row
''            Me.name_of_the_menu.Enabled = False
'            Call PopupMenu(menuCM)
'        End If
     ElseIf MySource = "FA" Or MySource = "MP" Or MySource = "SPCRR" Or MySource = "CBD" Or MySource = "MCCBD" Or MySource = "RFP" Or MySource = "MOTORCYCLERFP" Or MySource = "RFPOTHER" Then
        If Button = vbRightButton Then
            ActRow = flxSum.Row
            Call PopupMenu(MenuCancel)
        End If
     End If
End Sub

Private Sub Form_Load()
    Dim rsSupp As New ADODB.Recordset
    cbofilter.clear
    FRMMAIN.tbMain.Visible = False
    Me.Caption = MySource & " SUMMARY"
    Timer1.Enabled = True
    
    lblHeading(0).Caption = MySource & " Summary"
    Me.Caption = MySource & " Summary"
    flxSum.HighLight = flexHighlightNever
    
'    If MySource = "PURCHASE ORDER" Then
        With rsSupp
        Qry = "SELECT supplier.SUPPLIERNAME, supplier.SUPPLIERID From supplier Order By supplier.SUPPLIERNAME"
         If rsSummary.State = adStateOpen Then rsSummary.Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cbofilter.AddItem !SUPPLIERNAME
                .MoveNext
            Loop
        End If
        .Close
        End With
        
        
     If MySource = "RFPOTHER" Then
     
        With rsSupp
        Qry = "SELECT DISTINCT(ifnull(prepby,'')) as PrepBy FROM rfp order by PrepBy;"
         If rsSummary.State = adStateOpen Then rsSummary.Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboPrepBy.AddItem !PrepBy
                .MoveNext
            Loop
        End If
        .Close
        End With
        
        lblPrepBy.Visible = True
        cboPrepBy.Visible = True
        Call modObjTools.changeListIndex(cboPrepBy, UserFullName)
        
     Else
        lblPrepBy.Visible = False
        cboPrepBy.Visible = False
     
     End If
'    ElseIf MySource = "RFP" Then
'         With rsSupp
'        Qry = "SELECT supplier.SUPPLIERNAME, supplier.SUPPLIERID From supplier Order By supplier.SUPPLIERNAME"
'         If rsSummary.State = adStateOpen Then rsSummary.Close
'        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
'        If .RecordCount > 0 Then
'            .MoveFirst
'            Do While Not .EOF
'                cbofilter.AddItem !SUPPLIERNAME
'                .MoveNext
'            Loop
'        End If
'        .Close
'        End With
'     ElseIf MySource = "POBILLING" Then
'         With rsSupp
'        Qry = "SELECT supplier.SUPPLIERNAME, supplier.SUPPLIERID From supplier Order By supplier.SUPPLIERNAME"
'         If rsSummary.State = adStateOpen Then rsSummary.Close
'        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
'        If .RecordCount > 0 Then
'            .MoveFirst
'            Do While Not .EOF
'                cbofilter.AddItem !SUPPLIERNAME
'                .MoveNext
'            Loop
'        End If
'        .Close
'        End With
'    End If
End Sub
Private Sub checkPR()
    Dim i As Integer, Qry As String, prNO As Variant
     With flxSum
        For i = 1 To .Rows - 1
          Qry = "SELECT prsupplier.PONO FROM prsupplier WHERE prsupplier.PRNUMBER = '" & .TextMatrix(i, 2) & "' AND   prsupplier.`TYPE` IS NULL"
           'InputBox Qry, , Qry
            prNO = GetFieldValue(Qry)
            If prNO = "" Then prNO = 0
            If prNO > 0 Then
                .COL = 1:
                .Row = i:
                .ColSel = 11:
                .RowSel = i:
                .FillStyle = flexFillRepeat:
                .CellBackColor = RGB(192, 192, 192)
                .COL = 2
            End If
            Next i
        End With
End Sub
Private Sub checkPO()
    Dim i As Integer, Qry As String, prNO As Variant
     With flxSum
        For i = 1 To .Rows - 1
          Qry = "SELECT posupplier.PRNO FROM posupplier WHERE posupplier.PONUMBER = '" & .TextMatrix(i, 2) & "' AND  posupplier.`TYPE` IS NOT NULL"
           'InputBox Qry, , Qry
            prNO = GetFieldValue(Qry)
            If prNO = "" Then prNO = 0
            If prNO > 0 Then
                .COL = 1:
                .Row = i:
                .ColSel = 11:
                .RowSel = i:
                .FillStyle = flexFillRepeat:
                .CellBackColor = RGB(192, 192, 192)
                .COL = 2
            End If
            Next i
        End With
End Sub
Private Sub checkPODRinvoice()
    Dim i As Integer, Qry As String, prNO As Variant, ttl1 As Currency, ttl2 As Currency
    
    With flxSum
        For i = 1 To .Rows - 1
            Qry = "SELECT rfppoinvoice.RFID FROM rfppoinvoice WHERE rfppoinvoice.RFID ='" & .TextMatrix(i, 10) & "'"

           'InputBox Qry, , Qry
            PODRinvoice = GetFieldValue(Qry)
            If PODRinvoice = "" Then PODRinvoice = 0
            If PODRinvoice > 0 Then
                Qry = "SELECT sum(rfppoinvoice.POAMNT) as ttl FROM rfppoinvoice WHERE rfppoinvoice.RFID ='" & .TextMatrix(i, 10) & "'"
                ttl1 = Val(GetFieldValue(Qry))
                Qry = "SELECT sum(rfpdetails.POAMNT) as ttl2 FROM rfpdetails WHERE rfpdetails.RFPID = '" & .TextMatrix(i, 10) & "'"
                ttl2 = GetFieldValue(Qry)
                
               If ttl2 = ttl1 Then
                    .COL = 1:
                    .Row = i:
                    .ColSel = 10:
                    .RowSel = i:
                    .FillStyle = flexFillRepeat:
                    .CellBackColor = RGB(255, 251, 159)
                    .COL = 2
               Else
                    .COL = 1:
                    .Row = i:
                    .ColSel = 10:
                    .RowSel = i:
                    .FillStyle = flexFillRepeat:
                    .CellBackColor = &HC0C0C0
                    .COL = 2
               End If
               
            End If
            
            If .TextMatrix(i, 4) = "CANCELLED" Then
                .COL = 0:
                .Row = i:
                .ColSel = 11:
                .RowSel = i:
                .FillStyle = flexFillRepeat:
                .CellBackColor = &H8080FF
                .COL = 2
            End If
                
        Next i
    End With
 
    
End Sub
Private Sub checkPRMC()
    Dim i As Integer, Qry As String, prNO As Variant
     With flxSum
        For i = 1 To .Rows - 1
           Qry = "SELECT prsupplier.PONO FROM prsupplier WHERE prsupplier.PRNUMBER = '" & .TextMatrix(i, 2) & "' AND   prsupplier.`TYPE` IS NOT NULL"
           'InputBox Qry, , Qry
            prNO = GetFieldValue(Qry)
            If prNO = "" Then prNO = 0
            If prNO > 0 Then
                .COL = 1:
                .Row = i:
                .ColSel = 11:
                .RowSel = i:
                .FillStyle = flexFillRepeat:
                .CellBackColor = RGB(192, 192, 192)
                .COL = 2
            End If
            Next i
        End With
End Sub


Private Sub mnuCancel_Click()
    If MySource = "SPCRR" Then
        Call CancelSPCRR
    ElseIf MySource = "RFP" Or MySource = "MOTORCYCLERFP" Then
        Call CancelRFPCOD
    ElseIf MySource = "CBD" Or MySource = "MCCBD" Then
        Call CancelRFPCBD
    ElseIf MySource = "FA" Or MySource = "MP" Then
        Call CancelRFPGSSD
    ElseIf MySource = "RFPOTHER" Then
        Call CancelRFPOTHER
    End If
End Sub

Private Sub CancelSPCRR()
On Error GoTo er
    Dim Ans As Integer, MyReason As String, Remarks As String, PrintCopy As Integer, SPRRNO As Integer
'  0   1        2       3           4       5          6          7           8       9          10            11
'"  |^DATE|^SPCRR NO.|^RR NO. |^RR DATE |<BRANCH |<SUPPLIER |<PREPARED BY |^TERMS |^STATUS |>TOTAL PAYABLE |SPCRRID"
    If flxSum.TextMatrix(ActRow, 9) = "PENDING" Then
        MsgBox "SPCRR no. was still pending!Cancellation not allowed!", vbExclamation + vbOKOnly, "Validation"
        Exit Sub
    End If
    If InStr(UCase$(flxSum.TextMatrix(ActRow, 9)), "CANCELLED") > 0 Then
        MsgBox "SPCRR no. " & flxSum.TextMatrix(ActRow, 2) & " was already cancelled!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    
    Ans = MsgBox("Do you want to Cancel SPCRR no.?" & vbCrLf & vbCrLf & _
        "SPCRR Date : " & flxSum.TextMatrix(ActRow, 1) & vbCrLf & "SPCRR NO: " & flxSum.TextMatrix(ActRow, 2), vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    MyReason = InputBox("Please indicate reason for Cancellation of SPCRR no.", "Reason for Cancel")
    If Replace(MyReason, " ", "") = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Packing no.!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    
    'Update SPCRRMAIN
    LYRcon.Execute "UPDATE SPCRRMAIN SET INVSTATUS = 'CANCELLED', INVNOTES = 'CANCELLED' ', ' '" & MyReason & "' WHERE SPCRRID = '" & flxSum.TextMatrix(ActRow, 11) & "';"
    
    'Delete to item cost table
    LYRcon.Execute "DELETE FROM lyrpurchase.itemcost where DRNO = '" & flxSum.TextMatrix(ActRow, 3) & "' AND UPDBY = 'SPCRR';"

    'Insert History
    LYRcon.Execute "INSERT INTO transhist(DATE, TRANSNAME, REFID, REMARKS, RESPONSIBLE, ACTION) VALUES " & _
                   "(Now(), 'SPCRR'," & flxSum.TextMatrix(ActRow, 11) & ",'" & MyReason & "', " & UserID & ", 'CANCEL') "
    
    flxSum.TextMatrix(ActRow, 9) = "CANCELLED"
    With flxSum     'Change row color
        .COL = 1: .Row = ActRow: .ColSel = 10: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
        .COL = 2
    End With
    

    
    MsgBox "SPCRR was number Cancelled!", vbInformation + vbOKOnly


Exit Sub
er:
    DoEvents
    MsgBox err.Description, vbCritical, "Error->CancelSPCRR()"
End Sub

Private Sub CancelRFPCOD()
On Error GoTo er
    Dim Ans As Integer, MyReason As String, Remarks As String, PrintCopy As Integer, SPRRNO As Integer

    If Val(flxSum.TextMatrix(ActRow, 16)) = 0 Then
        MsgBox "RFP no. was still unprinted! Cancellation not allowed!", vbExclamation + vbOKOnly, "Validation"
        Exit Sub
    End If
    If InStr(UCase$(flxSum.TextMatrix(ActRow, 15)), "CANCELLED") > 0 Then
        MsgBox "RFP no. " & flxSum.TextMatrix(ActRow, 2) & " was already cancelled!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    
    Ans = MsgBox("Do you want to Cancel RFP no.?" & vbCrLf & vbCrLf & _
        "RFP Date : " & flxSum.TextMatrix(ActRow, 1) & vbCrLf & "RFP NO: " & flxSum.TextMatrix(ActRow, 2), vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    MyReason = InputBox("Please indicate reason for Cancellation of RFP no.", "Reason for Cancel")
    If Replace(MyReason, " ", "") = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of RFP no.!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    
    'Update RFP
    LYRcon.Execute "UPDATE rfp SET `STATUS` = 'CANCELLED', REMARKS = 'CANCELLED' ', ' '" & MyReason & "' WHERE RFID = " & Val(flxSum.TextMatrix(ActRow, 14)) & ";"
               

    'Insert History
    LYRcon.Execute "INSERT INTO transhist(DATE, TRANSNAME, REFID, REMARKS, RESPONSIBLE, ACTION) VALUES " & _
                   "(Now(), 'RFPCOD'," & flxSum.TextMatrix(ActRow, 14) & ",'" & MyReason & "', " & UserID & ", 'CANCEL') "
    
    flxSum.TextMatrix(ActRow, 15) = "CANCELLED"
    With flxSum     'Change row color to red
        .COL = 1: .Row = ActRow: .ColSel = 15: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = &H8080FF
        .COL = 2
    End With
    

    MsgBox "RFP number was Cancelled!", vbInformation + vbOKOnly


Exit Sub
er:
    DoEvents
    MsgBox err.Description, vbCritical, "Error->CancelRFPCOD()"
End Sub

Private Sub CancelRFPCBD()
On Error GoTo er
    Dim Ans As Integer, MyReason As String, Remarks As String, PrintCopy As Integer, SPRRNO As Integer
    
    
    If Val(flxSum.TextMatrix(ActRow, 11)) = 0 Then
        MsgBox "RFP no. was still unprinted! Cancellation not allowed!", vbExclamation + vbOKOnly, "Validation"
        Exit Sub
    End If
    If InStr(UCase$(flxSum.TextMatrix(ActRow, 4)), "CANCELLED") > 0 Then
        MsgBox "RFP no. " & flxSum.TextMatrix(ActRow, 2) & " was already cancelled!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    
    Ans = MsgBox("Do you want to Cancel RFP no.?" & vbCrLf & vbCrLf & _
        "RFP Date : " & flxSum.TextMatrix(ActRow, 1) & vbCrLf & "RFP NO: " & flxSum.TextMatrix(ActRow, 2), vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    MyReason = InputBox("Please indicate reason for Cancellation of RFP no.", "Reason for Cancel")
    If Replace(MyReason, " ", "") = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of RFP no.!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    
    'Update RFP
    LYRcon.Execute "UPDATE rfp SET `STATUS` = 'CANCELLED', REMARKS = 'CANCELLED' ', ' '" & MyReason & "' WHERE RFID = " & Val(flxSum.TextMatrix(ActRow, 10)) & ";"
               

    'Insert History
    LYRcon.Execute "INSERT INTO transhist(DATE, TRANSNAME, REFID, REMARKS, RESPONSIBLE, ACTION) VALUES " & _
                   "(Now(), 'RFPCBD'," & flxSum.TextMatrix(ActRow, 10) & ",'" & MyReason & "', " & UserID & ", 'CANCEL') "
    
    flxSum.TextMatrix(ActRow, 4) = "CANCELLED"
    With flxSum     'Change row color to red
        .COL = 1: .Row = ActRow: .ColSel = 11: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = &H8080FF
        .COL = 2
    End With
    

    MsgBox "RFP number was Cancelled!", vbInformation + vbOKOnly


Exit Sub
er:
    DoEvents
    MsgBox err.Description, vbCritical, "Error->CancelRFPCBD()"
End Sub

Private Sub CancelRFPOTHER()
On Error GoTo er
    Dim Ans As Integer, MyReason As String, Remarks As String, PrintCopy As Integer, SPRRNO As Integer
    
'0    1      2        3        4          5           6            7           8        9              10           11
' |^DATE|^RFP NO.|<SOLDTO  |^ITEMS |<PREPARED BY |<CHECKED BY |<APPROVED BY |^STATUS |^STATUTORY |>TOTAL PAYABLE |^PRINT COPY "
    
    If Val(flxSum.TextMatrix(ActRow, 11)) = 0 Then
        MsgBox "RFP no. was still unprinted! Cancellation not allowed!", vbExclamation + vbOKOnly, "Validation"
        Exit Sub
    End If
    If InStr(UCase$(flxSum.TextMatrix(ActRow, 8)), "CANCELLED") > 0 Then
        MsgBox "RFP no. " & flxSum.TextMatrix(ActRow, 2) & " was already cancelled!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    
    Ans = MsgBox("Do you want to Cancel RFP no.?" & vbCrLf & vbCrLf & _
        "RFP Date : " & flxSum.TextMatrix(ActRow, 1) & vbCrLf & "RFP NO: " & flxSum.TextMatrix(ActRow, 2), vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    MyReason = InputBox("Please indicate reason for Cancellation of RFP no.", "Reason for Cancel")
    If Replace(MyReason, " ", "") = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of RFP no.!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    
    'Update RFP
    LYRcon.Execute "UPDATE rfp SET `STATUS` = 'CANCELLED', REMARKS = 'CANCELLED' ', ' '" & MyReason & "' WHERE RFID = " & Val(flxSum.TextMatrix(ActRow, 12)) & ";"
               

    'Insert History
    LYRcon.Execute "INSERT INTO transhist(DATE, TRANSNAME, REFID, REMARKS, RESPONSIBLE, ACTION) VALUES " & _
                   "(Now(), 'RFPOTHER'," & flxSum.TextMatrix(ActRow, 12) & ",'" & MyReason & "', " & UserID & ", 'CANCEL') "
    
    flxSum.TextMatrix(ActRow, 8) = "CANCELLED"
    With flxSum     'Change row color to red
        .COL = 1: .Row = ActRow: .ColSel = 11: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = &H8080FF
        .COL = 2
    End With
    

    MsgBox "RFP number was Cancelled!", vbInformation + vbOKOnly


Exit Sub
er:
    DoEvents
    MsgBox err.Description, vbCritical, "Error->CancelRFPOTHER()"
End Sub

Private Sub CancelRFPGSSD()
On Error GoTo er
    Dim Ans As Integer, MyReason As String, Remarks As String, PrintCopy As Integer, SPRRNO As Integer
    
'                  0    1      2        3          4            5           6            7           8        9              10           11          12
'  .FormatString = "  |^DATE|^RFP NO.|<SOLDTO  |^ITEMS |<PREPARED BY |<CHECKED BY |<APPROVED BY |^STATUS |^P.O. TYPE |>TOTAL PAYABLE |^PRINT COPY  |RFPID "
    
    If Val(flxSum.TextMatrix(ActRow, 11)) = 0 Then
        MsgBox "RFP no. was still unprinted! Cancellation not allowed!", vbExclamation + vbOKOnly, "Validation"
        Exit Sub
    End If
    If InStr(UCase$(flxSum.TextMatrix(ActRow, 8)), "CANCELLED") > 0 Then
        MsgBox "RFP no. " & flxSum.TextMatrix(ActRow, 2) & " was already cancelled!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    
    Ans = MsgBox("Do you want to Cancel RFP no.?" & vbCrLf & vbCrLf & _
        "RFP Date : " & flxSum.TextMatrix(ActRow, 1) & vbCrLf & "RFP NO: " & flxSum.TextMatrix(ActRow, 2), vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    MyReason = InputBox("Please indicate reason for Cancellation of RFP no.", "Reason for Cancel")
    If Replace(MyReason, " ", "") = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of RFP no.!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    
    'Update RFP
    LYRcon.Execute "UPDATE rfp SET `STATUS` = 'CANCELLED', REMARKS = 'CANCELLED' ', ' '" & MyReason & "' WHERE RFID = " & Val(flxSum.TextMatrix(ActRow, 12)) & ";"
               

    'Insert History
    LYRcon.Execute "INSERT INTO transhist(DATE, TRANSNAME, REFID, REMARKS, RESPONSIBLE, ACTION) VALUES " & _
                   "(Now(), 'RFPGSSD'," & flxSum.TextMatrix(ActRow, 12) & ",'" & MyReason & "', " & UserID & ", 'CANCEL') "
    
    flxSum.TextMatrix(ActRow, 8) = "CANCELLED"
    With flxSum     'Change row color to red
        .COL = 1: .Row = ActRow: .ColSel = 11: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = &H8080FF
        .COL = 2
    End With
    

    MsgBox "RFP number was Cancelled!", vbInformation + vbOKOnly


Exit Sub
er:
    DoEvents
    MsgBox err.Description, vbCritical, "Error->CancelRFPGSSD()"
End Sub

'Private Sub mnuConfirm_Click()
'    If MySource = "POBILLINGMC" Then
'        If flxSum.TextMatrix(ActRow, 2) <> "" And flxSum.TextMatrix(ActRow, 2) <> "" Then
'            MsgBox "DR number already Confirmed", vbInformation + vbOKOnly, "System Mesaage"
'            Exit Sub
'        End If
'        Call frmBillingMotor.ShowRecPO(flxSum.TextMatrix(ActRow, 14))
'        frmBillingMotor.Show
'        frmBillDR.Branchcode = frmBillingMotor.lblBranchShipToCode.Caption
'        frmBillDR.Recfrom = ""
'        frmBillDR.Recfrom = frmBillingMotor.lblSupplier.Caption
'        frmBillDR.cmdApply.Visible = False
'        frmBillDR.cmdconfrim.Visible = True
'        frmBillDR.cmdView.Visible = True
'        frmBillDR.cmdClose.Visible = False
'        frmBillDR.ItemcostID = flxSum.TextMatrix(ActRow, 14)
'        frmBillDR.txtDRnumber = flxSum.TextMatrix(ActRow, 3)
'        Call frmBillDR.GetDRModel
'        frmBillDR.Show vbModal
'        Unload Me
'    End If
'End Sub

Private Sub mnuRFPCancel_Click()

End Sub

Private Sub mnuUpdateDev_Click()
    Dim rsdr As New ADODB.Recordset
    Dim Qry As String, pbillid As Long, DRDATE As Date, DRNO As String
    Dim i As Integer, j As Integer

    If MySource = "POBILLINGMC" Or MySource = "POBILLING" Then
        If flxSum.TextMatrix(ActRow, 1) <> "" And flxSum.TextMatrix(ActRow, 2) <> "" Then
            MsgBox "DR number already Confirmed", vbInformation + vbOKOnly, "System Mesaage"
            Exit Sub
        End If
    If MySource = "POBILLING" Then
        Qry = "CALL sp_lyrpurchase.Sp_Loadbranches('" & flxSum.TextMatrix(ActRow, 6) & "', '" & flxSum.TextMatrix(ActRow, 3) & "','" & flxSum.TextMatrix(ActRow, 12) & "');"
    ElseIf MySource = "POBILLINGMC" Then
        If flxSum.TextMatrix(ActRow, 12) = "HONDA PHILIPPINES INC." Then
        Qry = "CALL sp_lyrpurchase.Sp_Loadbranches('" & flxSum.TextMatrix(ActRow, 6) & "', '" & flxSum.TextMatrix(ActRow, 3) & "','" & flxSum.TextMatrix(ActRow, 12) & "');"
        'InputBox Qry, , Qry
    ElseIf flxSum.TextMatrix(ActRow, 12) = "KAWASAKI MOTORS (PHIL)CORP." Or flxSum.TextMatrix(ActRow, 12) = "YAMAHA MOTOR PHIL. INC." Or flxSum.TextMatrix(ActRow, 12) = "MOTOITALIA PHILIPPINES, INC." Or flxSum.TextMatrix(ActRow, 12) = "AUTOITALIA PHILIPPINES ENTERPRISES INC." Or flxSum.TextMatrix(ActRow, 12) = "SUZUKI PHILS., INC." Or flxSum.TextMatrix(ActRow, 12) = "TRIMOTORS TECHNOLOGY CORP." Then
        Qry = "CALL sp_lyrpurchase.Sp_Loadbranches('" & flxSum.TextMatrix(ActRow, 6) & "', '" & flxSum.TextMatrix(ActRow, 2) & "','" & flxSum.TextMatrix(ActRow, 12) & "');"
        End If
     End If
    If rsdr.State = adStateOpen Then rsdr.Close
    rsdr.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsdr.RecordCount = 0 Then
        flxUpdateDr.clear
        flxUpdateDr.Rows = 2
        MsgBox "DR number not available", vbExclamation + vbOKOnly, "System Message"
        Exit Sub
    End If
    Set flxUpdateDr.DataSource = rsdr
    If flxUpdateDr.Rows = 1 Then
        MsgBox "DR not found", vbCritical + vbOKOnly, "System Message"
        Exit Sub
    End If
    With flxUpdateDr
        For i = 1 To flxUpdateDr.Rows - 1
        If flxSum.TextMatrix(ActRow, 4) = flxUpdateDr.TextMatrix(i, 3) Then
                MsgBox "Already Update!", vbExclamation + vbOKOnly, "System Message!"
            Exit Sub
         End If
        Next i
     End With
      For i = 1 To flxUpdateDr.Rows - 1
        If flxSum.TextMatrix(ActRow, 12) = flxUpdateDr.TextMatrix(i, 1) Then
            flxSum.TextMatrix(ActRow, 1) = flxUpdateDr.TextMatrix(i, 4)
            flxSum.TextMatrix(ActRow, 2) = flxUpdateDr.TextMatrix(i, 3)
            
        DRDATE = flxSum.TextMatrix(ActRow, 1)
        DRNO = flxSum.TextMatrix(ActRow, 2)
        Qry = "SELECT purchasebill.PBILLID From  purchasebill Where  purchasebill.INVNO ='" & flxSum.TextMatrix(ActRow, 3) & "' AND purchasebill.Recfrom  ='" & flxSum.TextMatrix(ActRow, 12) & "'"
        pbillid = GetFieldValue(Qry)
        LYRcon.Execute "Update purchasebill SET DRNO ='" & DRNO & "' Where  purchasebill.pbillid = '" & pbillid & "'"
        LYRcon.Execute "Update purchasebill SET DRDATE ='" & Format(DRDATE, "yyyy-mm-dd") & "' Where  purchasebill.pbillid = '" & pbillid & "'"
                Exit Sub
            End If
        Next i
        End If
'    End If
End Sub
Private Sub mnuConShipto_Click()
    Dim rsdr As New ADODB.Recordset
    Dim Qry As String, pbillid As Long, DRDATE As Date, DRNO As String
    Dim i As Integer, j As Integer

    If MySource = "POBILLINGMC" Then
        If flxSum.TextMatrix(ActRow, 1) <> "" And flxSum.TextMatrix(ActRow, 2) <> "" Then
            MsgBox "DR number already Confirmed", vbInformation + vbOKOnly, "System Mesaage"
            Exit Sub
        End If
    If flxSum.TextMatrix(ActRow, 12) = "HONDA PHILIPPINES INC." Then
    Qry = "CALL sp_LoadUrsbrnch('" & flxSum.TextMatrix(ActRow, 3) & "');"
'    InputBox Qry, , Qry
   ElseIf flxSum.TextMatrix(ActRow, 12) = "KAWASAKI MOTORS (PHIL)CORP." Or flxSum.TextMatrix(ActRow, 12) = "YAMAHA MOTOR PHIL. INC." Or flxSum.TextMatrix(ActRow, 12) = "SUZUKI PHILS., INC." Or flxSum.TextMatrix(ActRow, 12) = "TRIMOTORS TECHNOLOGY CORP." Then
    Qry = "CALL sp_LoadDRbrnch('" & flxSum.TextMatrix(ActRow, 3) & "');"
   End If
    If rsdr.State = adStateOpen Then rsdr.Close
    rsdr.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsdr.RecordCount = 0 Then
        flxUpdateDr.clear
        flxUpdateDr.Rows = 2
        MsgBox "DR Number not Found", vbExclamation + vbOKOnly, "System Message"
        Exit Sub
    End If
    Set flxUpdateDr.DataSource = rsdr
    If flxUpdateDr.Rows = 1 Then
        MsgBox "DR number not found", vbCritical + vbOKOnly, "System Message"
        Exit Sub
    End If
    With flxUpdateDr
        For i = 1 To flxUpdateDr.Rows - 1
        If flxSum.TextMatrix(ActRow, 4) = flxUpdateDr.TextMatrix(i, 3) Then
                MsgBox "Already Update!", vbExclamation + vbOKOnly, "System Message"
            Exit Sub
         End If
        Next i
     End With
      For i = 1 To flxUpdateDr.Rows - 1
        If flxSum.TextMatrix(ActRow, 12) = flxUpdateDr.TextMatrix(i, 2) Then
            flxSum.TextMatrix(ActRow, 6) = flxUpdateDr.TextMatrix(i, 1)
'            flxSum.TextMatrix(ActRow, 2) = flxUpdateDr.TextMatrix(i, 3)
      
      
'      DRDATE = flxSum.TextMatrix(ActRow, 1)
'      DRNO = flxSum.TextMatrix(ActRow, 2)
'        Qry = "SELECT purchasebill.PBILLID From  purchasebill Where  purchasebill.INVNO = '" & flxSum.TextMatrix(ActRow, 3) & "'"
'        pbillid = GetFieldValue(Qry)
'        LYRcon.Execute "Update purchasebill SET DRNO ='" & DRNO & "' Where  purchasebill.pbillid = '" & pbillid & "'"
'        LYRcon.Execute "Update purchasebill SET DRDATE ='" & Format(DRDATE, "yyyy-mm-dd") & "' Where  purchasebill.pbillid = '" & pbillid & "'"
'            Exit Sub
'        End If
'
        End If
    Next i
End If
End Sub
Private Sub mnuDisSet_Click()
    Dim rsdr As New ADODB.Recordset
    Dim Qry As String, pbillid As Long, DRDATE As Date, DRNO As String
    Dim i As Integer, j As Integer

    If MySource = "POBILLINGMC" Then
        If flxSum.TextMatrix(ActRow, 1) <> "" And flxSum.TextMatrix(ActRow, 2) <> "" Then
            MsgBox "DR number already Confirmed", vbInformation + vbOKOnly, "System Mesaage"
            Exit Sub
        End If
    If flxSum.TextMatrix(ActRow, 12) = "HONDA PHILIPPINES INC." Then
'    Qry = "CALL sp_lyrpurchase.sp_LoadPRBranch('" & flxSum.TextMatrix(ActRow, 3) & "', '" & flxSum.TextMatrix(ActRow, 6) & "');"
'    InputBox Qry, , Qry
   ElseIf flxSum.TextMatrix(ActRow, 12) = "KAWASAKI MOTORS (PHIL)CORP." Or flxSum.TextMatrix(ActRow, 12) = "YAMAHA MOTOR PHIL. INC." Or flxSum.TextMatrix(ActRow, 12) = "SUZUKI PHILS., INC." Or flxSum.TextMatrix(ActRow, 12) = "TRIMOTORS TECHNOLOGY CORP." Then
'    Qry = "CALL sp_lyrpurchase.sp_LoadPRBranch('" & flxSum.TextMatrix(ActRow, 2) & "', '" & flxSum.TextMatrix(ActRow, 6) & "');"
   End If
    If rsdr.State = adStateOpen Then rsdr.Close
    rsdr.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsdr.RecordCount = 0 Then
        flxUpdateDr.clear
        flxUpdateDr.Rows = 2
        MsgBox "DR number not available", vbExclamation + vbOKOnly, "System Message"
        Exit Sub
    End If
    Set flxUpdateDr.DataSource = rsdr
    If flxUpdateDr.Rows = 1 Then
        MsgBox "DR not found", vbCritical + vbOKOnly, "System Message"
        Exit Sub
    End If
    With flxUpdateDr
        For i = 1 To flxUpdateDr.Rows - 1
        If flxSum.TextMatrix(ActRow, 4) = flxUpdateDr.TextMatrix(i, 3) Then
                MsgBox "Already Update!", vbExclamation + vbOKOnly, "System Message!"
            Exit Sub
         End If
        Next i
     End With
      For i = 1 To flxUpdateDr.Rows - 1
        If flxSum.TextMatrix(ActRow, 12) <> flxUpdateDr.TextMatrix(i, 1) Then
            flxSum.TextMatrix(ActRow, 1) = flxUpdateDr.TextMatrix(i, 4)
            flxSum.TextMatrix(ActRow, 2) = flxUpdateDr.TextMatrix(i, 3)


      DRDATE = flxSum.TextMatrix(ActRow, 1)
      DRNO = flxSum.TextMatrix(ActRow, 2)
        Qry = "SELECT purchasebill.PBILLID From  purchasebill Where  purchasebill.INVNO ='" & flxSum.TextMatrix(ActRow, 3) & "' AND purchasebill.Recfrom  ='" & flxSum.TextMatrix(ActRow, 12) & "'"
        pbillid = GetFieldValue(Qry)
        LYRcon.Execute "Update purchasebill SET DRNO ='" & DRNO & "' Where  purchasebill.pbillid = '" & pbillid & "'"
        LYRcon.Execute "Update purchasebill SET DRDATE ='" & Format(DRDATE, "yyyy-mm-dd") & "' Where  purchasebill.pbillid = '" & pbillid & "'"
            Exit Sub
        End If
     Next i
'End If
End If
End Sub
Private Sub mnuCM_Click()
   If MySource = "RFP" Or MySource = "MOTORCYCLERFP" Then
        If flxSum.TextMatrix(ActRow, 2) <> "" And flxSum.TextMatrix(ActRow, 2) <> "" Then
            MsgBox "NO RECORD TO SHOW", vbInformation + vbOKOnly, "System Mesaage"
            Exit Sub
        End If
        Call frmBillingMotor.ShowRecPO(flxSum.TextMatrix(ActRow, 14))
        frmBillingMotor.Show
        frmBillDR.Branchcode = frmBillingMotor.lblBranchShipToCode.Caption
        frmBillDR.Recfrom = ""
        frmBillDR.Recfrom = frmBillingMotor.lblSupplier.Caption
        frmBillDR.cmdApply.Visible = False
        frmBillDR.cmdconfrim.Visible = True
        frmBillDR.cmdView.Visible = True
        frmBillDR.cmdClose.Visible = False
        frmBillDR.ItemcostID = flxSum.TextMatrix(ActRow, 14)
        frmBillDR.txtDRnumber = flxSum.TextMatrix(ActRow, 3)
        Call frmBillDR.GetDRModel
        frmBillDR.Show vbModal
        Unload Me
    End If
End Sub
Private Sub RXResizer1_ResizeComplete()
    lblHeading(0).Left = Frame2.Left + (Frame2.Width - lblHeading(0).Width) - 10
    
    tabwait.Left = (Me.ScaleWidth - tabwait.Width) / 2
    tabwait.Top = (Me.ScaleHeight - tabwait.Height) / 2
    
    SSTab1.Left = (Me.ScaleWidth - tabwait.Width) / 2
    SSTab1.Top = (Me.ScaleHeight - tabwait.Height) / 2
    
    With flxSum
        If MySource = "PRODUCTS" Then '1        2          3          4              5         6             7         8          9          10
            .FormatString = "   |<PROD. CODE|<MODEL|<MODEL GROUP|<DESCRIPTION|<PRODUCT TYPE|^INITIAL|<ACCOUNT TITLE|<BRAND|<CLASSIFICATION|<PRDID"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.09 'CODE
            .ColWidth(2) = (.Width - 300) * 0.11 'MODEL
            .ColWidth(3) = (.Width - 300) * 0.12 'MODEL GROUP
            .ColWidth(4) = (.Width - 300) * 0.11 'DESCRIPTION
            .ColWidth(5) = (.Width - 300) * 0.12  'TYPEs
            .ColWidth(6) = (.Width - 300) * 0.04 'INITIAL
            .ColWidth(7) = (.Width - 300) * 0.18 'TITLE
            .ColWidth(8) = (.Width - 300) * 0.1  'BRAND
            .ColWidth(9) = (.Width - 300) * 0.12 'CLASS
            .ColWidth(10) = (.Width - 300) * 0 'PRDID
         ElseIf MySource = "Employees" Then
                                    '1        2          3          4              5         6             7         8          9          10            11          12
'            .FormatString = "   |<LASTNAME|<FIRSTNAME|<MIDDLENAME|<ADDRESS|<CONTACTNOS|^GENDER|<RELIGION|<CIVILSTATUS|<BIRTHDATE|<DATEHIRED|<POSITION|<EMPSTATUS|<DEPARTMENT"
            .ColWidth(0) = 300
            .ColWidth(1) = 0 'EMPID
            .ColWidth(2) = (.Width - 300) * 0.08  'LASTNAME
            .ColWidth(3) = (.Width - 300) * 0.08 'FIRSTNAME
            .ColWidth(4) = (.Width - 300) * 0.07 'MIDDLE
            .ColWidth(5) = (.Width - 300) * 0.11  'ADDRESS
            .ColWidth(6) = (.Width - 300) * 0.08 'CONTACTNOS
            .ColWidth(7) = (.Width - 300) * 0.04 'GENDER
            .ColWidth(8) = (.Width - 300) * 0.06 'RELIGION
            .ColWidth(9) = (.Width - 300) * 0.06 'CIVILSTATUS
            .ColWidth(10) = 0 'SPOUSE
            .ColWidth(11) = 0 'SPOUSEEMP
            .ColWidth(12) = (.Width - 300) * 0.08 'BIRTHDATE
            .ColWidth(13) = (.Width - 300) * 0.08 'DATEHIRED
            .ColWidth(14) = (.Width - 300) * 0.08 'POSITION
            .ColWidth(15) = (.Width - 300) * 0.08 'EMPSTATUS
            .ColWidth(16) = 0 'AREAID
            .ColWidth(17) = 0 'EMPPICTURE
            .ColWidth(18) = (.Width - 300) * 0.08 'PCUSER
            .ColWidth(19) = 0  'DEPARTMENT
            .ColWidth(20) = 0  'INACTIVE
            .ColWidth(21) = 0 'UPDATEDBY
            .ColWidth(22) = 0  'LASTUPDATE
            .ColWidth(23) = 0 'USERNAME
            .ColWidth(24) = 0  'PASSWORD
            .ColWidth(25) = 0 'USERTYPE
        ElseIf MySource = "RFP" Or MySource = "MOTORCYCLERFP" Then
             .FormatString = "   |<DATE |<RFPNO|^SOLDTO|<SUPPLIER NAME|<SI# REC|^DR REC|^TTL ANMT|^TTL DEDUCT.|^NET PAYABLE|<DUEDATE|<PREPARED BY|<RECIEVED BY|^APPROVED BY |<RFID |<STATUS"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.06 'DATE
            .ColWidth(2) = (.Width - 300) * 0.05 'RFPNO
            .ColWidth(3) = (.Width - 300) * 0.06 'SOLD TO
            .ColWidth(4) = (.Width - 300) * 0.16 'SUPPLIER NAME
            .ColWidth(5) = (.Width - 300) * 0.04 'TTLINVOICE
            .ColWidth(6) = (.Width - 300) * 0.04 'TTLDR
            .ColWidth(7) = (.Width - 300) * 0.07  'TTL ANMT
            .ColWidth(8) = (.Width - 300) * 0.06  'TTL DEDUCT.
            .ColWidth(9) = (.Width - 300) * 0.08  'NET PAYABLE
            .ColWidth(10) = (.Width - 300) * 0.06   'due
            .ColWidth(11) = (.Width - 300) * 0.1   'prep
            .ColWidth(12) = (.Width - 300) * 0.11  'rec BY
            .ColWidth(13) = (.Width - 300) * 0.11 'APPR BY
            .ColWidth(14) = (.Width - 300) * 0 'RFPID
            .ColWidth(15) = (.Width - 300) * 0.11 'STATUS
            .ColWidth(16) = (.Width - 300) * 0.08 'Print Copy
            
        ElseIf MySource = "CBD" Or MySource = "MCCBD" Then
            .FormatString = "   |<DATE |<RFPNO|<SUPPLIER NAME|<STATUS|<CHECK AMNT|^DUEDATE|<PREPARED BY|<RECEIVED BY|<APPROVED BY|>RFID"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.08 'DATE
            .ColWidth(2) = (.Width - 300) * 0.08 'RFPNO
            .ColWidth(3) = (.Width - 300) * 0.18 'SUPPLIERNAME
            .ColWidth(4) = (.Width - 300) * 0.11 'STATUS
            .ColWidth(5) = (.Width - 300) * 0.12  'CHAECK AMNT
            .ColWidth(6) = (.Width - 300) * 0.09 'DUEDATE
            .ColWidth(7) = (.Width - 300) * 0.18 'PREPARED BY
            .ColWidth(8) = (.Width - 300) * 0.1  'RECEIVED BY
            .ColWidth(9) = (.Width - 300) * 0.12 'APPROVED BY
            .ColWidth(10) = (.Width - 300) * 0 'RFID
            .ColWidth(11) = (.Width - 300) * 0.08 'Print Copy
            
        ElseIf MySource = "CHECKPO" Then
            .FormatString = "|^P.O. DATE |<P.O NUM |^BRANCH NAME |^TOTAL QTY |^INVOICE NUMBER |^STATUS |<SHIP TO"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1
            .ColWidth(2) = (.Width - 300) * 0.1
            .ColWidth(3) = (.Width - 300) * 0.19
            .ColWidth(4) = (.Width - 300) * 0.08
            .ColWidth(5) = (.Width - 300) * 0.12
            .ColWidth(6) = (.Width - 300) * 0.12
            .ColWidth(7) = (.Width - 300) * 0.19
            .ColWidth(8) = 0
         '------------------------------1/09/2018---------------------
        ElseIf MySource = "BRANCH ALLOCATION" Then
            .FormatString = " |^P.O. DATE|^P.O NUMBER|^P.O REF|<SUPPLIER NAME|^BRANCHCODE |^TOTAL PO AMNT|^PREPARED BY |^STATUS |^TERMS|^BRANDNAME"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.08
            .ColWidth(2) = (.Width - 300) * 0.08
            .ColWidth(3) = (.Width - 300) * 0.07
            .ColWidth(4) = (.Width - 300) * 0.17
            .ColWidth(5) = (.Width - 300) * 0.07
            .ColWidth(6) = (.Width - 300) * 0.11
            .ColWidth(7) = (.Width - 300) * 0.11
            .ColWidth(8) = (.Width - 300) * 0.08
            .ColWidth(9) = (.Width - 300) * 0.08
            .ColWidth(10) = 0
            .ColWidth(11) = (.Width - 300) * 0.08
            .ColWidth(12) = (.Width - 300) * 0.05
            .ColWidth(13) = 0
            .ColWidth(14) = 0
        '------------------------------ARIAS ---------------------
        ElseIf MySource = "CHART OF ACCOUNTS" Then
            .FormatString = "  |<ACCOUNT NUMBER|<ACCOUNT TITLE|<ACCOUNT TYPE|<GROUP ACCOUNT|<ACCOUNT REPORT|<NORMAL BALANCE|<REMARKS|STATUS|<ID"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.08
            .ColWidth(2) = (.Width - 300) * 0.12
            .ColWidth(3) = (.Width - 300) * 0.15
            .ColWidth(4) = (.Width - 300) * 0.1
            .ColWidth(5) = (.Width - 300) * 0.12
            .ColWidth(6) = (.Width - 300) * 0.07
            .ColWidth(7) = (.Width - 300) * 0.19
            .ColWidth(8) = (.Width - 300) * 0.15
            .ColWidth(9) = 0
            
        ElseIf MySource = "BACK ORDER" Then
            .FormatString = " |^DATE|^REF.NO.|<PARTNUMBER|<DESCRIPTION|^BBO QTY.|^BBO BAL.|<BRANCH |<P.O. REF#. "
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.08
            .ColWidth(2) = (.Width - 300) * 0.08
            .ColWidth(3) = (.Width - 300) * 0.18
            .ColWidth(4) = (.Width - 300) * 0.2
            .ColWidth(5) = (.Width - 300) * 0.08
            .ColWidth(6) = (.Width - 300) * 0.08
            .ColWidth(7) = (.Width - 300) * 0.15
            .ColWidth(8) = (.Width - 300) * 0.12
        ElseIf MySource = "INST-ACCT" Then
           .FormatString = " |^DATE|^ACCT.NO.|<CUSTOMER|<ADDRESS|^DESCRIPTION|^COD|<LCP |<INV. AMOUNT|<INVOICE TYPE|^TERMS|^SALES AGENT|<COMMISION |<DATE RELEASED|^PCV NO.|^REBCOMM|<REMARKS "
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.05
            .ColWidth(2) = (.Width - 300) * 0.08
            .ColWidth(3) = (.Width - 300) * 0.15
            .ColWidth(4) = (.Width - 300) * 0.1
            .ColWidth(5) = (.Width - 300) * 0.1
            .ColWidth(6) = (.Width - 300) * 0.05
            .ColWidth(7) = (.Width - 300) * 0.03
            .ColWidth(8) = (.Width - 300) * 0.03
            .ColWidth(9) = (.Width - 300) * 0.05
            .ColWidth(10) = (.Width - 300) * 0.05
            .ColWidth(11) = (.Width - 300) * 0.08
            .ColWidth(12) = (.Width - 300) * 0.03
            .ColWidth(13) = (.Width - 300) * 0.03
            .ColWidth(14) = (.Width - 300) * 0.03
            .ColWidth(15) = (.Width - 300) * 0.03
            .ColWidth(16) = (.Width - 300) * 0.1
        ElseIf MySource = "BRANCH" Then
'                                    1         2        3           4        5     6    7      8        9      10      11        12      13       14     15     16      17        18
            .FormatString = " |<BRANCHNAME|^BRANCHCODE|<CGROUP|<MANAGER|<ADDRESS|^TEL|<FAX|<SPARK|<BRANCHTYPE|<TIN|<DATESTART|<REMARKS|<EMAIL|<PRODUCTS|<AREA|<MAIN|<DSN NAME|<BRANCHID "
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.18 'branchname
            .ColWidth(2) = (.Width - 300) * 0.08 'bcode
            .ColWidth(3) = (.Width - 300) * 0.12 'cgroup
            .ColWidth(4) = (.Width - 300) * 0.15  'mgr
            .ColWidth(5) = (.Width - 300) * 0.3   'address
            .ColWidth(6) = (.Width - 300) * 0.09 'tel
            .ColWidth(7) = (.Width - 300) * 0 'fax
            .ColWidth(8) = (.Width - 300) * 0 'spark
            .ColWidth(9) = (.Width - 300) * 0 'branchtype
            .ColWidth(10) = (.Width - 300) * 0 'tin
            .ColWidth(11) = (.Width - 300) * 0 'datestart
            .ColWidth(12) = (.Width - 300) * 0.06 'remarks
            .ColWidth(13) = (.Width - 300) * 0.04 'email
            .ColWidth(14) = (.Width - 300) * 0 'products
            .ColWidth(15) = (.Width - 300) * 0 'area
            .ColWidth(16) = (.Width - 300) * 0 'main
            .ColWidth(17) = (.Width - 300) * 0.07 'dsn name
            .ColWidth(18) = (.Width - 300) * 0 'id
        ElseIf MySource = "AGENT" Then
            .FormatString = " |<AGENT|^RANK|<DATE CLOSED|<ACCT. NO.|<CUST. NAME|<CUST. ADDRESS|<UNIT DESCRIPTION |<REMARKS "
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.11
            .ColWidth(2) = (.Width - 300) * 0.03
            .ColWidth(3) = (.Width - 300) * 0.06
            .ColWidth(4) = (.Width - 300) * 0.06
            .ColWidth(5) = (.Width - 300) * 0.14
            .ColWidth(6) = (.Width - 300) * 0.27
            .ColWidth(7) = (.Width - 300) * 0.25
            .ColWidth(8) = (.Width - 300) * 0.05
            
        ElseIf MySource = "CHART OF ACCOUNTS" Then
            .FormatString = " |<ACCOUNT NO.|^ACCOUNT TITLE|<ACCOUNT TYPE|<GROUP ACCOUNT|<ACCOUNT REPORT|<NORMAL BALANCE|<REMARKS |<STATUS|<ACCTID "
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.07
            .ColWidth(2) = (.Width - 300) * 0.1
            .ColWidth(3) = (.Width - 300) * 0.1
            .ColWidth(4) = (.Width - 300) * 0.12
            .ColWidth(5) = (.Width - 300) * 0.14
            .ColWidth(6) = (.Width - 300) * 0.1
            .ColWidth(7) = (.Width - 300) * 0.2
            .ColWidth(8) = (.Width - 300) * 0.14
            .ColWidth(9) = 0
        ElseIf MySource = "SUPPLIER" Then
                               '1       2         3      4        5       6          7          8          9      10        11         12      13    14     15       16    17     18     19       20
             .FormatString = " |<|<SUPPLIERNAME|<CODE|<ADDRESS|<CONTACT#|<FAX|<CONTACTPERSON |<POSITION |<MOBILE|<EMAIL|<UPDATEDBY|<LASTUPDT |<TEL |<TIN|<INPUTTAX|<TERMS|<BANK |<REM |<BANKACCT|<WEB "
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0 'supplierid
            .ColWidth(2) = (.Width - 300) * 0.17 'suppliername
            .ColWidth(3) = (.Width - 300) * 0.03 'suppcode
            .ColWidth(4) = (.Width - 300) * 0.15  'address
            .ColWidth(5) = (.Width - 300) * 0.1 'contactno
            .ColWidth(6) = (.Width - 300) * 0.07 'fax
            .ColWidth(7) = (.Width - 300) * 0.1 'conperson
            .ColWidth(8) = (.Width - 300) * 0 'position
            .ColWidth(9) = (.Width - 300) * 0 'mobile
            .ColWidth(10) = (.Width - 300) * 0.1 'email
            .ColWidth(11) = (.Width - 300) * 0 'updated by
            .ColWidth(12) = (.Width - 300) * 0 'lastupd
            .ColWidth(13) = (.Width - 300) * 0.1 'tel
            .ColWidth(14) = (.Width - 300) * 0.1 'tin
            .ColWidth(15) = (.Width - 300) * 0.1 'inputtax
            .ColWidth(16) = 0 'TERMS
            .ColWidth(17) = (.Width - 300) * 0.1 'bank
            .ColWidth(18) = (.Width - 300) * 0 'rem
            .ColWidth(19) = (.Width - 300) * 0 'bankacct
            .ColWidth(20) = (.Width - 300) * 0 'web
            .ColWidth(21) = (.Width - 300) * 0.1 'statutory
            .ColWidth(22) = (.Width - 300) * 0 ' billtytpe
            .ColWidth(23) = (.Width - 300) * 0 'suptype
    '--------------------UPDATE BY ARIAS 07/10/2019----------------------------------
        ElseIf MySource = "PURCHASE ORDER" Or MySource = "PR" Or MySource = "PRMC" Or MySource = "PURCHASE ORDER2" Then
            Dim i As Integer
                                 '1           2       3       4           5               6      7        8           9              10                  11
            .FormatString = " |<P.O. DATE|<P.O. REF|^TERMS|^STATUS|<SUPPLIER NAME|<REMARKS|^TTL ITEMS |>TTL PO COST|>TTL PO LESS |> TOTAL PO AMNT |<PREPARED BY |>PURCHASEID "
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.06
            .ColWidth(2) = (.Width - 300) * 0.06
            .ColWidth(3) = (.Width - 300) * 0.06
            .ColWidth(4) = (.Width - 300) * 0.07
            .ColWidth(5) = (.Width - 300) * 0.2
            .ColWidth(6) = (.Width - 300) * 0.19
            .ColWidth(7) = (.Width - 300) * 0.05
            .ColWidth(8) = (.Width - 300) * 0.07
            .ColWidth(9) = (.Width - 300) * 0.04 '
            .ColWidth(10) = (.Width - 300) * 0.07
            .ColWidth(11) = (.Width - 300) * 0.1
            .ColWidth(12) = 0
        ElseIf MySource = "PRCM" Then
                                '1          2           3             4             5          6              7            8           9            10
            .FormatString = "|^RFID |< RF CMDATE |<RFCM NO |<SUPPLIERS NAME |<TOTAL AMNT |<STATUS  |<PREPARED BY  |<CHECKED BY |<NOTED BY  |<APPROVED BY "
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.03
            .ColWidth(2) = (.Width - 300) * 0.08
            .ColWidth(3) = (.Width - 300) * 0.12
            .ColWidth(4) = (.Width - 300) * 0.21
            .ColWidth(5) = (.Width - 300) * 0.08
            .ColWidth(6) = (.Width - 300) * 0.09
            .ColWidth(7) = (.Width - 300) * 0.12
            .ColWidth(8) = (.Width - 300) * 0.1
            .ColWidth(9) = (.Width - 300) * 0.1
            .ColWidth(10) = (.Width - 300) * 0.1
            .ColWidth(11) = 0
          ElseIf MySource = "PRDM" Then
                                   '1       2           3          4          5         6           7           8           9           10           11           12
            .FormatString = "|^DM DATE |^DM NO |<SUPPLIER NAME  |^CM ID |^TTL AMNT  |^TTL DISC |^TTL NETPAY |<STATUS |<PREPARED BY |<APPROVED BY |<CHECK BY |<CERTIFIED BY"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.06
            .ColWidth(2) = (.Width - 300) * 0.07
            .ColWidth(3) = (.Width - 300) * 0.2
            .ColWidth(4) = (.Width - 300) * 0.06
            .ColWidth(5) = (.Width - 300) * 0.07
            .ColWidth(6) = (.Width - 300) * 0.07
            .ColWidth(7) = (.Width - 300) * 0.07
            .ColWidth(8) = (.Width - 300) * 0.07
            .ColWidth(9) = (.Width - 300) * 0.11
            .ColWidth(10) = (.Width - 300) * 0.11
            .ColWidth(11) = (.Width - 300) * 0.11
            .ColWidth(12) = (.Width - 300) * 0.11
            .ColWidth(13) = 0
        ElseIf MySource = "POBILLINGMC" Then
                                '1        2         3        4      5          6            7         8      9        10        11              12        13           14          15          16
            .FormatString = " |^DRDATE|^DRNUMBER|^INVNUM|^P.O NUM|^BILL TO|^SHIP TO|^BILLING STAT |^TTLPO|^TTLINV |^TTLREC |>TTL DR AMNT  |^REC FROM     |^         |^COD RFPNO    |^         |^ CBD RFPNO"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.07
            .ColWidth(2) = (.Width - 300) * 0.07
            .ColWidth(3) = (.Width - 300) * 0.07
            .ColWidth(4) = (.Width - 300) * 0.07
            .ColWidth(5) = (.Width - 300) * 0.05
            .ColWidth(6) = (.Width - 300) * 0.05
            .ColWidth(7) = (.Width - 300) * 0.1
            .ColWidth(8) = (.Width - 300) * 0.05
            .ColWidth(9) = (.Width - 300) * 0.06
            .ColWidth(10) = (.Width - 300) * 0.06
            .ColWidth(11) = (.Width - 300) * 0.07
            .ColWidth(12) = (.Width - 300) * 0.14
            .ColWidth(13) = (.Width - 300) * 0
            .ColWidth(14) = (.Width - 300) * 0.06
            .ColWidth(15) = (.Width - 300) * 0
            .ColWidth(16) = (.Width - 300) * 0.06
            .ColWidth(17) = (.Width - 300) * 0
        ElseIf MySource = "POBILLING" Then
                                 '1        2         3        4      5          6            7         8      9        10        11            12             13         14      15          16
            .FormatString = " |^DRDATE|^DRNUMBER|^INVNUM|^P.O NUM|^BILL TO|^SHIP TO|^BILLING STAT |^TTLPO|^TTLINV |^TTLREC |>TTL DR AMNT    |^REC FROM   |^         |^ COD RFPNO |^   |^CBD RFPNO"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.07
            .ColWidth(2) = (.Width - 300) * 0.07
            .ColWidth(3) = (.Width - 300) * 0.07
            .ColWidth(4) = (.Width - 300) * 0.07
            .ColWidth(5) = (.Width - 300) * 0.05
            .ColWidth(6) = (.Width - 300) * 0.05
            .ColWidth(7) = (.Width - 300) * 0.1
            .ColWidth(8) = (.Width - 300) * 0.05
            .ColWidth(9) = (.Width - 300) * 0.06
            .ColWidth(10) = (.Width - 300) * 0.06
            .ColWidth(11) = (.Width - 300) * 0.07
            .ColWidth(12) = (.Width - 300) * 0.14
            .ColWidth(13) = 0
            .ColWidth(14) = (.Width - 300) * 0.06
            .ColWidth(15) = 0
            .ColWidth(16) = (.Width - 300) * 0.06
            .ColWidth(17) = 0
        ElseIf MySource = "Price List" Then
            .FormatString = "  |^EFFECTIVE DATE|^REF. NO.|^PROD. TYPE|<PRODUCT BRAND|<DESCRIPTION |^ACTIVE |^FINAL |^ |^ |^ |^ |^ |^DAY`S ACTIVE "
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1
            .ColWidth(2) = (.Width - 300) * 0.11
            .ColWidth(3) = (.Width - 300) * 0.1
            .ColWidth(4) = (.Width - 300) * 0.15
            .ColWidth(5) = (.Width - 300) * 0.21
            .ColWidth(6) = (.Width - 300) * 0.1
            .ColWidth(7) = (.Width - 300) * 0.1
            .ColWidth(8) = 0
            .ColWidth(9) = 0
            .ColWidth(10) = 0
            .ColWidth(11) = 0
            .ColWidth(12) = 0
            .ColWidth(13) = (.Width - 300) * 0.1
            For i = 1 To .Rows - 1
                '.TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                .TextMatrix(i, 6) = IIf(.TextMatrix(i, 6) = "1" Or .TextMatrix(i, 6) = "X", "X", "")
                .TextMatrix(i, 7) = IIf(.TextMatrix(i, 7) = "1" Or .TextMatrix(i, 7) = "X", "X", "")
                If .TextMatrix(i, 7) = "" Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .COL = 1: .Row = i: .ColSel = 13: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .COL = 13
                End If
            Next i

        ElseIf MySource = "SPCRR" Then
            'DDATECREATED, SPCRRNO, RRNO, RRDATE, BRANCH, SUPPLIER, PREPAREDBY, TERMS, INVSTATUS, INVOICETTLPAY
            .FormatString = "  |^DATE|^SPCRR NO.|^RR NO. |^RR DATE |<BRANCH |<SUPPLIER |<PREPARED BY |^TERMS |^STATUS |>TOTAL PAYABLE"
            .ColWidth(0) = 200
            .ColWidth(1) = (.Width - 300) * 0.06 'DATE
            .ColWidth(2) = (.Width - 300) * 0.07  'SPCRR NO.
            .ColWidth(3) = (.Width - 300) * 0.07  'RR NO.
            .ColWidth(4) = (.Width - 300) * 0.07  'RR DATE
            .ColWidth(5) = (.Width - 300) * 0.17   'BRANCH
            .ColWidth(6) = (.Width - 300) * 0.17   'SUPPLIER
            .ColWidth(7) = (.Width - 300) * 0.15 'PREPARED BY
            .ColWidth(8) = (.Width - 300) * 0.05 'TERMS
            .ColWidth(9) = (.Width - 300) * 0.06 'STATUS
            .ColWidth(10) = (.Width - 300) * 0.08 'TOTAL PAYABLE
            .ColWidth(11) = 0 'SPCRRID
            
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 9) = "PENDING" Then
                    .COL = 1: .Row = i: .ColSel = 10: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .COL = 10
                ElseIf .TextMatrix(i, 9) = "CANCELLED" Then
                    .COL = 1: .Row = i: .ColSel = 10: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
                    .COL = 10
                Else
                    .COL = 1: .Row = i: .ColSel = 10: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = vbWhite
                    .COL = 10
                End If
            Next i
            
        ElseIf MySource = "FA" Or MySource = "MP" Then

          '                  0    1      2        3          4            5           6            7           8        9              10           11
            .FormatString = "  |^DATE|^RFP NO.|<SOLDTO  |^ITEMS |<PREPARED BY |<CHECKED BY |<APPROVED BY |^STATUS |^P.O. TYPE |>TOTAL PAYABLE |^PRINT COPY "
            .ColWidth(0) = 200
            .ColWidth(1) = (.Width - 300) * 0.05 'DATE
            .ColWidth(2) = (.Width - 300) * 0.05  'RFP NO.
            .ColWidth(3) = (.Width - 300) * 0.14   'SOLDTO
            .ColWidth(4) = (.Width - 300) * 0.04  'ITEMS
            .ColWidth(5) = (.Width - 300) * 0.1  'PREPARED BY
            .ColWidth(6) = (.Width - 300) * 0.1 'CHECKED BY
            .ColWidth(7) = (.Width - 300) * 0.1 'APPROVED BY
            .ColWidth(8) = (.Width - 300) * 0.14 'STATUS
            .ColWidth(9) = (.Width - 300) * 0.05 'PO TYPE
            .ColWidth(10) = (.Width - 300) * 0.12 'TOTAL PAYABLE
            .ColWidth(11) = (.Width - 300) * 0.07  'Print Copy.
            .ColWidth(12) = (.Width - 300) * 0 'RFP ID.
        
        ElseIf MySource = "RFPOTHER" Then

          '                  0    1      2        3          4            5           6            7           8        9              10           11
            .FormatString = "  |^DATE|^RFP NO.|<SOLDTO  |^ITEMS |<PREPARED BY |<CHECKED BY |<APPROVED BY |^STATUS |^STATUTORY |>TOTAL PAYABLE |^PRINT COPY "
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
            .ColWidth(12) = (.Width - 300) * 0 'RFP ID.
            
        End If
    End With
    
End Sub

Private Sub txtdr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtdr.Text = Empty
        flxSum.TextMatrix(ActRow, 6) = txtdr.Tag
        txtdr.Tag = "skiplostfocus"
        flxSum.SetFocus
        flxSum.COL = 6
        txtdr.Visible = False
    End If
End Sub
Private Sub txtdr_KeyPress(KeyAscii As Integer)
    If KeyAscii > 26 Then KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxSum.SetFocus
        flxSum.COL = 6
    End If
End Sub
Private Sub txtdr_LostFocus()
    If txtdr.Tag = "skiplostfocus" Then
        txtdr.Tag = ""
        Exit Sub
    End If
    flxSum.TextMatrix(ActRow, 6) = txtdr.Text
    txtdr.Visible = False
End Sub
Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txt.Text = Empty
        flxSum.TextMatrix(ActRow, 2) = txt.Tag
        txt.Tag = "skiplostfocus"
        flxSum.SetFocus
        flxSum.COL = 2
        txt.Visible = False
    End If
End Sub
Private Sub txt_KeyPress(KeyAscii As Integer)
    If KeyAscii > 26 Then KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxSum.SetFocus
        flxSum.COL = 2
    End If
End Sub
Private Sub txt_LostFocus()
    If txt.Tag = "skiplostfocus" Then
        txt.Tag = ""
        Exit Sub
    End If
    flxSum.TextMatrix(ActRow, 2) = txt.Text
    txt.Visible = False
End Sub

'Private Sub RXResizer2_ResizeComplete()
'    If MySource = "RFP" Then
'
'
'
'
'    ElseIf MySource = "BRANCH ALLOCATION" Then
'
'
'    ElseIf MySource = "POBILLING" Then
'
'
'    ElseIf MySource = "POBILLINGMC" Then
'
'End Sub

Private Sub Timer1_Timer()
    Dim i As Integer
    If MySource = "PRODUCTS" Then
        FixRange = True
        DailyRange = False
        MonthlyRange = False
        TCol = 9
    ElseIf MySource = "CHART OF ACCOUNTS" Then
        FixRange = True
        DailyRange = False
        MonthlyRange = False
        TCol = 9
    ElseIf MySource = "BACK ORDER" Then
        FixRange = False
        DailyRange = True
        MonthlyRange = False
        TCol = 9
        If DateStart <> "" Then
            Dfrom.Value = DateValue(DateStart)
        Else
            Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        End If
        If DateEnd <> "" Then
            Dto.Value = DateValue(DateEnd)
        Else
            Dto.Value = Date
        End If
        ShowDate = 2
    ElseIf MySource = "INST-ACCT" Then
        FixRange = False
        DailyRange = True
        MonthlyRange = False
        TCol = 8
        If DateStart <> "" Then
            Dfrom.Value = DateValue(DateStart)
        Else
            Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        End If
        If DateEnd <> "" Then
            Dto.Value = DateValue(DateEnd)
        Else
            Dto.Value = Date
        End If
        ShowDate = 2
     ElseIf MySource = "Employees" Then
        FixRange = True
        DailyRange = False
        MonthlyRange = False
        TCol = 13
    ElseIf MySource = "BRANCH" Then
        FixRange = False
        DailyRange = False
        MonthlyRange = False
        TCol = 8
    ElseIf MySource = "SUPPLIER" Then
        FixRange = False
        DailyRange = False
        MonthlyRange = False
        TCol = 8
    ElseIf MySource = "AGENT" Then
        FixRange = False
        DailyRange = True
        MonthlyRange = False
        TCol = 8
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto.Value = Date
        Call RefreshRecord(2)
    ElseIf MySource = "PURCHASE ORDER" Or MySource = "PURCHASE ORDER2" Or MySource = "PR" Or MySource = "PRMC" Or MySource = "BRANCH ALLOCATION" Then
        FixRange = False
        DailyRange = True
        MonthlyRange = False
        TCol = 8
        If DateStart <> "" Then
            Dfrom.Value = DateValue(DateStart)
        Else
            Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        End If
        If DateEnd <> "" Then
            Dto.Value = DateValue(DateEnd)
        Else
            Dto.Value = Date
        End If
        ShowDate = 2
    ElseIf MySource = "POBILLING" Or MySource = "POBILLINGMC" Or MySource = "CHECKPO" Then
        FixRange = False
        DailyRange = True
        MonthlyRange = False
        TCol = 9
        If DateStart <> "" Then
            Dfrom.Value = DateValue(DateStart)
        Else
            Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        End If
        If DateEnd <> "" Then
            Dto.Value = DateValue(DateEnd)
        Else
            Dto.Value = Date
        End If
        ShowDate = 2
    ElseIf MySource = "RFP" Or MySource = "CBD" Or MySource = "MCCBD" Or MySource = "MOTORCYCLERFP" Or MySource = "PRCM" Or MySource = "PRDM" Then
        FixRange = False
        DailyRange = True
        MonthlyRange = False
        TCol = 10
        If DateStart <> "" Then
            Dfrom.Value = DateValue(DateStart)
        Else
            Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        End If
        If DateEnd <> "" Then
            Dto.Value = DateValue(DateEnd)
        Else
            Dto.Value = Date
        End If
        ShowDate = 2
    ElseIf MySource = "Price List" Then
        FixRange = True
        DailyRange = False
        MonthlyRange = False
        TCol = 9
        Dfrom.Visible = False
        Dto.Visible = False
        
'        cmdNext.Visible = False
'        cmdPrev.Visible = False
'        cmdRefresh.Visible = False
'        Label1.Visible = False
'        ShowDateRange = True
'        Call ShowCommands(1)

    ElseIf MySource = "SPCRR" Then
        FixRange = False
        DailyRange = True
        MonthlyRange = False
        TCol = 11
        Dto.Value = Date
        Dfrom.Value = DateAdd("D", -30, Date)
        ShowDate = 2
        framStatus.Visible = True
        cboStat.ListIndex = 0
        
    ElseIf MySource = "FA" Or MySource = "MP" Or MySource = "RFPOTHER" Then
        FixRange = False
        DailyRange = True
        MonthlyRange = False
        cmdFilter.Enabled = False
        TCol = 12
        Dto.Value = Date
        Dfrom.Value = DateAdd("D", -30, Date)
        ShowDate = 2
        
    End If
    
    If DailyRange Then
        Frame1.Visible = True
        Frame3.Visible = False
    ElseIf MonthlyRange Then
        Frame1.Visible = False
        Frame3.Visible = True
        For i = 1 To 12
            cboMonth.AddItem MonthName(i)
        Next i
        UpdYear.Value = Year(Date)
        TXTYear.Text = Year(Date)
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto.Value = DateAdd("d", -1, DateAdd("m", 1, Dfrom.Value))
        cboMonth.ListIndex = Month(Dfrom.Value) - 1
    ElseIf FixRange Then
        Frame1.Visible = False
        Frame3.Visible = False

    End If
    Call frmHSummary.RefreshRecord(2)
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














