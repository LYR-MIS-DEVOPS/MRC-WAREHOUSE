VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSPNew 
   Caption         =   "PARTNUMBER"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   13080
   StartUpPosition =   3  'Windows Default
   Tag             =   "xyf"
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSPsum 
      Height          =   3375
      Left            =   120
      TabIndex        =   18
      Tag             =   "xyf"
      Top             =   3360
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   5953
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Tag             =   "xyf"
      Top             =   600
      Width           =   12855
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2400
         TabIndex        =   20
         Tag             =   "xyf"
         Top             =   1800
         Width           =   4815
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   2400
         TabIndex        =   17
         Tag             =   "xyf"
         Top             =   1080
         Width           =   3615
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   9840
         TabIndex        =   16
         Tag             =   "xyf"
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   9840
         TabIndex        =   15
         Tag             =   "xyf"
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   9840
         TabIndex        =   14
         Tag             =   "xyf"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   9840
         TabIndex        =   13
         Tag             =   "xyf"
         Top             =   720
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2400
         TabIndex        =   12
         Tag             =   "xyf"
         Top             =   1440
         Width           =   4815
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   6120
         TabIndex        =   11
         Tag             =   "xyf"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2400
         TabIndex        =   10
         Tag             =   "xyf"
         Top             =   720
         Width           =   4815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Tag             =   "xyf"
         Top             =   330
         Width           =   4815
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "APPLICABLE MODEL"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Tag             =   "xyf"
         Top             =   1830
         Width           =   1935
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "MAX. DISCOUNT"
         Height          =   255
         Left            =   7800
         TabIndex        =   8
         Tag             =   "xyf"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "RETAIL PRICE"
         Height          =   255
         Left            =   7800
         TabIndex        =   7
         Tag             =   "xyf"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "COST"
         Height          =   255
         Left            =   7800
         TabIndex        =   6
         Tag             =   "xyf"
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "PRICE FACTOR"
         Height          =   255
         Left            =   7800
         TabIndex        =   5
         Tag             =   "xyf"
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "TYPE"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Tag             =   "xyf"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "SUPPLIER"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Tag             =   "xyf"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "DESCRIPTION"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Tag             =   "xyf"
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "PARTNUMBER"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Tag             =   "xyf"
         Top             =   360
         Width           =   1935
      End
   End
   Begin lvButton.lvButtons_H cmdAdd 
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "Add &New"
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
      Image           =   "frmSPNew.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdEdit 
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   22
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
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
      Image           =   "frmSPNew.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Index           =   3
      Left            =   2880
      TabIndex        =   23
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
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
      Image           =   "frmSPNew.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   375
      Index           =   4
      Left            =   4320
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
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
      Image           =   "frmSPNew.frx":10CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Index           =   5
      Left            =   5760
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
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
      Image           =   "frmSPNew.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   9120
      TabIndex        =   26
      Top             =   3000
      Width           =   1440
      _ExtentX        =   2540
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
      Image           =   "frmSPNew.frx":1C02
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   13200
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   13320
      m_MDI_Height    =   7470
   End
End
Attribute VB_Name = "frmSPNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



