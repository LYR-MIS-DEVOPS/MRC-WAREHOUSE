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
Begin VB.Form SPorderSum 
   Caption         =   "Spare Parts P.O. Summary"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9120
   ScaleWidth      =   12960
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flxOrder 
      Height          =   7935
      Left            =   0
      TabIndex        =   62
      Tag             =   "XY"
      Top             =   1125
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   13996
      _Version        =   393216
      Rows            =   3
      Cols            =   30
      FixedRows       =   2
      RowHeightMin    =   275
      BackColorBkg    =   -2147483633
      AllowUserResizing=   1
      FormatString    =   $"SPOrderSum.frx":0000
   End
   Begin TabDlg.SSTab SSConfirmTab 
      Height          =   1575
      Left            =   13320
      TabIndex        =   57
      Tag             =   "XYF"
      Top             =   3120
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   2778
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Confirmation"
      TabPicture(0)   =   "SPOrderSum.frx":013E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblConfirmation"
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
         TabIndex        =   58
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
         TabIndex        =   59
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
         Image           =   "SPOrderSum.frx":015A
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   375
         Left            =   2040
         TabIndex        =   60
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
         Image           =   "SPOrderSum.frx":06F4
         cBack           =   -2147483633
      End
      Begin VB.Label lblConfirmation 
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
         TabIndex        =   61
         Tag             =   "XYF"
         Top             =   480
         Width           =   1935
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   7
      Left            =   11880
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "SPOrderSum.frx":0C8E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   10080
      TabIndex        =   56
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Save to Excel"
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
      Image           =   "SPOrderSum.frx":1228
      cBack           =   -2147483633
   End
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   12240
      Top             =   0
   End
   Begin VB.Timer tmrAutoLoad 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   11760
      Top             =   0
   End
   Begin VB.TextBox txtQTYCONFIRM 
      Height          =   285
      Left            =   12120
      TabIndex        =   55
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   50
      TabIndex        =   40
      Tag             =   "XY"
      Top             =   9240
      Width           =   12855
      Begin VB.Shape Shape8 
         BorderStyle     =   6  'Inside Solid
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   10440
         Tag             =   "XY"
         Top             =   255
         Width           =   135
      End
      Begin VB.Shape Shape7 
         BorderStyle     =   6  'Inside Solid
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   7920
         Tag             =   "XY"
         Top             =   615
         Width           =   135
      End
      Begin VB.Shape Shape6 
         BorderStyle     =   6  'Inside Solid
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   7920
         Tag             =   "XY"
         Top             =   255
         Width           =   135
      End
      Begin VB.Shape Shape5 
         BorderStyle     =   6  'Inside Solid
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   4320
         Tag             =   "XY"
         Top             =   620
         Width           =   135
      End
      Begin VB.Shape Shape4 
         BorderStyle     =   6  'Inside Solid
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   4320
         Tag             =   "XY"
         Top             =   260
         Width           =   135
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   6  'Inside Solid
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   240
         Tag             =   "XY"
         Top             =   620
         Width           =   135
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BorderColor     =   &H80000007&
         BorderStyle     =   6  'Inside Solid
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   240
         Tag             =   "XY"
         Top             =   260
         Width           =   135
      End
      Begin VB.Label Label16 
         Caption         =   "Order is Cancelled"
         Height          =   255
         Left            =   11280
         TabIndex        =   54
         Tag             =   "XY"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "RED:"
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
         TabIndex        =   53
         Tag             =   "XY"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "Not yet Printed"
         Height          =   255
         Left            =   9000
         TabIndex        =   52
         Tag             =   "XY"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Final Process"
         Height          =   255
         Left            =   9000
         TabIndex        =   51
         Tag             =   "XY"
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label12 
         Caption         =   "YELLOW:"
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
         Left            =   8160
         TabIndex        =   50
         Tag             =   "XY"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "WHITE:"
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
         Left            =   8160
         TabIndex        =   49
         Tag             =   "XY"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Shipment is not Confirm"
         Height          =   255
         Left            =   5760
         TabIndex        =   48
         Tag             =   "XY"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "No Qty Confirmed Encoded"
         Height          =   255
         Left            =   5760
         TabIndex        =   47
         Tag             =   "XY"
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "LIGHT GRAY:"
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
         TabIndex        =   46
         Tag             =   "XY"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "ORANGE:"
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
         TabIndex        =   45
         Tag             =   "XY"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Not yet Confirm"
         Height          =   255
         Left            =   2040
         TabIndex        =   44
         Tag             =   "XY"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Fax/Email is not yet Confirm"
         Height          =   255
         Left            =   2040
         TabIndex        =   43
         Tag             =   "XY"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "SKY BLUE:"
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
         Left            =   480
         TabIndex        =   42
         Tag             =   "XY"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "YELLOW GREEN:"
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
         Left            =   480
         TabIndex        =   41
         Tag             =   "XY"
         Top             =   240
         Width           =   1575
      End
   End
   Begin TabDlg.SSTab tabCancel 
      Height          =   1935
      Left            =   4080
      TabIndex        =   30
      Top             =   9360
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3413
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "SPOrderSum.frx":17C2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Shape1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdTabCancel"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Pbar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin MSComctlLib.ProgressBar Pbar 
         Height          =   135
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
      End
      Begin lvButton.lvButtons_H cmdTabCancel 
         Height          =   375
         Left            =   1320
         TabIndex        =   32
         Top             =   1320
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         Caption         =   "OK"
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
         Image           =   "SPOrderSum.frx":17DE
         cBack           =   -2147483633
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   255
         Left            =   20
         Top             =   20
         Width           =   3700
      End
      Begin VB.Label Label2 
         Caption         =   "Please wait cancel on progress..."
         Height          =   255
         Left            =   960
         TabIndex        =   31
         Top             =   600
         Width           =   2415
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   360
         Picture         =   "SPOrderSum.frx":1D78
         Stretch         =   -1  'True
         Top             =   480
         Width           =   360
      End
   End
   Begin VB.TextBox txtSuppID 
      Height          =   285
      Left            =   240
      TabIndex        =   26
      Top             =   9240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   50
      TabIndex        =   11
      Tag             =   "XYF"
      Top             =   9360
      Width           =   12855
      Begin VB.TextBox txtRemarks 
         Alignment       =   2  'Center
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
         Left            =   9480
         TabIndex        =   23
         Tag             =   "XYF"
         Top             =   503
         Width           =   3255
      End
      Begin VB.TextBox txtBackOrder 
         Alignment       =   2  'Center
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
         Left            =   9480
         TabIndex        =   22
         Tag             =   "XYF"
         Top             =   173
         Width           =   1095
      End
      Begin VB.TextBox txtRec 
         Alignment       =   2  'Center
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
         Left            =   6120
         TabIndex        =   21
         Tag             =   "XYF"
         Top             =   503
         Width           =   1095
      End
      Begin VB.TextBox txtOrder 
         Alignment       =   2  'Center
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
         Left            =   6120
         TabIndex        =   20
         Tag             =   "XYF"
         Top             =   173
         Width           =   1095
      End
      Begin VB.ComboBox cboPO 
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
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   480
         Width           =   2895
      End
      Begin VB.ComboBox cboSupplier 
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
         TabIndex        =   18
         Tag             =   "XYF"
         Top             =   150
         Width           =   2895
      End
      Begin VB.Label lblPctBO 
         Caption         =   "0.00%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   10680
         TabIndex        =   25
         Tag             =   "XYF"
         Top             =   203
         Width           =   855
      End
      Begin VB.Label lblPctRec 
         Caption         =   "0.00%"
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
         Left            =   7320
         TabIndex        =   24
         Tag             =   "XYF"
         Top             =   533
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "REMARKS"
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
         Left            =   8280
         TabIndex        =   17
         Tag             =   "XYF"
         Top             =   518
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "BACK ORDER:"
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
         Left            =   8280
         TabIndex        =   16
         Tag             =   "XYF"
         Top             =   188
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "QTY. RECEIVED"
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
         Left            =   4800
         TabIndex        =   15
         Tag             =   "XYF"
         Top             =   518
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "QTY. ORDERED"
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
         Left            =   4800
         TabIndex        =   14
         Tag             =   "XYF"
         Top             =   188
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "P.O. NUMBER"
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
         TabIndex        =   13
         Tag             =   "XYF"
         Top             =   518
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "SUPPLIER"
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
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   188
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   700
      Left            =   45
      TabIndex        =   8
      Tag             =   "XYF"
      Top             =   400
      Width           =   12855
      Begin VB.TextBox txtFind 
         Height          =   315
         Left            =   10080
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   230
         Width           =   2655
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   8760
         TabIndex        =   10
         Tag             =   "XY"
         Top             =   195
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
         Image           =   "SPOrderSum.frx":2302
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dFrom 
         Height          =   315
         Left            =   1440
         TabIndex        =   34
         Top             =   180
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   141950977
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   375
         Left            =   3240
         TabIndex        =   35
         Tag             =   "XYF"
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
         Image           =   "SPOrderSum.frx":289C
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   375
         Left            =   5400
         TabIndex        =   36
         ToolTipText     =   "Execute"
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
         Image           =   "SPOrderSum.frx":2E36
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dTo 
         Height          =   315
         Left            =   3720
         TabIndex        =   37
         Top             =   180
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   141950977
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdPrev 
         Height          =   375
         Left            =   5880
         TabIndex        =   38
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
         Image           =   "SPOrderSum.frx":33D0
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
         Index           =   6
         Left            =   120
         TabIndex        =   39
         Top             =   210
         Width           =   1335
      End
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4560
      Top             =   8280
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4560
      Top             =   7800
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   8280
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   7800
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&View Order"
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
      Image           =   "SPOrderSum.frx":396A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Grid View"
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
      Image           =   "SPOrderSum.frx":3F04
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "SPOrderSum.frx":449E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   8400
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "SPOrderSum.frx":4A38
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Back Order"
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
      Image           =   "SPOrderSum.frx":4FD2
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Refresh"
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
      Image           =   "SPOrderSum.frx":556C
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxOrd 
      Height          =   1095
      Left            =   240
      TabIndex        =   28
      Top             =   9600
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1931
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      FormatString    =   "Date   |<PO Number    |<Orders"
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSP 
      Height          =   3015
      Left            =   7800
      TabIndex        =   29
      Top             =   9480
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComDlg.CommonDialog dlgDEMO 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblRow 
      Height          =   255
      Left            =   10320
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblP 
      Caption         =   "è"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   8280
      Width           =   255
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   480
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   13200
      m_MDI_Height    =   9690
   End
   Begin VB.Menu mnuOrder 
      Caption         =   "Sales Report"
      Visible         =   0   'False
      Begin VB.Menu mnuFaxDate 
         Caption         =   "Confirm Fax/Email Date"
      End
      Begin VB.Menu mnuConfirmDate 
         Caption         =   "Confirmation Date"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuShipDate 
         Caption         =   "Confirm Ship Date"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel P.O."
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "SPorderSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRPT As New ADODB.Recordset
Dim rsSummary As New ADODB.Recordset
Dim rsSupplier As New ADODB.Recordset
Dim rsSP As New ADODB.Recordset
Dim BtnIndx As Integer, OSrow As Integer, fRow As Integer, fCol As Integer, ActCol As Long
Dim ans As Integer, myDcode As Integer
Dim Qry As String, myConfirmStatus As String
Dim SuppSrch As Boolean
Public MySource As String
Dim prevRowSel As Long

Public Sub RefreshRecord(dCode As Integer)
    Dim i As Integer
    myDcode = dCode
    
    Qry = "call sp_mrcwarehouse.sp_sporder_summary(" & Val(dCode) & ",'" & Format(Dfrom.Value, "YYYY-MM-DD") & "','" & Format(Dto.Value, "YYYY-MM-DD") & "','" & MySource & "')"
    
    If rsSummary.State = adStateOpen Then rsSummary.Close
    rsSummary.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsSummary.RecordCount = 0 Then
        flxOrder.Clear
        flxOrder.Rows = 2
        flxOrder.Rows = 3
        flxOrder.Cols = 30
        Call RXResizer1_ResizeComplete
        Exit Sub
    Else
        With rsSummary
            .MoveFirst
            flxOrder.Rows = .RecordCount + 2
            For i = 2 To flxOrder.Rows - 1
                'PURCHASE ORDER DETAILS
                flxOrder.TextMatrix(i, 1) = Format$(!ORDERDATE, "MM/DD/YYYY")
                flxOrder.TextMatrix(i, 2) = !SUPPLIERNAME
                flxOrder.TextMatrix(i, 3) = !REFERENCE
                flxOrder.TextMatrix(i, 4) = !ORDERTYPE
                flxOrder.TextMatrix(i, 5) = IIf(IsNull(!TTLPN), 0, !TTLPN)
                flxOrder.TextMatrix(i, 6) = IIf(IsNull(!TTLORDER), 0, !TTLORDER)
                flxOrder.TextMatrix(i, 7) = IIf(IsNull(!TTLAMOUNT), 0, Format$(!TTLAMOUNT, "###,###,###.00"))
                flxOrder.TextMatrix(i, 8) = IIf(IsNull(!CONFIRMNO), "", !CONFIRMNO)
                flxOrder.TextMatrix(i, 9) = !PURCHASEDBY
                flxOrder.TextMatrix(i, 10) = IIf(IsNull(!FAXDATE), "", Format$(!FAXDATE, "MM/DD/YYYY"))
                'CONFIRMATION(7DAYS)
                flxOrder.TextMatrix(i, 11) = IIf(IsNull(!ConfirmRDays), "", !ConfirmRDays)
                flxOrder.TextMatrix(i, 12) = IIf(IsNull(!CONFIRMDATE), "", Format$(!CONFIRMDATE, "MM/DD/YYYY"))
                flxOrder.TextMatrix(i, 13) = IIf(IsNull(!ConfirmLT), "", !ConfirmLT)
                'SHIPMENT(7DAYS)
                flxOrder.TextMatrix(i, 14) = IIf(IsNull(!ShipRDays), "", !ShipRDays)
                flxOrder.TextMatrix(i, 15) = IIf(IsNull(!SHIPDATE), "", Format$(!SHIPDATE, "MM/DD/YYYY"))
                flxOrder.TextMatrix(i, 16) = IIf(IsNull(!ShipLT), "", !ShipLT)
                'ARRIVAL(28DAYS)
                flxOrder.TextMatrix(i, 17) = IIf(IsNull(!RecRDays), "", !RecRDays)
                flxOrder.TextMatrix(i, 18) = IIf(IsNull(!DATEREC), "", !DATEREC)
                flxOrder.TextMatrix(i, 19) = IIf(IsNull(!RecLT), "", !RecLT)
                flxOrder.TextMatrix(i, 20) = IIf(IsNull(!AVELT), "", !AVELT)
                'SERVICE RATIO
                flxOrder.TextMatrix(i, 21) = IIf(IsNull(!QTYORDERED), 0, !QTYORDERED)
                flxOrder.TextMatrix(i, 22) = IIf(IsNull(!QTYCONFIRM), 0, !QTYCONFIRM)
                flxOrder.TextMatrix(i, 23) = IIf(IsNull(!PERCCONFIRM), "0%", !PERCCONFIRM)
                flxOrder.TextMatrix(i, 24) = IIf(IsNull(!RECQTY), 0, !RECQTY)
                flxOrder.TextMatrix(i, 25) = IIf(IsNull(!PERCSERVED), "0%", !PERCSERVED)
                flxOrder.TextMatrix(i, 26) = IIf(IsNull(!BACKORDER), 0, !BACKORDER)
                
                flxOrder.TextMatrix(i, 27) = IIf(IsNull(!SPORDERID), 0, !SPORDERID)
                flxOrder.TextMatrix(i, 28) = IIf(IsNull(!PrintCopy), 0, !PrintCopy)
                flxOrder.TextMatrix(i, 29) = IIf(IsNull(!Remarks), "", !Remarks)
                .MoveNext
            Next i
            .Close
            'Set flxOrder.DataSource = rsSummary
        End With
        
        With flxOrder
            For i = 2 To .Rows - 1

'                If Val(.TextMatrix(i, 22)) > 0 Then 'NO QTY CONFIRMED
'                    .Col = 1: .row = i: .ColSel = 26: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 255, 255)
'                    .Col = 2
'                End If
'
'                If Val(.TextMatrix(i, 22)) = 0 Then 'NO QTY CONFIRMED
'                    .Col = 1: .row = i: .ColSel = 26: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(211, 211, 211)
'                    .Col = 2
'                End If
'
'                If .TextMatrix(i, 15) = "" Then 'NO SHIP DATE
'                    .Col = 1: .row = i: .ColSel = 26: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 165, 0)
'                    .Col = 2
'                End If
'
'                If .TextMatrix(i, 12) = "" Then 'NO CONFIRMATION DATE
'                    .Col = 1: .row = i: .ColSel = 26: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(135, 206, 250)
'                    .Col = 2
'                End If
'
'                If .TextMatrix(i, 10) = "" Then 'NO FAX DATE
'                    .Col = 1: .row = i: .ColSel = 26: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(173, 255, 47)
'                    .Col = 2
'                End If
'
                If Val(.TextMatrix(i, 28)) = 0 Then 'PRINT COPY NOT YET PRINTED or FINAL
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .col = 1: .row = i: .ColSel = 26: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .col = 2
                ElseIf Val(.TextMatrix(i, 28)) = 1 Then
                    .col = 1: .row = i: .ColSel = 26: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 255, 255)
                    .col = 2
                End If
                If InStr(.TextMatrix(i, 29), "CANCELLED") > 0 Then 'REMARKS
                    .col = 1: .row = i: .ColSel = 26: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(240, 0, 0)
                    .col = 2
                End If
            Next i
        End With
        
    End If
    ActRow = 1
    flxOrder.Visible = True
    'If mySrc = "Branch Purchase Order" Then Call DisplayRef
    Call RXResizer1_ResizeComplete
End Sub


Private Sub cboPO_Change()
    Dim rsPO As New ADODB.Recordset
    Dim Qry As String, i As Integer, TtlRec As Long, TtlOrd As Long
    
    
    If cboPO.Text <> "" Then
    
        With flxOrder
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 3) = cboPO.Text Then
                    txtOrder.Text = Format(.TextMatrix(i, 10), "#,##0")
                    txtRec.Text = Format(.TextMatrix(i, 11), "#,##0")
                    TtlOrd = PureNum(.TextMatrix(i, 10))
                    TtlRec = PureNum(.TextMatrix(i, 11))
                    txtBackOrder.Text = Format(TtlOrd - TtlRec, "#,##0")
                    If TtlOrd > 0 And TtlRec > 0 Then
                        lblPctRec.Caption = Format(TtlRec / TtlOrd, "#0.00%")
                        lblPctBO.Caption = Format((TtlOrd - TtlRec) / TtlOrd, "#0.00%")
                    Else
                        lblPctRec.Caption = ""
                        lblPctBO.Caption = ""
                    End If
                    txtRemarks.Text = "Updated " & .TextMatrix(i, 13)
                    .col = 1: .row = i: .ColSel = 13: .RowSel = i
                    Exit Sub
                End If
            Next i
        End With
        MsgBox "P.O. Number not found!", vbExclamation + vbOKOnly, "Back Order"
        txtOrder.Text = ""
        txtRec.Text = ""
        lblPctRec.Caption = ""
        txtBackOrder.Text = ""
        lblPctBO.Caption = ""
        txtRemarks.Text = ""
    End If
    
    
End Sub

Private Sub cboPO_Click()
    Call cboPO_Change
End Sub

Private Sub cboSupplier_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
    If SuppSrch Then
        SuppSrch = False
        Exit Sub
    End If
    With cboSupplier
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsSupplier.MoveFirst
            rsSupplier.Move i
            txtSuppID.Text = rsSupplier!supplierid
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                SuppSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtSuppID.Text = ""
        End If
    End With
End Sub

Private Sub cboSupplier_Click()
    Dim rsOrder As New ADODB.Recordset
    Dim Qry As String
    
    Call cboSupplier_Change
    If txtSuppID.Text <> "" Then
        'Display P.O. of Suppliers
        '***Access
        'Qry = "SELECT SPORDER.ORDERDATE, SPORDER.REFERENCE, SUPPLIER.SUPPLIERNAME, SPORDER.ORDERTYPE, SPORDER.CONFIRMNO, SPORDER.SUBMITTEDBY, SPORDER.APPROVEDBY, SPORDER.PRINTCOPY, SPORDER.QTYORDER, SPORDER.QTYREC, SPORDER.BACKORDER, SPORDER.BOLASTUPDATE " & _
              "FROM SPORDER INNER JOIN SUPPLIER ON SPORDER.SUPPLIERID = SUPPLIER.SUPPLIERID Where (((SUPPLIER.SUPPLIERID) = " & txtSuppID.Text & ")) ORDER BY SPORDER.ORDERDATE DESC;"
        'mySQL
        Qry = "SELECT SPORDER.ORDERDATE, SPORDER.REFERENCE, SUPPLIER.SUPPLIERNAME, SPORDER.ORDERTYPE, SPORDER.CONFIRMNO, SPORDER.SUBMITTEDBY, SPORDER.APPROVEDBY, SPORDER.PRINTCOPY, SPORDER.QTYORDER, SPORDER.QTYREC, SPORDER.BACKORDER, SPORDER.BOLASTUPDATE " & _
              "FROM SPORDER INNER JOIN SUPPLIER ON SPORDER.SUPPLIERID = SUPPLIER.SUPPLIERID Where SUPPLIER.SUPPLIERID = " & txtSuppID.Text & " ORDER BY SPORDER.ORDERDATE DESC;"
        
        With rsOrder
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                flxOrder.Rows = 1
                cboPO.Clear
                TtlOrd = 0
                TtlRec = 0
                Do While Not .EOF
                    flxOrder.AddItem vbTab & vbTab & Format$(!ORDERDATE, "MM/DD/YYYY") & vbTab & !REFERENCE & vbTab & !SUPPLIERNAME & vbTab & !ORDERTYPE & vbTab & !CONFIRMNO & vbTab & !SubmittedBy & vbTab & !ApprovedBy & vbTab & !PrintCopy & vbTab & IIf(IsNull(!QtyOrder), 0, !QtyOrder) & vbTab & IIf(IsNull(!QtyRec), 0, Format(!QtyRec, "#,##0")) & vbTab & IIf(IsNull(!BACKORDER), 0, Format(!BACKORDER, "#,##0")) & vbTab & IIf(IsNull(!BOLastUpdate), "", !BOLastUpdate)
                    cboPO.AddItem !REFERENCE
                    TtlOrd = TtlOrd + IIf(IsNull(!QtyOrder), 0, !QtyOrder)
                    TtlRec = TtlRec + IIf(IsNull(!QtyRec), 0, !QtyRec)
                    .MoveNext
                Loop
            Else
                flxOrder.Rows = 1
                flxOrder.Rows = 2
                TtlOrd = 0
                TtlRec = 0
            End If
            .Close
        End With
        txtOrder.Text = Format(TtlOrd, "#,##0")
        txtRec.Text = Format(TtlRec, "#,##0")
        txtBackOrder.Text = Format(TtlOrd - TtlRec, "#,##0")
        If TtlOrd > 0 And TtlRec > 0 Then
            lblPctRec.Caption = Format(TtlRec / TtlOrd, "#0.00%")
            lblPctBO.Caption = Format((TtlOrd - TtlRec) / TtlOrd, "#0.00%")
        Else
            lblPctRec.Caption = ""
            lblPctBO.Caption = ""
        End If
        With flxOrder
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 9)) = 0 Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .col = 1: .row = i: .ColSel = 12: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .col = 2
                End If
            Next i
        End With
        
        
    End If
End Sub

Private Sub cboSupplier_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(cboSupplier.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then SuppSrch = True
End Sub

Private Sub cboSupplier_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cboSupplier_Click
End Sub

Private Sub cmdCancel_Click()
SSConfirmTab.Visible = False
End Sub

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxOrder
        For i = fRow To .Rows - 1
            For j = fCol To 26
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    With flxOrder
                        .col = 1: .row = i: .ColSel = 26: .RowSel = i: .TopRow = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(135, 206, 250)
                        .col = 2
                    End With
                    fRow = IIf(j = 26, i + 1, i)
                    fCol = IIf(j = 26, 1, j + 1)
                    Exit Sub
                End If
            Next j
            fCol = 1
            If Val(flxOrder.TextMatrix(i, 28)) = 0 Then
                .col = 1: .row = i: .ColSel = 26: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                .col = 2
            ElseIf InStr(flxOrder.TextMatrix(i, 29), "CANCELLED") > 0 Then 'REMARKS
                .col = 1: .row = i: .ColSel = 26: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(240, 0, 0)
                .col = 2
            Else
                .col = 1: .row = i: .ColSel = 26: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 255, 255)
                .col = 2
            End If
        Next i
        OSrow = fRow
    End With
    MsgBox "No match found!", vbInformation + vbOKOnly, "Find"
    fRow = 2
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

Private Sub cmdOK_Click()
Dim myFaxDate As String, myShipDate As String, myConfirmDate As String

If myConfirmStatus = "FAXDATE" Then
    myFaxDate = ""
    myFaxDate = GetFieldValue("Select FAXDATE From sporder WHERE SPORDERID = " & Val(flxOrder.TextMatrix(OSrow, 27)) & " ")
    
    If myFaxDate = "" Then
        LYRcon.Execute "Update SPORDER SET FAXDATE = '" & Format(txtConfirmDate.Text, "YYYY-MM-DD") & "' " & _
                       "Where SPORDERID = " & Val(flxOrder.TextMatrix(OSrow, 27)) & " "
    End If
    
    MsgBox "Date Fax/Email was successfully Updated", vbInformation + vbOKOnly, "Validation"
    
ElseIf myConfirmStatus = "SHIPDATE" Then
    myShipDate = ""
    myShipDate = GetFieldValue("Select SHIPDATE From sporder WHERE SPORDERID = " & Val(flxOrder.TextMatrix(OSrow, 27)) & " ")
    
    If myShipDate = "" Then
        LYRcon.Execute "Update SPORDER SET SHIPDATE = '" & Format(txtConfirmDate.Text, "YYYY-MM-DD") & "' " & _
                       "Where SPORDERID = " & Val(flxOrder.TextMatrix(OSrow, 27)) & " "
    End If
    
    MsgBox "Shipment Date was successfully Updated", vbInformation + vbOKOnly, "Validation"
    
ElseIf myConfirmStatus = "CONFIRMDATE" Then
    myConfirmDate = ""
    myConfirmDate = GetFieldValue("Select CONFIRMDATE From sporder WHERE SPORDERID = " & Val(flxOrder.TextMatrix(OSrow, 27)) & " ")
    
    If myConfirmDate = "" Then
        LYRcon.Execute "Update SPORDER SET CONFIRMDATE = '" & Format(txtConfirmDate.Text, "YYYY-MM-DD") & "' " & _
                       "Where SPORDERID = " & Val(flxOrder.TextMatrix(OSrow, 27)) & " "
    End If

    MsgBox "Confirmation Date was successfully Updated", vbInformation + vbOKOnly, "Validation"
End If

    SSConfirmTab.Visible = False
    Call cmdRefresh_Click
End Sub

Private Sub cmdPrev_Click()
    cmdNext.Tag = "1"
    cmdNext.Visible = True
    cmdPrev.Visible = False
    Dto.Visible = False
    cmdRefresh.Visible = False
    Call RefreshRecord(1)
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

Private Sub cmdTrans_Click(Index As Integer)
On Error GoTo er
    Dim XLapp As Excel.Application
    Dim wbkMVR As Excel.Workbook
    Dim wksLineA1 As Excel.Worksheet
    Dim sSaveAsFileName
    Dim i As Integer, j As Integer
    Dim MyQry As String
    
    Select Case Index
        Case 0  'Grid View
            BtnIndx = 1
            tmrFloat.Enabled = True
            Call PrintPreView(False)
           ' Call cboMonth_Click
        Case 1  'Exit
            MotoRep.tbMain.Visible = True
            Unload Me
        Case 2  'SP P.O.
            With flxOrder
                If .TextMatrix(OSrow, 3) <> "" Then
                    If .TextMatrix(OSrow, 5) = "(06) - BRANCH B.O." Then
                        Call BrnBackOrderPO.DisplaySPOrder(.TextMatrix(OSrow, 3))
                    Else
                        If OSrow = 0 And .Rows > 2 Then OSrow = 1
                        If MySource = "WAREHOUSE" Then
                            SPordering.Show
                            Call SPordering.DisplaySPOrder(.TextMatrix(OSrow, 3))
                        Else
                            SPorderingWS.Show
                            Call SPorderingWS.DisplaySPOrder(.TextMatrix(OSrow, 3))
                        End If
                    End If
                End If
            End With
            Unload Me
        Case 3
            With flxOrder
                If .TextMatrix(OSrow, 3) <> "" Then
                    If OSrow = 0 And .Rows > 2 Then OSrow = 1
                    SPBackOrder.Show
                    Call SPBackOrder.DisplaySPOrder(.TextMatrix(OSrow, 3))
                End If
            End With
            Unload Me
        Case 4  'Refresh
            Call Form_Load
        Case 5  'Print Preview

            Qry = "call sp_mrcwarehouse.sp_sporder_summary('" & myDcode & "','" & Format(Dfrom.Value, "yyyy-mm-dd") & "','" & Format(Dto.Value, "yyyy-mm-dd") & "','" & MySource & "')"
            frmPrinCDSummary.QryPrn = Qry
            frmPrinCDSummary.ParamValue1 = SysBranch
            frmPrinCDSummary.ParamValue2 = SysAddress
            If MySource = "WHOLESALE" Then
                frmPrinCDSummary.ParamValue3 = MySource & " WAREHOUSE PURCHASE ORDER SUMMARY"
            Else
                frmPrinCDSummary.ParamValue3 = "RETAIL " & MySource & " PURCHASE ORDER SUMMARY"
            End If
            frmPrinCDSummary.ParamValue4 = "DATE COVERED: " & Format(Dfrom.Value, "MMMM dd, YYYY") & " TO " & Format(Dto.Value, "MMMM dd, YYYY")

            frmPrinCDSummary.MySource = "Purchase Order Summary Report"
            frmPrinCDSummary.Show
        Case 6

            With dlgDEMO
                .CancelError = True
                .Flags = cdlOFNHideReadOnly
                .filename = ""
                .filter = "Microsoft Excel (*.xls)|*.xls|All Files(*.*)|*.*"
                .ShowOpen
                If dlgDEMO.filename = "" Then Exit Sub
                sSaveAsFileName = .filename
            End With
            
            Screen.MousePointer = vbHourglass
            Me.Enabled = False
            Set XLapp = CreateObject("Excel.Application")
            XLapp.Visible = True
            XLapp.Workbooks.Add
            Set wbkMVR = XLapp.ActiveWorkbook
            Set wksLineA1 = wbkMVR.Sheets(1)

            wksLineA1.Name = "SUPPLIER'S PO SUMMARY"
            wksLineA1.Cells(1, 1).Value = "PO DATE"
            wksLineA1.Cells(1, 2).Value = "SUPPLIER"
            wksLineA1.Cells(1, 3).Value = "PO NUMBER"
            wksLineA1.Cells(1, 4).Value = "ORDER TYPE"
            wksLineA1.Cells(1, 5).Value = "TOTAL PN"
            wksLineA1.Cells(1, 6).Value = "TOTAL ORDER"
            wksLineA1.Cells(1, 7).Value = "TOTAL AMOUNT"
            wksLineA1.Cells(1, 8).Value = "CONFIRM NO"
            wksLineA1.Cells(1, 9).Value = "PURCHASED BY"
            wksLineA1.Cells(1, 10).Value = "DATE FAX"
            wksLineA1.Cells(1, 11).Value = "DAYS"
            wksLineA1.Cells(1, 12).Value = "DATE CONFIRMED"
            wksLineA1.Cells(1, 13).Value = "LEAD TIME"
            wksLineA1.Cells(1, 14).Value = "DAYS"
            wksLineA1.Cells(1, 15).Value = "DATE SHIPPED"
            wksLineA1.Cells(1, 16).Value = "LEAD TIME"
            wksLineA1.Cells(1, 17).Value = "DAYS"
            wksLineA1.Cells(1, 18).Value = "DATE RECEIVED"
            wksLineA1.Cells(1, 19).Value = "LEAD TIME"
            wksLineA1.Cells(1, 20).Value = "AVE. LEAD TIME"
            wksLineA1.Cells(1, 21).Value = "QTY ORDERED"
            wksLineA1.Cells(1, 22).Value = "QTY CONFIRMED"
            wksLineA1.Cells(1, 23).Value = "% CONFIRMED"
            wksLineA1.Cells(1, 24).Value = "QTY SERVED"
            wksLineA1.Cells(1, 25).Value = "% SERVED"
            wksLineA1.Cells(1, 26).Value = "BACK ORDER"
            
            With flxOrder
                For i = 2 To .Rows - 1
                    For j = 1 To 26
                        wksLineA1.Cells((i - 1) + 1, j).Value = .TextMatrix(i, j)
                    Next j
                DoEvents
                Next i
            End With
            
            wksLineA1.Range("A1:Z" & (i - 1)).Columns.AutoFit
            wksLineA1.Range("A1:Z1").Font.Bold = True
            XLapp.DisplayAlerts = False
        
            
            wbkMVR.SaveAs filename:=sSaveAsFileName, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
            
            wbkMVR.Close SaveChanges:=False
            XLapp.Quit
            
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            
            Set XLapp = Nothing
            Set wbkMVR = Nothing
            Set wksLineA1 = Nothing
            MsgBox "Finish Saving to Excel!", vbInformation + vbOKOnly, Me.Caption
     
        Case 7  'Exit
            'Warehouse.tbMain.Visible = True
            Unload Me
    End Select
    Exit Sub
er:
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    MsgBox Err.Description, vbCritical, "Error!"
End Sub

Private Sub flxOrder_Click()
On Error GoTo er

    prevRowSel = OSrow

    ActCol = flxOrder.col
    OSrow = flxOrder.row
    With flxOrder
        If Val(flxOrder.TextMatrix(prevRowSel, 28)) = 0 And InStr(flxOrder.TextMatrix(prevRowSel, 29), "CANCELLED") <= 0 Then
            .col = 1: .row = prevRowSel: .ColSel = 26: .RowSel = prevRowSel: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
            .col = 2
        ElseIf InStr(flxOrder.TextMatrix(prevRowSel, 29), "CANCELLED") > 0 Then 'REMARKS
            .col = 1: .row = prevRowSel: .ColSel = 26: .RowSel = prevRowSel: .FillStyle = flexFillRepeat: .CellBackColor = RGB(240, 0, 0)
            .col = 2
        Else
            .col = 1: .row = prevRowSel: .ColSel = 26: .RowSel = prevRowSel: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 255, 255)
            .col = 2
        End If
        
        .col = 1: .row = OSrow: .ColSel = 26: .RowSel = OSrow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(135, 206, 250)
        .col = 2
    End With
    
    
    If ActCol = 0 Or OSrow = 0 Then Exit Sub
     
    Select Case ActCol
    Case 22
        If flxOrder.TextMatrix(OSrow, 10) <> "" And flxOrder.TextMatrix(OSrow, 12) <> "" Then
            If Val(flxOrder.TextMatrix(OSrow, 24)) = 0 Then
                With txtQTYCONFIRM
                    If flxOrder.TextMatrix(OSrow, 22) <> "" Then .Tag = flxOrder.TextMatrix(OSrow, 22)
                    .Top = flxOrder.CellTop + flxOrder.Top
                    .Left = flxOrder.CellLeft + flxOrder.Left
                    .Height = flxOrder.CellHeight
                    .Width = flxOrder.CellWidth
                    .Text = flxOrder.Text
                    .Visible = True
                    .Enabled = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .ZOrder
                    .SetFocus
                End With
            Else
                MsgBox "Item was already Served! Qty Confirm is not allowed", vbInformation + vbOKOnly, "Validation"
                flxOrder.SetFocus
                flxOrder.row = OSrow
                flxOrder.col = ActCol
                Exit Sub
            End If
        Else
            MsgBox "Confirmation Date is empty! Adding Qty Confirm is not allowed", vbInformation + vbOKOnly, "Validation"
            flxOrder.SetFocus
            flxOrder.row = OSrow
            flxOrder.col = ActCol
            Exit Sub
        End If
    End Select
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->flxOrder_Click()")
End Sub



Private Sub flxOrder_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call flxOrder_Click
    End If
End Sub

Private Sub flxORDER_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim myPOSITION As String

    myPOSITION = ""
    myPOSITION = GetFieldValue("Select POSITION From employees Where EMPID = " & Val(UserID) & " ")
    If flxOrder.TextMatrix(OSrow, 1) <> "" Then
        If myPOSITION = "Admin. Officer" Or myPOSITION = "CEO" Or myPOSITION = "Purchaser" Or myPOSITION = "I.T. Personnel" Then
            If InStr(flxOrder.TextMatrix(OSrow, 29), "CANCELLED") > 0 Then Exit Sub
            
            If flxOrder.TextMatrix(OSrow, 10) = "" Then
                If Button = vbRightButton Then
                    mnuCancel.Visible = True
                    mnuFaxDate.Visible = True
                    mnuConfirmDate.Visible = False
                    mnuShipDate.Visible = False
                    Call PopupMenu(mnuOrder)
                End If
            ElseIf flxOrder.TextMatrix(OSrow, 12) = "" Then
                If Button = vbRightButton Then
                    mnuCancel.Visible = True
                    mnuFaxDate.Visible = False
                    mnuConfirmDate.Visible = True
                    mnuShipDate.Visible = False
                    Call PopupMenu(mnuOrder)
                End If
            ElseIf flxOrder.TextMatrix(OSrow, 10) <> "" And flxOrder.TextMatrix(OSrow, 12) <> "" And flxOrder.TextMatrix(OSrow, 15) = "" Then
                If Button = vbRightButton Then
                    mnuCancel.Visible = True
                    mnuFaxDate.Visible = False
                    mnuConfirmDate.Visible = False
                    mnuShipDate.Visible = True
                    Call PopupMenu(mnuOrder)
                End If
            ElseIf flxOrder.TextMatrix(OSrow, 15) = "" Then
               If Button = vbRightButton Then
                    mnuCancel.Visible = True
                    mnuFaxDate.Visible = False
                    mnuConfirmDate.Visible = False
                    mnuShipDate.Visible = True
                    Call PopupMenu(mnuOrder)
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim rsOrder As New ADODB.Recordset
    Dim Qry As String
    Dim i As Integer, TtlOrd As Long, TtlRec As Long
On Error GoTo er

    DoEvents
    tmrRefresh.Enabled = False

    tmrAutoLoad.Enabled = True
    myConfirmStatus = ""
    DoEvents
    Call Load(frm_OnProcess2)
    frm_OnProcess2.Show vbModal
    
    
Exit Sub
er:
    tmrAutoLoad.Enabled = False
    Call modObjTools.ErrorMessage(Err.Description, "Error!->Form_Load()")
End Sub
Private Sub PrintPreView(Ppvw As Boolean)
    If Ppvw Then
'        Frame1.Visible = False
        flxOrder.Visible = False
        vwrDmCm.Visible = True
    Else
        vwrDmCm.Visible = False
 '       Frame1.Visible = True
        flxOrder.Visible = True
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsRPT.State = adStateOpen Then rsRPT.Close
    If rsSupplier.State = adStateOpen Then rsSupplier.Close
End Sub






Private Sub mnuCancel_Click()
    Dim ans As Integer, MyBO As Integer
    Dim Qry As String, MyAcc As String
    
Dim myPOSITION As String

    myPOSITION = ""
    myPOSITION = GetFieldValue("Select POSITION From employees Where EMPID = " & Val(UserID) & " ")
    If myPOSITION = "Admin. Officer" Or myPOSITION = "CEO" Or myPOSITION = "Purchaser" Or myPOSITION = "I.T. Personnel" Then
        ans = MsgBox("Do you want to Cancel P.O. " & flxOrder.TextMatrix(OSrow, 3) & "?", vbQuestion + vbYesNo, "Cancel")
        If ans = vbNo Then Exit Sub
            
        'UPDATE ORDER FIELDS TO CANCELLED
        '***Access
        'LYRcon.Execute "UPDATE SPORDER SET SPORDER.CONFIRMNO = 'CANCELLED', SPORDER.REMARKS = 'CANCELLED' WHERE (((SPORDER.REFERENCE)='" & flxOrder.TextMatrix(OSrow, 3) & "'));"
        'LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE = 'CANCELLED' WHERE (((SPITR.REFERENCE)='" & flxOrder.TextMatrix(OSrow, 3) & "'));"
        
        'mySQL
        LYRcon.Execute "UPDATE SPORDER SET SPORDER.CONFIRMNO = 'CANCELLED', SPORDER.REMARKS = 'CANCELLED' WHERE SPORDER.REFERENCE='" & flxOrder.TextMatrix(OSrow, 3) & "';"
        LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE = 'CANCELLED' WHERE SPITR.REFERENCE='" & flxOrder.TextMatrix(OSrow, 3) & "';"
        
        'Update BackOrder of items in P.O.
        '***Access
        'Qry = "SELECT SPITR.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.ORDER " & _
              "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((SPITR.REFERENCE)='" & flxOrder.TextMatrix(OSrow, 3) & "'));"
        'mySQL
        Qry = "SELECT SPITR.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.`ORDER` " & _
              "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE SPITR.REFERENCE='" & flxOrder.TextMatrix(OSrow, 3) & "';"
        
        If rsSP.State = adStateOpen Then rsSP.Close
        rsSP.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        Set flxSP.DataSource = rsSP
        flxSP.Refresh
        tabCancel.Top = flxOrder.Top + 800
        tabCancel.Left = (Me.ScaleWidth - tabCancel.Width) / 2
        tabCancel.Visible = True
        tabCancel.ZOrder
        Frame2.Enabled = False
        For i = 0 To 6
            cmdTrans(i).Enabled = False
        Next i
        cmdTabCancel.Enabled = False
        With flxSP
            Pbar.Value = 0
            Pbar.Max = .Rows - 1
            For i = 1 To .Rows - 1
                'Start Progress bar
                MyBO = GetBackOrder(.TextMatrix(i, 1))
                '***Access
                'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.BACKORDER = " & MyBO & _
                               " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 1) & "));"
                'mySQL
                LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.BACKORDER = " & MyBO & _
                               " WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 1) & ";"
                Pbar.Value = Pbar.Value + 1
                DoEvents
                pPct = (Pbar.Value / Pbar.Max)
                DoEvents
            Next i
            DoEvents
        End With
        Frame2.Enabled = True
        For i = 0 To 6
            cmdTrans(i).Enabled = True
        Next i
        tabCancel.Visible = False
        With flxOrder
            .TextMatrix(OSrow, 29) = "CANCELLED"
            .col = 1: .row = OSrow: .ColSel = 26: .RowSel = OSrow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(240, 0, 0)
            .col = 2
        End With
        
        MsgBox "P.O. " & flxOrder.TextMatrix(OSrow, 3) & " Cancelled!", vbInformation + vbOKOnly
    Else
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
End Sub

Function GetBackOrder(MyID As Integer) As Integer
    Dim rsOrder As New ADODB.Recordset
    Dim rsRec As New ADODB.Recordset
    Dim Qry As String, MyPO As String
    Dim xBal As Integer, i As Integer
    
    'Clear FlxStock
    '***Access
    'Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.ORDER From SPITR " & _
          "WHERE (((SPITR.TRANSTYPE)='ORDER') AND ((SPITR.SPAREPARTID)=" & MyId & ")) ORDER BY SPITR.TRANSDATE;"
    'mySQL
    Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.`ORDER` From SPITR " & _
          "WHERE SPITR.TRANSTYPE='ORDER' AND SPITR.SPAREPARTID=" & MyID & " ORDER BY SPITR.TRANSDATE;"
    
    flxOrd.Rows = 1
    With rsOrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxOrd.AddItem Format(!TRANSDATE, "mm/dd/yyyy") & vbTab & !REFERENCE & vbTab & !Order
                .MoveNext
            Loop
        End If
        .Close
    End With
    xBal = 0
    With flxOrd
        For i = 1 To .Rows - 1
            'Display PO in Stock Card and obtain Balance
            xBal = xBal + Val(.TextMatrix(i, 2))
            MyPO = .TextMatrix(i, 1)
            
            '***Access
            'Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.RECEIVE, SPITR.PACKINGLIST From SPITR " & _
                  "WHERE (((SPITR.TRANSTYPE)='RECEIVE') AND ((SPITR.SPAREPARTID)=" & MyId & ") AND ((SPITR.PACKINGLIST)='" & MyPO & "')) ORDER BY SPITR.TRANSDATE;"
            'mySQL
            Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.RECEIVE, SPITR.PACKINGLIST From SPITR " & _
                  "WHERE SPITR.TRANSTYPE='RECEIVE' AND SPITR.SPAREPARTID=" & MyID & " AND SPITR.PACKINGLIST='" & MyPO & "' ORDER BY SPITR.TRANSDATE;"
            
            rsRec.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If rsRec.RecordCount > 0 Then
                rsRec.MoveFirst
                Do While Not rsRec.EOF
                    xBal = xBal - rsRec!RECEIVE
                    rsRec.MoveNext
                Loop
            End If
            rsRec.Close
        Next i
    End With
    GetBackOrder = xBal
End Function

Private Sub mnuFaxDate_Click()
    If flxOrder.TextMatrix(OSrow, 10) = "" Then
        SSConfirmTab.Visible = True
        txtConfirmDate.Text = Format(Now, "mm/dd/yyyy")
        With SSConfirmTab
            .Top = (Me.Height - .Height) \ 2
            .Left = (Me.Width - .Width) \ 2
            myConfirmStatus = "FAXDATE"
            .Visible = True
            .ZOrder
        End With
    ElseIf flxOrder.TextMatrix(OSrow, 10) <> "" Then
        MsgBox "Fax/Email Date is already confirm", vbInformation + vbOKOnly, "Validation"
        Exit Sub
    End If
End Sub

Private Sub mnuConfirmDate_Click()
    If flxOrder.TextMatrix(OSrow, 12) = "" Then
        SSConfirmTab.Visible = True
        txtConfirmDate.Text = Format(Now, "mm/dd/yyyy")
        With SSConfirmTab
            .Top = (Me.Height - .Height) \ 2
            .Left = (Me.Width - .Width) \ 2
            myConfirmStatus = "CONFIRMDATE"
            .Visible = True
            .ZOrder
        End With
    ElseIf flxOrder.TextMatrix(OSrow, 12) <> "" Then
        MsgBox "Confirmation Date is already confirm", vbInformation + vbOKOnly, "Validation"
        Exit Sub
'    Else
'        MsgBox "Not Yet Printed! Please Print to Finalize the Order", vbInformation + vbOKOnly, "Validation"
'        Exit Sub
    End If
End Sub



Private Sub mnuShipDate_Click()

    If flxOrder.TextMatrix(OSrow, 15) = "" Then
        SSConfirmTab.Visible = True
        txtConfirmDate.Text = Format(Now, "mm/dd/yyyy")
        With SSConfirmTab
            .Top = (Me.Height - .Height) \ 2
            .Left = (Me.Width - .Width) \ 2
            myConfirmStatus = "SHIPDATE"
            .Visible = True
            .ZOrder
        End With
    ElseIf flxOrder.TextMatrix(OSrow, 15) <> "" Then
        MsgBox "Shipment Date is already confirm", vbInformation + vbOKOnly, "Validation"
        Exit Sub
'    Else
'        MsgBox "Not Yet Printed! Please Print to Finalize the Order", vbInformation + vbOKOnly, "Validation"
'        Exit Sub
    End If
End Sub

Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer
    With flxOrder

        '|^P.O. DATE |<SUPPLIER NAME |^PO NUMBER |^ORDER TYPE |^TOTAL PN |^ORDER QTY |>TTL AMOUNT |^CONFIRMATION NO. |^DATE FAX |^PURCHASED BY
        .TextMatrix(0, 1) = "PURCHASE ORDER DETAILS"        'PO DATE
        .TextMatrix(0, 2) = "PURCHASE ORDER DETAILS"        'SUPPLIER NAME
        .TextMatrix(0, 3) = "PURCHASE ORDER DETAILS"        'PO NUMBER
        .TextMatrix(0, 4) = "PURCHASE ORDER DETAILS"        'ORDER TYPE
        .TextMatrix(0, 5) = "PURCHASE ORDER DETAILS"        'TOTAL PN
        .TextMatrix(0, 6) = "PURCHASE ORDER DETAILS"        'TOTAL ORDER
        .TextMatrix(0, 7) = "PURCHASE ORDER DETAILS"        'TTL AMOUNT
        .TextMatrix(0, 8) = "PURCHASE ORDER DETAILS"        'CONFIRM #
        .TextMatrix(0, 9) = "PURCHASE ORDER DETAILS"        'PURCHASED BY
        .TextMatrix(0, 10) = "PURCHASE ORDER DETAILS"       'DATE FAX
        
        .TextMatrix(0, 11) = "CONFIRMATION"                 'RUNNING DAYS
        .TextMatrix(0, 12) = "CONFIRMATION"                 'DATE CONFIRM
        .TextMatrix(0, 13) = "CONFIRMATION"                 'LEAD TIME

        .TextMatrix(0, 14) = "SHIPMENT"                     'RUNNING DAYS
        .TextMatrix(0, 15) = "SHIPMENT"                     'DATE SHIPPED
        .TextMatrix(0, 16) = "SHIPMENT"                     'LEAD TIME
        
        .TextMatrix(0, 17) = "ARRIVAL"                      'RUNNING DAYS
        .TextMatrix(0, 18) = "ARRIVAL"                      'DATE RECEIVED
        .TextMatrix(0, 19) = "ARRIVAL"                      'LEAD TIME
        .TextMatrix(0, 20) = "ARRIVAL"                      'AVERAGE LEAD TIME
        
        .TextMatrix(0, 21) = "SERVICE RATIO"                'QTY ORDERED
        .TextMatrix(0, 22) = "SERVICE RATIO"                'QTY CONFIRMED
        .TextMatrix(0, 23) = "SERVICE RATIO"                '% CONFIRMED
        .TextMatrix(0, 24) = "SERVICE RATIO"                'QTY SERVED
        .TextMatrix(0, 25) = "SERVICE RATIO"                '% SERVED
        .TextMatrix(0, 26) = "SERVICE RATIO"                'BACK ORDER
        
        
        .TextMatrix(1, 1) = "PO DATE"
        .TextMatrix(1, 2) = "SUPPLIER"
        .TextMatrix(1, 3) = "PO NUMBER"
        .TextMatrix(1, 4) = "ORDER TYPE"
        .TextMatrix(1, 5) = "TOTAL PN"
        .TextMatrix(1, 6) = "TOTAL ORDER"
        .TextMatrix(1, 7) = "TOTAL AMOUNT"
        .TextMatrix(1, 8) = "CONFIRM NO"
        .TextMatrix(1, 9) = "PURCHASED BY"
        .TextMatrix(1, 10) = "DATE FAX"
        
        'CONFIRMATION(7DAYS)
        .TextMatrix(1, 11) = "DAYS"
        .TextMatrix(1, 12) = "DATE CONFIRMED"
        .TextMatrix(1, 13) = "LEAD TIME"

        'SHIPMENT(7DAYS)
        .TextMatrix(1, 14) = "DAYS"
        .TextMatrix(1, 15) = "DATE SHIPPED"
        .TextMatrix(1, 16) = "LEAD TIME"
        
        'ARRIVAL(28DAYS)
        .TextMatrix(1, 17) = "DAYS"
        .TextMatrix(1, 18) = "DATE RECEIVED"
        .TextMatrix(1, 19) = "LEAD TIME"
        .TextMatrix(1, 20) = "AVE. LEAD TIME"
        
        'SERVICE RATIO
        .TextMatrix(1, 21) = "QTY ORDERED"
        .TextMatrix(1, 22) = "QTY CONFIRMED"
        .TextMatrix(1, 23) = "% CONFIRMED"
        .TextMatrix(1, 24) = "QTY SERVED"
        .TextMatrix(1, 25) = "% SERVED"
        .TextMatrix(1, 26) = "BACK ORDER"
        
        .MergeCells = flexMergeFree

        'SET MERGING
        .MergeCells = 1
        .MergeRow(0) = True
        .MergeCol(0) = True
        .MergeCol(1) = True
        
        For i = 1 To 26
            .MergeCol(i) = False
        Next i
        
        For i = 1 To 26
            .ColAlignment(i) = 4
        Next i
        
        .ColWidth(0) = 300
        'PURCHASE ORDER DETAILS
        .ColWidth(1) = (.Width - 300) * 0.07            'PO DATE
        .ColWidth(2) = (.Width - 300) * 0.2             'SUPPLIER
        .ColWidth(3) = (.Width - 300) * 0.1             'PO NO
        .ColWidth(4) = (.Width - 300) * 0.08            'ORDER TYPE
        .ColWidth(5) = (.Width - 300) * 0.06            'TOTAL PN
        .ColWidth(6) = (.Width - 300) * 0.07            'TOTAL ORDER
        .ColWidth(7) = (.Width - 300) * 0.08            'TOTAL AMOUNT
        .ColWidth(8) = (.Width - 300) * 0.08            'CONFIRM NO
        .ColWidth(9) = (.Width - 300) * 0.15            'PURCHASED BY
        .ColWidth(10) = (.Width - 300) * 0.07           'DATE FAX
        'CONFIRMATION(7DAYS)
        .ColWidth(11) = (.Width - 300) * 0.06           'DAYS
        .ColWidth(12) = (.Width - 300) * 0.1            'DATE CONFIRMED
        .ColWidth(13) = (.Width - 300) * 0.06           'LEAD TIME
        'SHIPMENT(7DAYS)
        .ColWidth(14) = (.Width - 300) * 0.06           'DAYS
        .ColWidth(15) = (.Width - 300) * 0.07           'DATE SHIPPED
        .ColWidth(16) = (.Width - 300) * 0.06           'LEAD TIME
        'ARRIVAL(28DAYS)
        .ColWidth(17) = (.Width - 300) * 0.06           'DAYS
        .ColWidth(18) = (.Width - 300) * 0.1            'DATE RECEIVED
        .ColWidth(19) = (.Width - 300) * 0.06           'LEAD TIME
        .ColWidth(20) = (.Width - 300) * 0.1            'AVE. LEAD TIME
        'SERVICE RATIO
        .ColWidth(21) = (.Width - 300) * 0.07           'QTY ORDERED
        .ColWidth(22) = (.Width - 300) * 0.08           'QTY CONFIRMED
        .ColWidth(23) = (.Width - 300) * 0.07           '% CONFIRMED
        .ColWidth(24) = (.Width - 300) * 0.07           'QTY SERVED
        .ColWidth(25) = (.Width - 300) * 0.06           '% SERVED
        .ColWidth(26) = (.Width - 300) * 0.07           'BACK ORDER

        .ColWidth(27) = 0 '(.Width - 300) * 0.07           'SPORDERID
        .ColWidth(28) = 0 '(.Width - 300) * 0.06           'PRINT COPY
        .ColWidth(29) = 0 '(.Width - 300) * 0.07           'REMARKS
    End With

End Sub

Private Sub tmrAutoLoad_Timer()
On Error GoTo er

    Dim Qry As String, ShowDateRange As Boolean, i As Integer

    If tmrAutoLoad.Enabled = False Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    tmrAutoLoad.Enabled = False
    
    Dto.Value = Date
    Dfrom.Value = Date
    cmdNext.Tag = "1"
    
    Shape2.FillColor = RGB(173, 255, 47)
    Shape3.FillColor = RGB(135, 206, 250)
    Shape5.FillColor = RGB(255, 165, 0)
    Shape4.FillColor = RGB(211, 211, 211)
    Shape7.FillColor = RGB(255, 255, 255)
    Shape6.FillColor = RGB(250, 250, 180)
    Shape8.FillColor = RGB(240, 0, 0)
    
'    With rsSupplier
'        If .State = adStateOpen Then .Close
'        Qry = "SELECT SUPPLIER.SUPPLIERNAME, SUPPLIER.SUPPLIERID " & _
'              "From SUPPLIER ORDER BY SUPPLIER.SUPPLIERNAME;"
'        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
'        If .RecordCount > 0 Then
'            .MoveFirst
'            cboSupplier.Clear
'            Do While Not .EOF
'                cboSupplier.AddItem !SUPPLIERNAME
'                .MoveNext
'            Loop
'        End If
'    End With
    
    cboPO.Clear
    BtnIndx = 1
    tmrFloat.Enabled = True
    Warehouse.sBar.Panels(6).Text = "Spare Parts Order Summary"
    
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

Private Sub tmrDrop_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 8 Then     'last index + 1
            tmrDrop.Enabled = False
        End If
    End If
End Sub

Private Sub tmrDropRev_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 2 Then
            tmrDropRev.Enabled = False
        End If
    End If
End Sub

Private Sub tmrFloat_Timer()
    Dim i As Integer
    
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top - 50
    If cmdTrans(BtnIndx).Top <= -375 Then
        cmdTrans(BtnIndx).Visible = False
        BtnIndx = BtnIndx - 1
        If BtnIndx = -1 Then
            tmrFloat.Enabled = False
            BtnIndx = 2 'Index of Button view grid
            For i = 2 To 7
                cmdTrans(i).Top = -375
                cmdTrans(i).Left = IIf(i = 2, 0, cmdTrans(2).Width * (i - 2))
                cmdTrans(i).Visible = True
            Next i
            tmrDrop.Enabled = True
        End If
    End If
End Sub

Private Sub tmrFloatRev_Timer()
    
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top - 50
    If cmdTrans(BtnIndx).Top <= -375 Then
        cmdTrans(BtnIndx).Visible = False
        BtnIndx = BtnIndx - 1
        If BtnIndx = 1 Then 'Index of Button Exit
            tmrFloatRev.Enabled = False
            BtnIndx = 0
            cmdTrans(0).Visible = True
            cmdTrans(1).Visible = True
            tmrDropRev.Enabled = True
        End If
    End If
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
    Dfrom.Enabled = True
    Screen.MousePointer = vbDefault
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    Call modObjTools.ErrorMessage(Err.Description, "Error->tmrRefresh_Timer()")
End Sub

Private Sub txtFind_Change()
    fRow = 2
    fCol = 1
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub txtQTYCONFIRM_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtQTYCONFIRM, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then
        flxOrder.TextMatrix(OSrow, 22) = Val(txtQTYCONFIRM.Text)
        txtQTYCONFIRM.Visible = False
    End If
End Sub

Private Sub txtQTYCONFIRM_LostFocus()

If Val(flxOrder.TextMatrix(OSrow, 22)) > 0 Or Val(txtQTYCONFIRM.Text) > 0 Then
    ans = MsgBox("Do you want to save QUANTITY CONFIRMATION?", vbQuestion + vbYesNo, "VALIDATION")
    
    If ans = vbYes Then
        LYRcon.Execute "UPDATE SPORDER SET QTYCONFIRM = " & Val(txtQTYCONFIRM.Text) & " Where SPORDERID = " & Val(flxOrder.TextMatrix(OSrow, 27)) & ""
        Call cmdRefresh_Click
        MsgBox "You have successfully encoded Quantity Confirmation!", vbInformation + vbOKOnly, "VALIDATION"
    Else
        Exit Sub
    End If
End If
End Sub
