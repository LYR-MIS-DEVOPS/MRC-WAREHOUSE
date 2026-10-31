VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form SPPOSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Existing Spare Parts"
   ClientHeight    =   10680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18690
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SPPOSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10680
   ScaleWidth      =   18690
   Begin VB.Frame Frame1 
      Caption         =   "Branch Database Update"
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   120
      TabIndex        =   58
      Top             =   8760
      Width           =   18495
      Begin VB.Label lblBrnCon2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Suggested Order Quantity Formula"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   6480
         TabIndex        =   60
         Top             =   480
         Width           =   2460
      End
      Begin VB.Label lblBrnCon 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Suggested Order Quantity Formula"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   6540
         TabIndex        =   59
         Top             =   240
         Width           =   2460
      End
   End
   Begin TabDlg.SSTab tabProcessing 
      Height          =   975
      Left            =   6000
      TabIndex        =   61
      Tag             =   "xyf"
      Top             =   3600
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   1720
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Processing"
      TabPicture(0)   =   "SPPOSearch.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label28"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Image2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProcess"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Label lblProcess 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Analyzing Inventory"
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
         Left            =   720
         TabIndex        =   63
         Tag             =   "xyf"
         Top             =   600
         Width           =   3975
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   240
         Picture         =   "SPPOSearch.frx":05A6
         Tag             =   "xyf"
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label28 
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
         Left            =   960
         TabIndex        =   62
         Tag             =   "xyf"
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   8280
      Top             =   135
   End
   Begin VB.Frame fmAnotation 
      Caption         =   "EXPLANATORY NOTE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   915
      Left            =   120
      TabIndex        =   38
      Top             =   9720
      Width           =   18480
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TTLQH ="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   1
         Left            =   10320
         TabIndex        =   64
         Top             =   270
         Width           =   720
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Branch Qty. On Hand + Warehouse Qty. On Hand"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   11160
         TabIndex        =   57
         Top             =   270
         Width           =   4050
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SOQ     = 6 Mos Sales Qty. - (Branch QH+Warehouse QH)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   0
         Left            =   10320
         TabIndex        =   56
         Top             =   585
         Width           =   4695
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "(6) months Sales Qty."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   8280
         TabIndex        =   55
         Top             =   585
         Width           =   1830
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6mSQ = "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   7560
         TabIndex        =   54
         Top             =   585
         Width           =   720
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Cost W/Out Vat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   8280
         TabIndex        =   53
         Top             =   270
         Width           =   1305
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CWOV = "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   7560
         TabIndex        =   52
         Top             =   270
         Width           =   720
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Suggested Order Qty."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5520
         TabIndex        =   51
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SOQ = "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   4800
         TabIndex        =   49
         Top             =   585
         Width           =   570
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "In-Transit Qty."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5520
         TabIndex        =   48
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "InTrn = "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   4800
         TabIndex        =   47
         Top             =   270
         Width           =   690
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Quantity ON-HAND"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   645
         TabIndex        =   46
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label Label6 
         Caption         =   " Supplier BACK-ORDER"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   600
         TabIndex        =   45
         Top             =   600
         Width           =   1920
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Sales Quantity"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3135
         TabIndex        =   44
         Top             =   270
         Width           =   1230
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Quantity ON-HAND"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3135
         TabIndex        =   43
         Top             =   585
         Width           =   1545
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QH ="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   150
         TabIndex        =   42
         Top             =   270
         Width           =   420
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BO ="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   180
         TabIndex        =   41
         Top             =   585
         Width           =   405
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SQ = "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   2640
         TabIndex        =   40
         Top             =   270
         Width           =   450
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QH = "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   2640
         TabIndex        =   39
         Top             =   585
         Width           =   555
      End
   End
   Begin MSComCtl2.UpDown UpDRef 
      Height          =   315
      Left            =   1800
      TabIndex        =   35
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      Max             =   1
      Enabled         =   -1  'True
   End
   Begin MSComctlLib.ProgressBar Pbar 
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   8280
      Visible         =   0   'False
      Width           =   14145
      _ExtentX        =   24950
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtOrder 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3240
      TabIndex        =   18
      Top             =   9600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtSuppCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   17
      Top             =   9600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cboFind 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin lvButton.lvButtons_H cmdInsert 
      Height          =   375
      Left            =   14400
      TabIndex        =   7
      Top             =   8280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Insert"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Image           =   "SPPOSearch.frx":182D0
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdOptHide 
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "&Hide Option"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      ImgAlign        =   3
      Image           =   "SPPOSearch.frx":1886A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdOptShow 
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "&Show Option"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      ImgAlign        =   3
      Image           =   "SPPOSearch.frx":18E04
      cBack           =   -2147483633
   End
   Begin VB.Frame FrameOpt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   60
      TabIndex        =   3
      Top             =   720
      Width           =   18555
      Begin VB.ComboBox cboClass 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6600
         TabIndex        =   66
         Tag             =   "XYF"
         Top             =   360
         Width           =   2895
      End
      Begin VB.CheckBox chkResize 
         Height          =   345
         Left            =   7440
         Picture         =   "SPPOSearch.frx":1939E
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   720
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.ComboBox cbomos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "SPPOSearch.frx":19C68
         Left            =   6600
         List            =   "SPPOSearch.frx":19C6A
         TabIndex        =   36
         Tag             =   "XYF"
         Top             =   720
         Width           =   735
      End
      Begin MSComCtl2.UpDown OptrSQ 
         Height          =   285
         Left            =   17400
         TabIndex        =   34
         Top             =   255
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   2
         Max             =   5
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown OptrBO 
         Height          =   285
         Left            =   17400
         TabIndex        =   33
         Top             =   615
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   2
         Max             =   5
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtQH 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   14895
         TabIndex        =   30
         Text            =   "0"
         Top             =   615
         Width           =   480
      End
      Begin MSComCtl2.UpDown updBO 
         Height          =   285
         Left            =   18120
         TabIndex        =   26
         Top             =   615
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtBO"
         BuddyDispid     =   196643
         OrigLeft        =   9840
         OrigTop         =   600
         OrigRight       =   10095
         OrigBottom      =   855
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtBO 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   17640
         TabIndex        =   25
         Text            =   "0"
         Top             =   615
         Width           =   480
      End
      Begin MSComCtl2.UpDown updSQ 
         Height          =   285
         Left            =   18120
         TabIndex        =   24
         Top             =   255
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtSQ"
         BuddyDispid     =   196644
         OrigLeft        =   11280
         OrigTop         =   255
         OrigRight       =   11535
         OrigBottom      =   540
         Max             =   1000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtSQ 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   17640
         TabIndex        =   23
         Text            =   "0"
         Top             =   255
         Width           =   480
      End
      Begin MSComCtl2.DTPicker SDfrom 
         Height          =   315
         Left            =   1650
         TabIndex        =   14
         Tag             =   "XYF"
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   94109697
         CurrentDate     =   39498
      End
      Begin VB.CheckBox chkAllSupp 
         Caption         =   "All Suppliers:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   420
         Width           =   1335
      End
      Begin VB.ComboBox cboSupplier 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1650
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   360
         Width           =   3735
      End
      Begin MSComCtl2.DTPicker SDto 
         Height          =   315
         Left            =   3810
         TabIndex        =   16
         Tag             =   "XYF"
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   94109697
         CurrentDate     =   39498
      End
      Begin MSComCtl2.UpDown OptrQH 
         Height          =   285
         Left            =   14640
         TabIndex        =   29
         Top             =   615
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   2
         Max             =   5
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updQH 
         Height          =   285
         Left            =   15360
         TabIndex        =   31
         Top             =   615
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtQH"
         BuddyDispid     =   196642
         OrigLeft        =   9015
         OrigTop         =   255
         OrigRight       =   9270
         OrigBottom      =   540
         Max             =   1000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Order Mos.:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5700
         TabIndex        =   37
         Tag             =   "XYF"
         Top             =   720
         Width           =   915
      End
      Begin VB.Label lblOprQH 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   ">="
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   14160
         TabIndex        =   32
         Top             =   615
         Width           =   480
      End
      Begin VB.Label lblOprSQ 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   ">="
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   16920
         TabIndex        =   28
         Top             =   255
         Width           =   480
      End
      Begin VB.Label lblOprBO 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   ">="
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   16920
         TabIndex        =   27
         Top             =   615
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "BRANCH SQ:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   15840
         TabIndex        =   22
         ToolTipText     =   "SALES QTY."
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "BRANCH QH:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   15360
         TabIndex        =   21
         ToolTipText     =   "BACK ORDER"
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "WAREHOUSE QH:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   12480
         TabIndex        =   20
         ToolTipText     =   "QTY-ON-HAND"
         Top             =   630
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "To:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   15
         Tag             =   "XYF"
         Top             =   750
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Sales Date From:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Tag             =   "XYF"
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Classification:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   420
         Width           =   1095
      End
   End
   Begin lvButton.lvButtons_H cmdFind 
      Height          =   495
      Left            =   6480
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   4
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
      Image           =   "SPPOSearch.frx":19C6C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   375
      Left            =   15780
      TabIndex        =   8
      Top             =   8280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&OK"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Image           =   "SPPOSearch.frx":1A206
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   17160
      TabIndex        =   9
      Top             =   8280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Cancel"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Image           =   "SPPOSearch.frx":1A7A0
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdReplenish2 
      Height          =   375
      Left            =   16080
      TabIndex        =   65
      Top             =   240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      Caption         =   "Stocks for Replenish"
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
      Image           =   "SPPOSearch.frx":1AD3A
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxParts 
      Height          =   6360
      Left            =   60
      TabIndex        =   67
      Top             =   1800
      Width           =   18525
      _ExtentX        =   32676
      _ExtentY        =   11218
      _Version        =   393216
      Rows            =   3
      Cols            =   17
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      AllowUserResizing=   1
      FormatString    =   "|<PARTNUMBER |<DESCRIPTION |<LAST DATE-IN |^QH |^SQ |^INTRN |^QH |^BO |^INTRN |^TTLQH |^6MSOQ |^ORDER |^CWOV |^APM"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Left            =   90
      TabIndex        =   10
      Top             =   7320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblRefSearch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FIND PARTS DESC."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   195
      Width           =   1485
   End
   Begin VB.Menu mnuSP 
      Caption         =   "SP"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuAct 
         Caption         =   "set to ""ACTIVE"""
      End
      Begin VB.Menu mnuInAct 
         Caption         =   "set to ""IN-ACTIVE"""
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMark 
         Caption         =   "Mark"
      End
      Begin VB.Menu mnuUnmark 
         Caption         =   "Unmark"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMarkAll 
         Caption         =   "Mark All"
      End
      Begin VB.Menu mnuUnmarkAll 
         Caption         =   "UnMark All"
      End
   End
   Begin VB.Menu mnurg 
      Caption         =   "RG"
      Visible         =   0   'False
      Begin VB.Menu mnuresize 
         Caption         =   "Resize Grid"
      End
   End
End
Attribute VB_Name = "SPPOSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSupplier As New ADODB.Recordset
Dim ShowOpt As Boolean, SuppSrch As Boolean
Dim ActRow As Integer, ActCol As Integer, PrvW As Integer
Public MySource As String, RwSel As Boolean
Dim tm As Long
Dim MRC1Connect As Boolean, MRC2Connect As Boolean, UlasConnect As Boolean, TorilConnect As Boolean, Van2Connect As Boolean
Dim PanaboConnect As Boolean, BinosPolConnect As Boolean, YamahaConnect As Boolean, Van1Connect As Boolean, GensanConnect As Boolean
Dim YmarbelConnect As Boolean, StotomasConnect As Boolean, SamalConnect As Boolean, Van3Connect As Boolean
Dim WsaleConnect As Boolean, Whouse As Boolean, CbntianConnect As Boolean
Dim rsClass As New ADODB.Recordset


Private Sub cboClass_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
    If SuppSrch Then
        SuppSrch = False
        Exit Sub
    End If
    With cboClass
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsClass.MoveFirst
            rsClass.Move i
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                SuppSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            'txtSuppCode.Text = ""
        End If
    End With

End Sub

Private Sub cboClass_Click()
    Call cboClass_Change
End Sub

Private Sub cboClass_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboClass.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboClass_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then SuppSrch = True
End Sub

Private Sub cboClass_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then cboFind.SetFocus
End Sub


Private Sub cboFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdFind_Click
End Sub

Private Sub cbomos_Change()
    cbomos_Click
End Sub

Private Sub cbomos_Click()
Dim mos As Integer
    mos = Val(cbomos.Text)
    pmos = "-" & mos
    SDfrom.Value = DateAdd("m", pmos, Date)
   
    If cbomos.Text = "All" Then
        SDfrom.Value = SPordering.SalesDateFrom.Value
    Else
        flxParts.TextMatrix(1, 12) = cbomos.Text & "mSOQ"
    End If
End Sub

Private Sub cbomos_DblClick()
    cbomos_Click
End Sub

Private Sub cbomos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii > 54 Or KeyAscii < 48 Then
        KeyAscii = 0
    End If
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
            txtSuppCode.Text = IIf(IsNull(rsSupplier!SUPPLIERCODE), "", rsSupplier!SUPPLIERCODE)
            chkAllSupp.Value = 0
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                SuppSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtSuppCode.Text = ""
        End If
    End With

End Sub

Private Sub cboSupplier_Click()
    Call cboSupplier_Change
End Sub

Private Sub cboSupplier_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboSupplier.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then SuppSrch = True
End Sub

Private Sub cboSupplier_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then cboFind.SetFocus
End Sub

Private Sub cboSupplier_LostFocus()
    If cboSupplier.Text = "" Then txtSuppCode.Text = ""
End Sub

Private Sub chkAllSupp_Click()
    If chkAllSupp.Value = 1 Then 'All Supplier
        cboSupplier.Text = ""
        txtSuppCode.Text = ""
    Else
        If txtSuppCode.Text = "" Then
            cboSupplier.Text = SPordering.cboSupplier.Text
            txtSuppCode.Text = SPordering.txtSCODE.Text
        End If
    End If
End Sub

Private Sub chkResize_Click()
    If chkResize.Value = 1 Then
        Timer1.Enabled = True
        tm = 0
    Else
        Timer1.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub DisableButtonOnProcess(Bln As Boolean)

    If Bln Then
        If ShowOpt = True Then
            cmdOptHide.Enabled = False
            FrameOpt.Enabled = False
        Else
            cmdOptShow.Enabled = False
        End If
        cmdEvaluate.Enabled = False
        cmdReplenish.Enabled = False
        cmdInsert.Enabled = False
        cmdOK.Enabled = False
        cmdCancel.Enabled = False
    Else
        If ShowOpt = True Then
            cmdOptHide.Enabled = True
            FrameOpt.Enabled = True
        Else
            cmdOptShow.Enabled = True
        End If
        cmdEvaluate.Enabled = True
        cmdReplenish.Enabled = True
        cmdInsert.Enabled = True
        cmdOK.Enabled = True
        cmdCancel.Enabled = True
    End If
End Sub

Private Sub cmdFind_Click()
    Dim rsSP As New ADODB.Recordset
    
    Dim MySP() As String, Mdl() As String, Qry As String
    Dim myW As String, myW1 As String, myW2 As String, myW3 As String
    Dim MySel As String, i As Long, r As Long, rM As Integer
    Dim QHGray As Boolean, BOGray As Boolean, SQGray As Boolean
    
    myW1 = ""
    If cboFind.Text <> "" Then
        MySP = Split(cboFind.Text, " ")
        For i = LBound(MySP) To UBound(MySP)
            If myW1 = "" Then
                'myW1 = "(SPAREPARTS.DESCRIPTION) Like '%" & MySP(i) & "%'"
                
                If updRef.Value = 0 Then
                    myW1 = "PARTNUMBER Like \'%" & MySP(i) & "%\'"
                Else
                    myW1 = "DESCRIPTION Like \'%" & MySP(i) & "%\'"
                End If
            Else
                'myW1 = myW1 & " AND (SPAREPARTS.DESCRIPTION) Like '%" & MySP(i) & "%'"
                If updRef.Value = 0 Then
                    myW1 = myW1 & " AND PARTNUMBER Like \'%" & MySP(i) & "%\'"
                Else
                    myW1 = myW1 & " AND DESCRIPTION Like \'%" & MySP(i) & "%\'"
                End If
            End If
        Next i
        myW = myW1
    End If
    If txtSuppCode.Text <> "" And myW <> "" Then
        myW = myW & " AND SUPPLIERCODE = \'" & txtSuppCode.Text & "\'"
    ElseIf txtSuppCode.Text <> "" And myW = "" Then
        myW = "SUPPLIERCODE = \'" & txtSuppCode.Text & "\'"
    ElseIf txtSuppCode.Text = "" And myW = "" Then
        Exit Sub
    End If
    If cboClass.ListIndex >= 0 Then myW = myW & " AND REMARKS = \'" & cboClass.Text & "\'"
    DoEvents
    tabProcessing.Visible = True
    flxSP.Clear
    flxSP.Rows = 1
    Call ResizeCompleteHead
    Call CollAlign
    Call DisableButtonOnProcess(True)
    Qry = "CALL mrcwarehouse.sp_pobyclass('" & myW & "', '" & Format(SDfrom.Value, "yyyy-mm-dd") & "', '" & Format(SDto.Value, "yyyy-mm-dd") & "');"
    'InputBox Qry, Qry, Qry
    Set rsSP = New ADODB.Recordset
    If rsSP.State = adStateOpen Then rsSP.Close
    With rsSP
        DoEvents
        .Open Qry, LYRcon, adOpenStatic
        If .RecordCount = 1 Then
            .MoveFirst
            myW2 = .Fields(0)
            myW3 = .Fields(1)
            If myW2 = "Error" Then
                MsgBox myW3, vbCritical + vbOKOnly, "DB ERROR"
                .Close
                flxParts.Clear
                flxParts.Rows = 1
                tabProcessing.Visible = False
                Call DisableButtonOnProcess(False)
                Exit Sub
            End If
        End If
        If .RecordCount > 0 Then
            flxParts.Rows = 1
            pBar.Value = 0
            pBar.Visible = True
            pBar.Max = .RecordCount + 1
            .MoveFirst
            For rM = 1 To .RecordCount
                QHGray = False
                spSOQ = 0
                flxParts.Rows = rM
                flxParts.TextMatrix(rM - 1, 1) = !PartNumber
                flxParts.TextMatrix(rM - 1, 2) = !Description
                flxParts.TextMatrix(rM - 1, 3) = !LASTDATEIN
                flxParts.TextMatrix(rM - 1, 4) = !B_QH
                flxParts.TextMatrix(rM - 1, 5) = !B_SQ
                flxParts.TextMatrix(rM - 1, 6) = !B_ST
                flxParts.TextMatrix(rM - 1, 7) = !Ware_QH
                flxParts.TextMatrix(rM - 1, 8) = !WARE_BO
                flxParts.TextMatrix(rM - 1, 9) = !WARE_ST 'NEW: IN-TRANSIT    old:6MSQ
                flxParts.TextMatrix(rM - 1, 10) = Val(flxSP.TextMatrix(rM - 1, 4)) + Val(flxSP.TextMatrix(rM - 1, 7))
                spSOQ = Val(flxSP.TextMatrix(rM - 1, 5)) - (Val(flxSP.TextMatrix(rM - 1, 10)) + Val(flxSP.TextMatrix(rM - 1, 9)) + Val(flxSP.TextMatrix(rM - 1, 8)))
                flxParts.TextMatrix(rM - 1, 11) = IIf(spSOQ < 0, 0, Int(spSOQ + 0.49))
                flxParts.TextMatrix(rM - 1, 13) = Format(!Cost, "#,##0.00")
                flxParts.TextMatrix(rM - 1, 14) = IIf(IsNull(!ApplicModel), "", !ApplicModel)
                flxParts.TextMatrix(rM - 1, 15) = !SPAREPARTID
                If QHGray And BOGray And SQGray Then
                    flxParts.Col = 0: flxParts.row = rM - 1: flxParts.ColSel = 79: flxParts.RowSel = rM - 1: flxParts.FillStyle = flexFillRepeat: flxParts.CellBackColor = RGB(255, 255, 255)
                End If
                DoEvents
                pBar.Value = pBar.Value + 1
                .MoveNext
            Next rM
            
            flxSP.Visible = True
            If ShowOpt Then
                'flxSP.Top = 1920
                flxSPH.Top = FrameOpt.Top + FrameOpt.Height
                flxSP.Top = FrameOpt.Top + FrameOpt.Height + flxSPH.Height
            Else
                'flxSP.Top = 1080
                flxSPH.Top = cmdOptHide.Top + cmdOptHide.Height + 100
                flxSP.Top = FrameOpt.Height + flxSPH.Height
            End If
            cmdInsert.Top = flxSP.Top + flxSP.Height + 120
            cmdOK.Top = flxSP.Top + flxSP.Height + 120
            cmdCancel.Top = flxSP.Top + flxSP.Height + 120
            pBar.Top = flxSP.Top + flxSP.Height
           ' Me.Height = fmAnotation.Height + cmdOK.Top + cmdOK.Height + 580
        Else
            flxSP.Clear
            flxSP.Rows = 1
        End If
        .Close
    End With
    
    If flxSP.Rows = 1 And flxSP.TextMatrix(0, 1) = "" Then Exit Sub
    Call AddInCombo
    pBar.Visible = False
    'Assign color light green to MRC Warehouse
    flxSP.Col = 71: flxSP.row = 0: flxSP.ColSel = 73: flxSP.RowSel = flxSP.Rows - 1: flxSP.FillStyle = flexFillRepeat: flxSP.CellBackColor = RGB(150, 230, 150)
    If flxSP.Rows >= 1 Then
        'If flxSP.TextMatrix(1, 3) <> "" Then
                         'QTY on Hand
        If flxSP.TextMatrix(0, 67) <> "" Then
            flxSP.SetFocus
        Else
            cboFind.SetFocus
        End If
    Else
        cboFind.SetFocus
    End If
    tabProcessing.Visible = False
    Call DisableButtonOnProcess(False)

 End Sub

Function GetSalesQty(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    'old Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE (((SPITR.TRANSDATE)>=#" & SDfrom.Value & "# And (SPITR.TRANSDATE)<=#" & SDto.Value & "#) AND ((SPITR.TRANSTYPE)='OUT') AND ((SPITR.SPAREPARTID)=" & SDto.Value & "));"
    'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE (((SPITR.TRANSDATE)>=#" & SDfrom.Value & "# And (SPITR.TRANSDATE)<=#" & SDto.Value & "#) AND ((SPITR.TRANSTYPE)='OUT') AND ((SPITR.REMARKS)<>'TRANSFER') AND ((SPITR.SPAREPARTID)=" & MyID & "));"
    'mySQL
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
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
    'Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE (((SPITR.SPAREPARTID)=" & MyID & ") AND ((SPITR.TRANSTYPE)='SALES RETURN') AND ((SPITR.TRANSDATE)>=#" & SDfrom.Value & "# And (SPITR.TRANSDATE)<=#" & SDto.Value & "#));"
    'mySQL
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValue(Qry))
    GetSalesQty = MySales - MyRet
End Function

Function GetSalesQty6mos(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
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
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValue(Qry))
    GetSalesQty6mos = MySales - MyRet
End Function

Function GetSalesQty6MBranch(MyID As Long, myCON As String) As Long
    Dim rsQty As New ADODB.Recordset
    Dim MyQ As String
    Dim MySales As Long, MyRet As Long
    
    MyQ = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        If myCON = "VAN1" Then
            .Open MyQ, conVAN1, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "VAN2" Then
            .Open MyQ, conVAN2, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "MRC2" Then
            .Open MyQ, conMRC2, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "MRC1" Then
            .Open MyQ, conMRC1, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "ULAS" Then
            .Open MyQ, conUlas, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "TORIL" Then
            .Open MyQ, conToril, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "YAMAHA" Then
            .Open MyQ, conYM, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "PANABO" Then
            .Open MyQ, conPanabo, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "POLOMOLOK" Then
            .Open MyQ, conPol, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "GENSAN" Then
            .Open MyQ, conGenSan, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "GANTE" Then
            .Open MyQ, conMRC2, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "MARBEL" Then
            .Open MyQ, conYmarbel, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "STOTOMAS" Then
            .Open MyQ, conStotomas, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "SAMAL" Then
            .Open MyQ, conSamal, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "VAN3" Then
            .Open MyQ, conVan3, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "WSALE" Then
            .Open MyQ, conWsale, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "CABANTIAN" Then
            .Open MyQ, conCabantian, adOpenKeyset, adLockOptimistic
        End If
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    MyQ = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueBranch(MyQ, myCON))
    GetSalesQty6MBranch = MySales - MyRet
End Function


Function GetSalesQty6mosMU(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, BrchConMU, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueMU(Qry))
    GetSalesQty6mosMU = MySales - MyRet
End Function

Function GetSalesQty6mosYM(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, BrchConYM, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueYM(Qry))
    GetSalesQty6mosYM = MySales - MyRet
End Function

Function GetSalesQty6mosMD(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, BrchConMD, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueMD(Qry))
    GetSalesQty6mosMD = MySales - MyRet
End Function

Function GetSalesQty6mosVAN(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, BrchConVAN, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueVAN(Qry))
    GetSalesQty6mosVAN = MySales - MyRet
End Function


Function GetSalesQty6mosMRC2(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, BrnConMRC2, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueMRC2(Qry))
    GetSalesQty6mosMRC2 = MySales - MyRet
End Function



Function GetSalesQtyMain(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
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
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValue(Qry))
    GetSalesQtyMain = MySales - MyRet
End Function
Private Sub AddInCombo()
    Dim x As Integer, W As Boolean
    
    If cboFind.Text = "" Then Exit Sub
    W = False
    For x = 0 To cboFind.ListCount - 1
        If cboFind.Text = cboFind.List(x) Then
            If x = 0 Then
                W = True
                Exit For
            Else
                cboFind.RemoveItem x
                Exit For
            End If
        End If
    Next x
    If W = False Then cboFind.AddItem cboFind.Text, 0
End Sub

Private Sub cmdInsert_Click()
    Dim i As Integer
    
    
    If MySource = "SPORDERING" Then
        With SPordering.flxParts
            If .TextMatrix(.Rows - 1, 1) = "" Then  'Remove last row if empty
                If .Rows = 2 Then
                    .Rows = 1
                Else
                    .RemoveItem .Rows - 1
                End If
            End If
            For i = 0 To flxSP.Rows - 1
                If flxSP.TextMatrix(i, 0) <> "" Then
                    If Not SPinFlx(flxSP.TextMatrix(i, 1)) Then
                        '.AddItem .Rows & vbTab & flxSP.TextMatrix(i, 1) & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 8) & vbTab & flxSP.TextMatrix(i, 7) & vbTab & vbTab & flxSP.TextMatrix(i, 9) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & .TextMatrix(i, 13)
                        'Add applicable model
                        '                        1-PARTNUMBER                     2-DESCRIPTION                    3-ORDER                          4-COST                                   6-TOTAL                                                                          13-APPLICABLE MODEL
                        '.AddItem .Rows & vbTab & flxSP.TextMatrix(i, 1) & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 8) & vbTab & flxSP.TextMatrix(i, 7) & vbTab & vbTab & flxSP.TextMatrix(i, 9) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & flxSP.TextMatrix(i, 10)
                        '                        1-PARTNUMBER                     2-DESCRIPTION                    3-ORDER                          4-COST                       6-TOTAL                   10-spid                                                                        12-APPLICABLE MODEL
                        '.AddItem .Rows & vbTab & flxSP.TextMatrix(i, 1) & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 9) & vbTab & flxSP.TextMatrix(i, 11) & vbTab & vbTab & flxSP.TextMatrix(i, 10) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & flxSP.TextMatrix(i, 12)
                        'Design for  BrANCH     1-PARTNUMBER                       2-DESCRIPTION                  3-ORDER                          4-cost                       5-Total                    6-Spid                                                                        12-APPLICABLE MODEL
                        '.AddItem .Rows & vbTab & flxSP.TextMatrix(i, 1) & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 12) & vbTab & flxSP.TextMatrix(i, 14) & vbTab & vbTab & flxSP.TextMatrix(i, 13) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & flxSP.TextMatrix(i, 15)
                        'Design for  Main     1-PARTNUMBER                       2-DESCRIPTION                  3-ORDER                          4-cost                       5-Total                    6-Spid                                                                        12-APPLICABLE MODEL
                        .AddItem .Rows & vbTab & flxSP.TextMatrix(i, 1) & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 76) & vbTab & flxSP.TextMatrix(i, 77) & vbTab & vbTab & flxSP.TextMatrix(i, 79) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & flxSP.TextMatrix(i, 78)
             
                    End If
                End If
            Next i
            If .Rows = 1 Then .AddItem "1"
        End With
    End If
    If MySource = "RECEIVING" Then
        'Unused code because we have separated program for receiving
        'ith SPincoming.flxOrd
        '    If .TextMatrix(.Rows - 1, 7) = "" And .TextMatrix(.Rows - 1, 2) = "" Then 'Remove last row if empty
                
        '        If .Rows = 2 Then
        '            .Rows = 1
        '        Else
        '             .RemoveItem .Rows - 1
        '        End If
        '    End If
        '    For i = 1 To flxSP.Rows - 1
        '        If flxSP.TextMatrix(i, 0) <> "" Then
        '            If Not SPinFlx(flxSP.TextMatrix(i, 9)) Then
                        '.AddItem vbTab & vbTab & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 3) & vbTab & flxSP.TextMatrix(i, 5) & vbTab & vbTab & flxSP.TextMatrix(i, 9) & vbTab & flxSP.TextMatrix(i, 1)
                        '
        '                .AddItem vbTab & vbTab & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 3) & vbTab & flxSP.TextMatrix(i, 5) & vbTab & vbTab & flxSP.TextMatrix(i, 9) & vbTab & flxSP.TextMatrix(i, 1)
        '
        '            End If
        '        End If
        '    Next i
        '    If .Rows = 1 Then .Rows = 2
        'End With
    End If
End Sub
Function SPinFlx(myPN As String) As Boolean
    Dim i As Integer
    
    If MySource = "SPORDERING" Then
        With SPordering.flxParts
            For i = 1 To .Rows - 1
                If myPN = .TextMatrix(i, 1) Then
                    SPinFlx = True
                    Exit Function
                End If
            Next i
        End With
        SPinFlx = False
        Exit Function
    End If
    'Unused code for receiving
    'If MySource = "RECEIVING" Then
    '    With SPincoming.flxOrd
    '        For i = 1 To .Rows - 1
    '            If MyID = .TextMatrix(i, 7) Then
    '                SPinFlx = True
    '                Exit Function
    '            End If
    '        Next i
    '    End With
    '    SPinFlx = False
    '    Exit Function
    'End If
End Function

Private Sub cmdOK_Click()
    Call cmdInsert_Click
    Call SPordering.ComputeTotal
    Unload Me
End Sub

Private Sub cmdOptHide_Click()
    FrameOpt.Visible = False
    cmdOptHide.Left = cmdOptShow.Left
    cmdOptHide.Visible = False
    cmdOptShow.Visible = True
    flxParts.Top = cmdOptHide.Top + cmdOptHide.Height + 100
    cmdInsert.Top = flxParts.Top + flxParts.Height + 120
    cmdOK.Top = flxParts.Top + flxParts.Height + 120
    cmdCancel.Top = flxParts.Top + flxParts.Height + 120
    pBar.Top = flxParts.Top + flxParts.Height
    'Me.Height = fmAnotation.Height + cmdOK.Top + cmdOK.Height + 880 + Frame1.Height
    Frame1.Top = cmdOK.Top + cmdOK.Height + 100
    fmAnotation.Top = Frame1.Top + Frame1.Height + 120

    'txtModel.Text = ""
    ShowOpt = False
End Sub

Private Sub cmdOptShow_Click()
    FrameOpt.Visible = True
    cmdOptHide.Left = cmdOptShow.Left
    cmdOptHide.Visible = True
    cmdOptShow.Visible = False
    flxParts.Top = FrameOpt.Top + FrameOpt.Height
    cmdInsert.Top = flxParts.Top + flxParts.Height + 120
    cmdOK.Top = flxParts.Top + flxParts.Height + 120
    cmdCancel.Top = flxParts.Top + flxParts.Height + 120
    pBar.Top = flxParts.Top + flxParts.Height
'    Me.Height = fmAnotation.Height + cmdOK.Top + cmdOK.Height + 580
'    fmAnotation.Top = cmdOK.Top + cmdOK.Height + 900 - fmAnotation.Height + 120
    'Me.Height = fmAnotation.Height + cmdOK.Top + cmdOK.Height + 880 + Frame1.Height
    Frame1.Top = cmdOK.Top + cmdOK.Height + 100
    fmAnotation.Top = Frame1.Top + Frame1.Height + 120
    
    ShowOpt = True
End Sub

Private Sub cmdReplenish2_Click()
    Dim rsRpnsh As New ADODB.Recordset
    Dim Qry As String, r As Long, spSOQ As Double, rM As Integer
    Dim myW As String
    Dim QHGray As Boolean, BOGray As Boolean, SQGray As Boolean

    cmdReplenish2.Enabled = False
    cmdInsert.Enabled = False
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    
    'identify conditions from Supplier and Classification
    If Trim(cboClass.Text) <> "" Then
        myW = "REMARKS = \'" & cboClass.Text & "\'"
    Else
        myW = "REMARKS is NOT NULL"
    End If
    If txtSuppCode.Text <> "" Then
        myW = myW & " AND SUPPLIERCODE = \'" & txtSuppCode.Text & "\'"
    End If
    
    DoEvents
    tabProcessing.Visible = True
    flxSP.Clear
    flxSP.Rows = 1
    Call ResizeCompleteHead
    Call CollAlign
    Call DisableButtonOnProcess(True)
    Qry = "CALL mrcwarehouse.sp_pobyclass('" & myW & "', '" & Format(SDfrom.Value, "yyyy-mm-dd") & "', '" & Format(SDto.Value, "yyyy-mm-dd") & "');"
    Set rsSP = New ADODB.Recordset
    If rsSP.State = adStateOpen Then rsSP.Close
    With rsSP
        DoEvents
        .Open Qry, LYRcon, adOpenStatic
        If .RecordCount = 1 Then
            .MoveFirst
            myW2 = .Fields(0)
            myW3 = .Fields(1)
            If myW2 = "Error" Then
                MsgBox myW3, vbCritical + vbOKOnly, "DB ERROR"
                .Close
                flxSP.Clear
                flxSP.Rows = 1
                Exit Sub
            End If
        End If
        If .RecordCount > 0 Then
            flxSP.Rows = 1
            pBar.Value = 0
            pBar.Visible = True
            pBar.Max = .RecordCount + 1
            .MoveFirst
            For rM = 1 To .RecordCount
                flxSP.Rows = rM
                flxSP.TextMatrix(rM - 1, 1) = !PartNumber
                flxSP.TextMatrix(rM - 1, 2) = !Description
                flxSP.TextMatrix(rM - 1, 3) = !Ulas_QH
                flxSP.TextMatrix(rM - 1, 4) = !Ulas_BO
                flxSP.TextMatrix(rM - 1, 5) = !Ulas_SQ
                flxSP.TextMatrix(rM - 1, 6) = !Ulas_ST
                flxSP.TextMatrix(rM - 1, 7) = !Rizal_QH
                flxSP.TextMatrix(rM - 1, 8) = !Rizal_BO
                flxSP.TextMatrix(rM - 1, 9) = !Rizal_SQ
                flxSP.TextMatrix(rM - 1, 10) = !Rizal_ST
                flxSP.TextMatrix(rM - 1, 11) = !TORIL_QH
                flxSP.TextMatrix(rM - 1, 12) = !Toril_BO
                flxSP.TextMatrix(rM - 1, 13) = !TORIL_SQ
                flxSP.TextMatrix(rM - 1, 14) = !TORIL_ST
                flxSP.TextMatrix(rM - 1, 15) = !VAN2_QH
                flxSP.TextMatrix(rM - 1, 16) = !Van2_BO
                flxSP.TextMatrix(rM - 1, 17) = !VAN2_SQ
                flxSP.TextMatrix(rM - 1, 18) = !VAN2_ST
                flxSP.TextMatrix(rM - 1, 19) = !VAN56_QH
                flxSP.TextMatrix(rM - 1, 20) = !VAN56_BO
                flxSP.TextMatrix(rM - 1, 21) = !VAN56_SQ
                flxSP.TextMatrix(rM - 1, 22) = !VAN56_ST
                flxSP.TextMatrix(rM - 1, 23) = !Panabo_QH
                flxSP.TextMatrix(rM - 1, 24) = !Panabo_BO
                flxSP.TextMatrix(rM - 1, 25) = !Panabo_SQ
                flxSP.TextMatrix(rM - 1, 26) = !Panabo_ST
                flxSP.TextMatrix(rM - 1, 31) = !Sob_QH
                flxSP.TextMatrix(rM - 1, 32) = !Sob_BO
                flxSP.TextMatrix(rM - 1, 33) = !Sob_SQ
                flxSP.TextMatrix(rM - 1, 34) = !Sob_ST
                flxSP.TextMatrix(rM - 1, 35) = !Gensan_QH
                flxSP.TextMatrix(rM - 1, 36) = !Gensan_BO
                flxSP.TextMatrix(rM - 1, 37) = !Gensan_SQ
                flxSP.TextMatrix(rM - 1, 38) = !Gensan_ST
                flxSP.TextMatrix(rM - 1, 39) = !Gante_QH
                flxSP.TextMatrix(rM - 1, 40) = !Gante_BO
                flxSP.TextMatrix(rM - 1, 41) = !Gante_SQ
                flxSP.TextMatrix(rM - 1, 42) = !Gante_ST
                flxSP.TextMatrix(rM - 1, 55) = !VAN55_QH
                flxSP.TextMatrix(rM - 1, 56) = !VAN55_BO
                flxSP.TextMatrix(rM - 1, 57) = !VAN55_SQ
                flxSP.TextMatrix(rM - 1, 58) = !VAN55_ST
                flxSP.TextMatrix(rM - 1, 59) = !WS_QH
                flxSP.TextMatrix(rM - 1, 60) = !WS_BO
                flxSP.TextMatrix(rM - 1, 61) = !WS_SQ
                flxSP.TextMatrix(rM - 1, 62) = !WS_ST
                flxSP.TextMatrix(rM - 1, 63) = !CAB_QH
                flxSP.TextMatrix(rM - 1, 64) = !CAB_BO
                flxSP.TextMatrix(rM - 1, 65) = !CAB_SQ
                flxSP.TextMatrix(rM - 1, 66) = !CAB_ST
                flxSP.TextMatrix(rM - 1, 71) = !Ware_QH
                flxSP.TextMatrix(rM - 1, 72) = !WARE_BO
                flxSP.TextMatrix(rM - 1, 73) = !WARE_ST 'NEW: IN-TRANSIT    old:6MSQ
                flxSP.TextMatrix(rM - 1, 77) = Format(!Cost, "#,##0.00")
                flxSP.TextMatrix(rM - 1, 78) = IIf(IsNull(!ApplicModel), "", !ApplicModel)
                flxSP.TextMatrix(rM - 1, 79) = !SPAREPARTID
                flxSP.TextMatrix(rM - 1, 80) = IIf(IsNull(!DESCRIP2), "", !DESCRIP2)
                QHGray = False
                spSOQ = 0
                If (lblOprQH.Caption = "=" And !Ware_QH = Val(txtQH.Text)) Or _
                   (lblOprQH.Caption = ">" And !Ware_QH > Val(txtQH.Text)) Or _
                   (lblOprQH.Caption = ">=" And !Ware_QH >= Val(txtQH.Text)) Or _
                   (lblOprQH.Caption = "<" And !Ware_QH >= Val(txtQH.Text)) Or _
                   (lblOprQH.Caption = "<=" And !Ware_QH <= Val(txtQH.Text)) Or _
                   (lblOprQH.Caption = "<>" And !Ware_QH <> Val(txtQH.Text)) Then QHGray = True
                BOGray = False
                flxSP.TextMatrix(rM - 1, 67) = !CAB_QH + !WS_QH + !VAN55_QH + !Gante_QH + !Gensan_QH + !Sob_QH + !Panabo_QH + !VAN56_QH + !VAN2_QH + !TORIL_QH + !Rizal_QH + !Ulas_QH
                If (lblOprBO.Caption = "=" And flxSP.TextMatrix(rM - 1, 67) = Val(txtBO.Text)) Or _
                    (lblOprBO.Caption = ">" And flxSP.TextMatrix(rM - 1, 67) > Val(txtBO.Text)) Or _
                    (lblOprBO.Caption = ">=" And flxSP.TextMatrix(rM - 1, 67) >= Val(txtBO.Text)) Or _
                    (lblOprBO.Caption = "<" And flxSP.TextMatrix(rM - 1, 67) < Val(txtBO.Text)) Or _
                    (lblOprBO.Caption = "<=" And flxSP.TextMatrix(rM - 1, 67) <= Val(txtBO.Text)) Or _
                    (lblOprBO.Caption = "<>" And flxSP.TextMatrix(rM - 1, 67) <> Val(txtBO.Text)) Then BOGray = True
                SQGray = False
                flxSP.TextMatrix(rM - 1, 69) = !CAB_SQ + !WS_SQ + !VAN55_SQ + !Gante_SQ + !Gensan_SQ + !Sob_SQ + !Panabo_SQ + !VAN56_SQ + !VAN2_SQ + !TORIL_SQ + !Rizal_SQ + !Ulas_SQ
                If (lblOprSQ.Caption = "=" And flxSP.TextMatrix(rM - 1, 69) = Val(txtSQ.Text)) Or _
                    (lblOprSQ.Caption = ">" And flxSP.TextMatrix(rM - 1, 69) > Val(txtSQ.Text)) Or _
                    (lblOprSQ.Caption = ">=" And flxSP.TextMatrix(rM - 1, 69) >= Val(txtSQ.Text)) Or _
                    (lblOprSQ.Caption = "<" And flxSP.TextMatrix(rM - 1, 69) < Val(txtSQ.Text)) Or _
                    (lblOprSQ.Caption = "<=" And flxSP.TextMatrix(rM - 1, 69) <= Val(txtSQ.Text)) Or _
                    (lblOprSQ.Caption = "<>" And flxSP.TextMatrix(rM - 1, 69) <> Val(txtSQ.Text)) Then SQGray = True
                
                flxSP.TextMatrix(rM - 1, 68) = !CAB_BO + !WS_BO + !VAN55_BO + !Gante_BO + !Gensan_BO + !Sob_BO + !Panabo_BO + !VAN56_BO + !Van2_BO + !Toril_BO + !Rizal_BO + !Ulas_BO
                flxSP.TextMatrix(rM - 1, 70) = !CAB_ST + !WS_ST + !VAN55_ST + !Gante_ST + !Gensan_ST + !Sob_ST + !Panabo_ST + !VAN56_ST + !VAN2_ST + !TORIL_ST + !Rizal_ST + !Ulas_ST
                flxSP.TextMatrix(rM - 1, 74) = Val(flxSP.TextMatrix(rM - 1, 67)) + Val(flxSP.TextMatrix(rM - 1, 71))
                spSOQ = Val(flxSP.TextMatrix(rM - 1, 69)) - (Val(flxSP.TextMatrix(rM - 1, 74)) + Val(flxSP.TextMatrix(rM - 1, 73)) + Val(flxSP.TextMatrix(rM - 1, 72)))
                flxSP.TextMatrix(rM - 1, 75) = IIf(spSOQ < 0, 0, Int(spSOQ + 0.49))
                If QHGray And BOGray And SQGray Then
                    flxSP.Col = 0: flxSP.row = rM - 1: flxSP.ColSel = 79: flxSP.RowSel = rM - 1: flxSP.FillStyle = flexFillRepeat: flxSP.CellBackColor = RGB(255, 255, 255)
                End If
                DoEvents
                pBar.Value = pBar.Value + 1
                .MoveNext
            Next rM
            
            flxSP.Visible = True
            If ShowOpt Then
                'flxSP.Top = 1920
                flxSPH.Top = FrameOpt.Top + FrameOpt.Height
                flxSP.Top = FrameOpt.Top + FrameOpt.Height + flxSPH.Height
            Else
                'flxSP.Top = 1080
                flxSPH.Top = cmdOptHide.Top + cmdOptHide.Height + 100
                flxSP.Top = FrameOpt.Height + flxSPH.Height
            End If
            cmdInsert.Top = flxSP.Top + flxSP.Height + 120
            cmdOK.Top = flxSP.Top + flxSP.Height + 120
            cmdCancel.Top = flxSP.Top + flxSP.Height + 120
            pBar.Top = flxSP.Top + flxSP.Height
           ' Me.Height = fmAnotation.Height + cmdOK.Top + cmdOK.Height + 580
        Else
            flxSP.Clear
            flxSP.Rows = 1
        End If
        .Close
    End With
    
    If flxSP.Rows = 1 And flxSP.TextMatrix(0, 1) = "" Then Exit Sub
    Call AddInCombo
    pBar.Visible = False
    flxSP.Col = 71: flxSP.row = 0: flxSP.ColSel = 73: flxSP.RowSel = flxSP.Rows - 1: flxSP.FillStyle = flexFillRepeat: flxSP.CellBackColor = RGB(150, 230, 150)
    If flxSP.Rows >= 1 Then
        'If flxSP.TextMatrix(1, 3) <> "" Then
                         'QTY on Hand
        If flxSP.TextMatrix(0, 67) <> "" Then
            flxSP.SetFocus
        Else
            cboFind.SetFocus
        End If
    Else
        cboFind.SetFocus
    End If
    tabProcessing.Visible = False
    cmdReplenish2.Enabled = True
    Call DisableButtonOnProcess(False)


End Sub





Private Sub Form_Load()
    Dim Qry As String, m As Integer
    
    RwSel = True
    Call cmdOptHide_Click
    With SPordering
        If .LockFlx = "True" Then
            cmdInsert.Enabled = False
            cmdOK.Enabled = False
        End If
        'SDfrom.Value = .SalesDateFrom.Value
        SDfrom.Value = DateAdd("m", -6, Date)
        SDto.Value = .SalesDateTo.Value
        cbomos.Text = 6
        For m = 1 To 12
            cbomos.AddItem m
        Next m
        'cboMos.AddItem "All"
    End With
    
    'Set Search Description
    updRef.Value = 1
    
    Call ResizeCompleteHead
   
    
    Timer1.Enabled = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'LYRmain.sBar.Panels(6).Text = ""
    'MainCon.Close
    'BrchConMU.Close
    'BrchConYM.Close
    'BrchConMD.Close
    'BrchConVAN.Close
    'BrnConMRC2.Close
    'Timer1.Enabled = False
    If rsSupplier.State = adStateOpen Then rsSupplier.Close
End Sub

Private Sub mnuAct_Click()
    Dim ans As Integer
    
    ans = MsgBox("Do you want ACTIVATE Partnumber " & flxSP.TextMatrix(flxSP.row, 1) & "?", vbQuestion + vbYesNo, "In-Active")
    If ans = vbNo Then Exit Sub
    
    flxSP.Col = 1:  flxSP.ColSel = 80: flxSP.RowSel = flxSP.row: flxSP.FillStyle = flexFillRepeat: flxSP.CellForeColor = RGB(0, 0, 0)
    flxSP.TextMatrix(flxSP.row, 80) = ""
    
    'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION2 = '' WHERE SPAREPARTS.SPAREPARTID=" & Val(flxSP.TextMatrix(flxSP.Row, 9)) & ";"
    'mySQL
    LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION2 = '' WHERE SPAREPARTS.SPAREPARTID=" & Val(flxSP.TextMatrix(flxSP.row, 79)) & ";"

End Sub

Private Sub mnuInAct_Click()
    Dim ans As Integer
    
    ans = MsgBox("Do you want Partnumber " & flxSP.TextMatrix(flxSP.row, 1) & " to be IN-ACTIVE?", vbQuestion + vbYesNo, "In-Active")
    If ans = vbNo Then Exit Sub
    
    flxSP.Col = 1:  flxSP.ColSel = 80: flxSP.RowSel = flxSP.row: flxSP.FillStyle = flexFillRepeat: flxSP.CellForeColor = RGB(255, 0, 0)
    flxSP.TextMatrix(flxSP.row, 80) = "INACTIVE"
    
    'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION2 = 'INACTIVE' WHERE (((SPAREPARTS.SPAREPARTID)=" & Val(flxSP.TextMatrix(flxSP.Row, 9)) & "));"
    'mySQL
    LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION2 = 'INACTIVE' WHERE SPAREPARTS.SPAREPARTID=" & Val(flxSP.TextMatrix(flxSP.row, 79)) & ";"
    
End Sub

Private Sub mnuMark_Click()
    With flxSP
        .Col = 0
        .CellFontName = "WINGDINGS"
        .CellFontSize = 12
        .Text = lblChk.Caption
        .CellAlignment = 4
    End With
End Sub

Private Sub mnuMarkAll_Click()
    Dim i As Integer
    
    With flxSP
        For i = 1 To .Rows - 1
            .row = i
            .Col = 0
            .CellFontName = "WINGDINGS"
            .CellFontSize = 12
            .Text = lblChk.Caption
            .CellAlignment = 4
            .TextMatrix(i, 0) = lblChk.Caption
        Next i
    End With
    
End Sub

Private Sub mnuUnmark_Click()
    flxSP.TextMatrix(flxSP.row, 0) = ""
End Sub

Private Sub mnuUnmarkAll_Click()
    Dim i As Integer
    
    For i = 1 To flxSP.Rows - 1
        flxSP.TextMatrix(i, 0) = ""
    Next i
End Sub


Private Sub OptrBO_Change()
    Select Case OptrBO.Value
        Case 0
            lblOprBO.Caption = "="
        Case 1
            lblOprBO.Caption = ">"
        Case 2
            lblOprBO.Caption = ">="
        Case 3
            lblOprBO.Caption = "<"
        Case 4
            lblOprBO.Caption = "<="
        Case 5
            lblOprBO.Caption = "<>"
    End Select

End Sub
Private Sub OptrQH_Change()
    Select Case OptrQH.Value
        Case 0
            lblOprQH.Caption = "="
        Case 1
            lblOprQH.Caption = ">"
        Case 2
            lblOprQH.Caption = ">="
        Case 3
            lblOprQH.Caption = "<"
        Case 4
            lblOprQH.Caption = "<="
        Case 5
            lblOprQH.Caption = "<>"
    End Select
End Sub
Private Sub OptrSQ_Change()
    Select Case OptrSQ.Value
        Case 0
            lblOprSQ.Caption = "="
        Case 1
            lblOprSQ.Caption = ">"
        Case 2
            lblOprSQ.Caption = ">="
        Case 3
            lblOprSQ.Caption = "<"
        Case 4
            lblOprSQ.Caption = "<="
        Case 5
            lblOprSQ.Caption = "<>"
    End Select
End Sub

Private Sub Timer1_Timer()
    Dim rsBackup As New ADODB.Recordset
    Dim Qry As String
    
    Qry = "SELECT SUPPLIER.SUPPLIERID, SUPPLIER.SUPPLIERNAME, SUPPLIER.SUPPLIERCODE " & _
          "From SUPPLIER ORDER BY SUPPLIER.SUPPLIERNAME;"
    With rsSupplier
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboSupplier.Clear
            Do While Not .EOF
                cboSupplier.AddItem !SUPPLIERNAME
                .MoveNext
            Loop
        End If
    End With
    
    cboSupplier.Text = SPordering.cboSupplier.Text
    txtSuppCode.Text = SPordering.txtSCODE.Text
   
   
    Qry = "SELECT DISTINCT spareparts.REMARKS From spareparts Order By spareparts.REMARKS"
    If rsClass.State = adStateOpen Then rsClass.Close
    With rsClass
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboClass.Clear
            Do While Not .EOF
                cboClass.AddItem !Remarks
                .MoveNext
            Loop
        End If
    End With

    lblBrnCon.Caption = ""
    lblBrnCon2.Caption = ""
    Qry = "CALL sp_mrc_backup_monitoring.sp_mrc_backup_monitoring;"
    If rsBackup.State = adStateOpen Then rsBackup.Close
    With rsBackup
        DoEvents
        .Open Qry, LYRcon, adOpenStatic
        If .RecordCount > 1 Then
            .MoveFirst
            Do While Not .EOF
                If !BranchCode = "MST" Then
                    lblBrnCon.Caption = IIf(lblBrnCon.Caption <> "", lblBrnCon.Caption & ", ", "") & "LYRMP SOBRECARY:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MRT" Then
                    lblBrnCon.Caption = IIf(lblBrnCon.Caption <> "", lblBrnCon.Caption & ", ", "") & "LYRMP GANTE:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "M2W" Then
                    lblBrnCon.Caption = IIf(lblBrnCon.Caption <> "", lblBrnCon.Caption & ", ", "") & "LYRMP WHOLESALE:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MUD" Then
                    lblBrnCon.Caption = IIf(lblBrnCon.Caption <> "", lblBrnCon.Caption & ", ", "") & "LYRMPULAS:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MPT" Then
                    lblBrnCon.Caption = IIf(lblBrnCon.Caption <> "", lblBrnCon.Caption & ", ", "") & "MCP DEPOT TORIL:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MCPR" Then
                    lblBrnCon.Caption = IIf(lblBrnCon.Caption <> "", lblBrnCon.Caption & ", ", "") & "MCP DEPOT RIZAL:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MG" Then
                    lblBrnCon2.Caption = IIf(lblBrnCon2.Caption <> "", lblBrnCon2.Caption & ", ", "") & "LYRMP LAGAO:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MCB" Then
                    lblBrnCon2.Caption = IIf(lblBrnCon2.Caption <> "", lblBrnCon2.Caption & ", ", "") & "LYRMP CABANTIAN:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MP" Then
                    lblBrnCon2.Caption = IIf(lblBrnCon2.Caption <> "", lblBrnCon2.Caption & ", ", "") & "LYRMP PANABO:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MV1" Then
                    lblBrnCon2.Caption = IIf(lblBrnCon2.Caption <> "", lblBrnCon2.Caption & ", ", "") & "VAN-2 [LHC-403]:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MV3" Then
                    lblBrnCon2.Caption = IIf(lblBrnCon2.Caption <> "", lblBrnCon2.Caption & ", ", "") & "VAN-55 [AAH-1250]:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MV2" Then
                    lblBrnCon2.Caption = IIf(lblBrnCon2.Caption <> "", lblBrnCon2.Caption & ", ", "") & "VAN-56 [AAH-1233]:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                End If
                .MoveNext
            Loop
        End If
        .Close
    End With
        

'    If flxSP.Rows = 1 And flxSP.TextMatrix(0, 1) = "" Then Exit Sub
'
'    For Cw = 0 To 24
'        flxSP.ColWidth(Cw) = flxSPH.ColWidth(Cw)
'    Next Cw
'    fScrol.Left = flxSPH.ColPos(6) + 60
'    flxSP.ColWidth(26) = flxSPH.ColWidth(25)
'    flxSP.ColWidth(27) = flxSPH.ColWidth(26)
'    tm = tm + 1
'    If tm = 60 Then chkResize.Value = 0
    Timer1.Enabled = False
End Sub
Private Sub txtOrder_GotFocus()
    'Timer1.Enabled = False
End Sub
Private Sub txtOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If Val(txtOrder.Text) > Val(flxSP.TextMatrix(ActRow, 75)) And UserID <> 2 Then
            flxSP.TextMatrix(ActRow, 76) = ""
            flxSP.TextMatrix(ActRow, 0) = ""
            txtOrder.Visible = False
            flxSP.Col = 76
            Exit Sub
        ElseIf Val(txtOrder.Text) > Val(flxSP.TextMatrix(ActRow, 75)) And UserID = 2 Then
            MsgBox ("Order Greater than Suggested Order Quantity (SOQ)"), vbExclamation, "Order"
            txtOrder.SetFocus
            Exit Sub
        ElseIf Val(txtOrder.Text) < Val(flxSP.TextMatrix(ActRow, 75)) And Val(txtOrder.Text) <> 0 Then
            MsgBox ("Order Lesser than Suggested Order Quantity (SOQ)"), vbExclamation, "Order"
            txtOrder.SetFocus
            Exit Sub
        End If
        txtOrder.Text = flxSP.TextMatrix(ActRow, 76)
        flxSP.SetFocus
        flxSP.row = ActRow
        flxSP.Col = 76
        txtOrder.Visible = False
    End If
End Sub

Private Sub txtOrder_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        flxSP.SetFocus
        flxSP.row = ActRow
        flxSP.Col = 76
   End If
End Sub
Private Sub txtOrder_LostFocus()
    With flxSP
    
        'If Val(txtOrder.Text) > Val(.TextMatrix(ActRow, 47)) And (UserID <> 4 Or UserID <> 5) Then
'        If Val(txtOrder.Text) > Val(.TextMatrix(ActRow, 47)) Then
'            MsgBox ("Can't add Order Quantity" & vbCrLf & _
'                    "which Greater than Suggested Order Quantity (SOQ)"), vbExclamation, "Order"
'            flxSP.TextMatrix(ActRow, 48) = ""
'            flxSP.TextMatrix(ActRow, 0) = ""
'            If txtOrder.Visible = True Then txtOrder.SetFocus ': SendKeys "{Home}+{End}"
'            Exit Sub
'        End If  'Else           'isumpay kay temporary lang pra kay iva
        If Val(txtOrder.Text) > Val(flxSP.TextMatrix(ActRow, 75)) Then
            MsgBox ("Order Greater than Suggested Order Quantity (SOQ)"), vbExclamation, "Order"
            If UserID = 2 Then 'Or UserID = 10
                .TextMatrix(ActRow, 76) = txtOrder.Text
            Else
                .TextMatrix(ActRow, 76) = ""
                .TextMatrix(ActRow, 0) = ""
                .Col = 76
                Exit Sub    'Requested by sir Jo 6/10/2016
            End If
        ElseIf Val(txtOrder.Text) < Val(flxSP.TextMatrix(ActRow, 75)) And Val(txtOrder.Text) <> 0 Then
            MsgBox ("Order Lesser than Suggested Order Quantity (SOQ)"), vbExclamation, "Order"
        End If
        
        .TextMatrix(ActRow, 76) = txtOrder.Text
        txtOrder.Visible = False
        If Val(.TextMatrix(ActRow, 76)) > 0 Then
            .row = ActRow
            .Col = 0
            .CellFontName = "WINGDINGS"
            .CellFontSize = 12
            .Text = lblChk.Caption
            .CellAlignment = 4
            .Col = 76
        Else
            .TextMatrix(ActRow, 76) = ""
            .TextMatrix(ActRow, 0) = ""
            .Col = 76
        End If
        
        'Timer1.Enabled = True
    End With
End Sub
Private Sub UpDRef_Change()
    If updRef.Value = 0 Then
        lblRefSearch.Caption = "FIND PARTNUMBER"
    Else
        lblRefSearch.Caption = "FIND PARTS DESC."
    End If
   
End Sub
Private Sub ResizeCompleteHead()
    Dim aCol As Integer, aCel As Integer, reAln As Integer
    Dim i As Integer
    
    With flxParts
        
        .TextMatrix(0, 1) = "S P A R E P A R T S"
        .TextMatrix(0, 2) = "S P A R E P A R T S"
        .TextMatrix(0, 3) = "S P A R E P A R T S"
             
        .TextMatrix(0, 4) = "BRANCH TOTAL"
        .TextMatrix(0, 5) = "BRANCH TOTAL"
        .TextMatrix(0, 6) = "BRANCH TOTAL"
        
        .TextMatrix(0, 7) = "MRC WAREHOUSE"
        .TextMatrix(0, 8) = "MRC WAREHOUSE"
        .TextMatrix(0, 9) = "MRC WAREHOUSE"
        
        .TextMatrix(0, 10) = "PURCHASE ORDER"
        .TextMatrix(0, 11) = "PURCHASE ORDER"
        .TextMatrix(0, 12) = "PURCHASE ORDER"
        .TextMatrix(0, 13) = "PURCHASE ORDER"
        .TextMatrix(0, 14) = "PURCHASE ORDER"
        
        'SPAREPARTS
        .TextMatrix(1, 1) = "PARTNUMBER"
        .TextMatrix(1, 2) = "DESCRIPTION"
        .TextMatrix(1, 3) = "LAST DATE-IN"

        'BRANCH TOTAL
        .TextMatrix(1, 4) = "QH"
        .TextMatrix(1, 5) = "SQ"
        .TextMatrix(1, 6) = "InTrn"
        
        'MRC WAREHOUSE
        .TextMatrix(1, 7) = "QH"
        .TextMatrix(1, 8) = "BO"
        .TextMatrix(1, 9) = "InTrn"
        
        'Purchase Order total
        .TextMatrix(1, 10) = "TTLQH"
        .TextMatrix(1, 11) = cbomos.Text & "mSOQ"
        .TextMatrix(1, 12) = "ORDER"
        .TextMatrix(1, 13) = "CWOV"
        .TextMatrix(1, 14) = "APM"
        
        .MergeCells = flexMergeFree
        
        .MergeRow(0) = True
        For i = 1 To 14
            .MergeCol(i) = True
        Next i

        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.13
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.3
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.1
        For aCol = 4 To 14
            .ColWidth(aCol) = (.Width - .ColWidth(0)) * 0.06
        Next aCol
        .ColWidth(15) = 0 ' (.Width - .ColWidth(0)) * 0.13
        .ColWidth(16) = 0 ' (.Width - .ColWidth(0)) * 0.13
    End With
End Sub

Private Sub CollAlign()
    Dim nCols As Integer
     
    flxSP.ColAlignment(1) = 2
    flxSP.ColAlignment(2) = 2
    For nCols = 3 To 67
        flxSP.ColAlignment(nCols) = 4
    Next nCols
End Sub

Private Sub FillColorGrd()
    Dim CCol As Integer, cRow As Integer
    
    flxSP.Col = 3
    flxSP.row = 0
    For CCol = 3 To 5
        flxSP.ColSel = CCol
        For cRow = 0 To flxSP.Rows - 1
            flxSP.RowSel = cRow
            flxSP.FillStyle = flexFillRepeat: flxSP.CellBackColor = &HC0FFC0
        Next cRow
    Next CCol
End Sub


