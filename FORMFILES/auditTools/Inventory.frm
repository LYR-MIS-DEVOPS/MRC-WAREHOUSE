VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form AuditInvntry 
   Caption         =   "Inventory Audit"
   ClientHeight    =   9150
   ClientLeft      =   15
   ClientTop       =   -1920
   ClientWidth     =   13035
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   13035
   Tag             =   "XYF"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar Pbar 
      Height          =   135
      Left            =   240
      TabIndex        =   65
      Tag             =   "XYF"
      Top             =   8460
      Visible         =   0   'False
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox txtAppModel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   62
      Top             =   9480
      Visible         =   0   'False
      Width           =   7035
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   135
      Left            =   5640
      TabIndex        =   61
      Top             =   3960
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   238
      _Version        =   393216
   End
   Begin TabDlg.SSTab tabDetails 
      Height          =   2655
      Left            =   8040
      TabIndex        =   36
      Top             =   9720
      Visible         =   0   'False
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   4683
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Inventory.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   120
         TabIndex        =   50
         Tag             =   "XYF"
         Top             =   1320
         Width           =   10455
         Begin VB.TextBox txtDispRef 
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
            Height          =   285
            Left            =   4440
            TabIndex        =   67
            Tag             =   "XYF"
            Top             =   840
            Width           =   1935
         End
         Begin VB.TextBox txtRemarks 
            Height          =   555
            Left            =   1080
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   54
            Tag             =   "XYF"
            Top             =   240
            Width           =   5295
         End
         Begin VB.TextBox txtTemp 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8520
            TabIndex        =   53
            Tag             =   "XYF"
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtFinal 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8520
            TabIndex        =   52
            Tag             =   "XYF"
            Top             =   525
            Width           =   1815
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8520
            TabIndex        =   51
            Tag             =   "XYF"
            Top             =   810
            Width           =   1815
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "BATCH NUMBER REFERENCE FOR DISPLAY:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1080
            TabIndex        =   66
            Tag             =   "XYF"
            Top             =   900
            Width           =   3585
         End
         Begin VB.Label lbl2 
            AutoSize        =   -1  'True
            Caption         =   "REMARKS:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   58
            Tag             =   "XYF"
            Top             =   240
            Width           =   1395
         End
         Begin VB.Label lblTemp 
            Alignment       =   1  'Right Justify
            Caption         =   "TEMPORARY RECORDS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6360
            TabIndex        =   57
            Tag             =   "XYF"
            Top             =   270
            Width           =   1980
         End
         Begin VB.Label lblFinal 
            Alignment       =   1  'Right Justify
            Caption         =   "FINAL RECORDS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6360
            TabIndex        =   56
            Tag             =   "XYF"
            Top             =   555
            Width           =   1980
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            Caption         =   "TOTAL RECORDS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6360
            TabIndex        =   55
            Tag             =   "XYF"
            Top             =   840
            Width           =   1980
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1245
         Left            =   120
         TabIndex        =   37
         Tag             =   "XYF"
         Top             =   60
         Width           =   10455
         Begin lvButton.lvButtons_H cmdCloseTab 
            Height          =   495
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
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
            Image           =   "Inventory.frx":001C
            cBack           =   -2147483633
         End
         Begin VB.TextBox txtBatchNo 
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
            Left            =   8520
            Locked          =   -1  'True
            TabIndex        =   43
            Tag             =   "XYF"
            Top             =   200
            Width           =   1815
         End
         Begin VB.TextBox txtDateStart 
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
            Left            =   8520
            Locked          =   -1  'True
            TabIndex        =   42
            Tag             =   "XYF"
            Top             =   520
            Width           =   1815
         End
         Begin VB.TextBox txtSubReport 
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
            Left            =   2400
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   41
            Tag             =   "XYF"
            Top             =   520
            Width           =   1815
         End
         Begin VB.TextBox txtStatus 
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
            Left            =   2400
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   40
            Tag             =   "XYF"
            Top             =   200
            Width           =   1815
         End
         Begin VB.TextBox txtAuditor 
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
            Left            =   2400
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   39
            Tag             =   "XYF"
            Top             =   840
            Width           =   3855
         End
         Begin VB.TextBox txtDateFin 
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
            Left            =   8520
            Locked          =   -1  'True
            TabIndex        =   38
            Tag             =   "XYF"
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "REPORT STATUS"
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
            Left            =   840
            TabIndex        =   49
            Tag             =   "XYF"
            Top             =   225
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "SUB REPORTS:"
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
            Left            =   840
            TabIndex        =   48
            Tag             =   "XYF"
            Top             =   555
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "BATCH NO."
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
            Left            =   6720
            TabIndex        =   47
            Tag             =   "XYF"
            Top             =   230
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "DATE START:"
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
            Left            =   6720
            TabIndex        =   46
            Tag             =   "XYF"
            Top             =   550
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "AUDITOR'S NAME:"
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
            Left            =   840
            TabIndex        =   45
            Tag             =   "XYF"
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "DATE FINISH:"
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
            Left            =   6840
            TabIndex        =   44
            Tag             =   "XYF"
            Top             =   870
            Width           =   1575
         End
      End
   End
   Begin VB.TextBox txtPrnQry 
      Height          =   285
      Left            =   4920
      TabIndex        =   35
      Top             =   9360
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.TextBox txtHomeID 
      Height          =   285
      Left            =   3960
      TabIndex        =   34
      Top             =   9240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSeqNo 
      Height          =   285
      Left            =   3000
      TabIndex        =   33
      Top             =   9240
      Visible         =   0   'False
      Width           =   855
   End
   Begin TabDlg.SSTab tabLoc 
      Height          =   3255
      Left            =   4080
      TabIndex        =   28
      Top             =   10080
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   5741
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "B1002032"
      TabPicture(0)   =   "Inventory.frx":05B6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdLocCancel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdLocOK"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "flxLoc"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin MSFlexGridLib.MSFlexGrid flxLoc 
         Height          =   2175
         Left            =   240
         TabIndex        =   30
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   3836
         _Version        =   393216
         Rows            =   1
         Cols            =   6
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   325
         BackColorBkg    =   16777215
         ScrollBars      =   2
         SelectionMode   =   1
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
      Begin lvButton.lvButtons_H cmdLocOK 
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   2805
         Width           =   1380
         _ExtentX        =   2434
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
         Image           =   "Inventory.frx":05D2
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdLocCancel 
         Height          =   375
         Left            =   1800
         TabIndex        =   32
         Top             =   2805
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   661
         Caption         =   "Cancel"
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
         Image           =   "Inventory.frx":0B6C
         cBack           =   -2147483633
      End
      Begin VB.Label Label2 
         Caption         =   "Please select a Spare Part:"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.TextBox txtLocation 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8640
      TabIndex        =   27
      Top             =   9600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtBatchID 
      Height          =   285
      Left            =   7800
      TabIndex        =   26
      Top             =   9600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtFind 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10200
      TabIndex        =   23
      Tag             =   "XYF"
      Top             =   8640
      Width           =   2415
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "&A"
      Height          =   375
      Left            =   9600
      TabIndex        =   18
      Top             =   10320
      Width           =   135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   2040
   End
   Begin VB.TextBox txtPartNo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3240
      TabIndex        =   17
      Top             =   9600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtQty 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4440
      TabIndex        =   16
      Top             =   9600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cboParts 
      Height          =   315
      Left            =   6480
      TabIndex        =   15
      Top             =   9600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxParts 
      Height          =   1815
      Left            =   120
      TabIndex        =   14
      Top             =   9960
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3201
      _Version        =   393216
      BackColor       =   16777215
      FixedRows       =   0
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      GridColorUnpopulated=   16777215
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtQtyTip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   3120
      TabIndex        =   13
      Text            =   "Qty. Purchased: 0"
      Top             =   8760
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txtPrnCopy 
      Height          =   285
      Left            =   5640
      TabIndex        =   12
      Top             =   9600
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid flxInv 
      Height          =   7455
      Left            =   240
      TabIndex        =   0
      Tag             =   "XYF"
      Top             =   960
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   13150
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      RowHeightMin    =   420
      BackColorBkg    =   -2147483633
      ScrollBars      =   2
      FormatString    =   "^    |<PART NUMBER      |<DESCRIPTION          |>PRICE         |<LOCATION     |^QTY.         |^ TIME         |^DATE    "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   873
      Caption         =   "&Add New"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Image           =   "Inventory.frx":1106
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   495
      Index           =   2
      Left            =   3240
      TabIndex        =   6
      Top             =   0
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   873
      Caption         =   "&Discrepancy"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Image           =   "Inventory.frx":16A0
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   495
      Index           =   3
      Left            =   4860
      TabIndex        =   7
      Top             =   0
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   873
      Caption         =   "&Mgr's Report"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Image           =   "Inventory.frx":1C3A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   495
      Index           =   5
      Left            =   8100
      TabIndex        =   8
      Top             =   0
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   873
      Caption         =   "E&xit"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Image           =   "Inventory.frx":21D4
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   495
      Index           =   6
      Left            =   9720
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   873
      Caption         =   "&Save"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Image           =   "Inventory.frx":276E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   495
      Index           =   7
      Left            =   11340
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   873
      Caption         =   "&Cancel"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Image           =   "Inventory.frx":2D08
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   495
      Index           =   1
      Left            =   1620
      TabIndex        =   11
      Top             =   0
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   873
      Caption         =   "&Edit Report"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Image           =   "Inventory.frx":32A2
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   495
      Index           =   4
      Left            =   6480
      TabIndex        =   19
      Top             =   0
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   873
      Caption         =   "S&ummary"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Image           =   "Inventory.frx":383C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   495
      Index           =   9
      Left            =   12840
      TabIndex        =   20
      Top             =   1560
      Visible         =   0   'False
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   873
      Caption         =   "&Print Preview"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Image           =   "Inventory.frx":3DD6
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   495
      Index           =   8
      Left            =   12840
      TabIndex        =   21
      Top             =   1080
      Visible         =   0   'False
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   873
      Caption         =   "&Sub Report"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Image           =   "Inventory.frx":4370
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   495
      Index           =   10
      Left            =   12840
      TabIndex        =   22
      Top             =   2040
      Visible         =   0   'False
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   873
      Caption         =   "&Main Report"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Image           =   "Inventory.frx":490A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdFind 
      Height          =   420
      Left            =   8820
      TabIndex        =   24
      Tag             =   "XYF"
      Top             =   8640
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   741
      Caption         =   "&Find:"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Image           =   "Inventory.frx":4EA4
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdInsertSP 
      Height          =   420
      Left            =   7440
      TabIndex        =   25
      Tag             =   "XYF"
      Top             =   8640
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   741
      Caption         =   "&Insert SP"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Image           =   "Inventory.frx":543E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdShowDtls 
      Height          =   420
      Left            =   6060
      TabIndex        =   59
      Tag             =   "XYF"
      Top             =   8640
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   741
      Caption         =   "&Details"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Image           =   "Inventory.frx":59D8
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdSort 
      Height          =   420
      Left            =   360
      TabIndex        =   63
      Tag             =   "XYF"
      Top             =   8640
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Image           =   "Inventory.frx":5F72
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   420
      Left            =   840
      TabIndex        =   64
      Tag             =   "XYF"
      Top             =   8640
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Image           =   "Inventory.frx":650C
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   0
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   13275
      m_MDI_Height    =   9720
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INVENTORY AUDIT"
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
      Height          =   240
      Index           =   2
      Left            =   10860
      TabIndex        =   3
      Tag             =   "TRF"
      Top             =   690
      Width           =   1755
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "MOTO>R REPUBLIC CORP."
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
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Tag             =   "TLF"
      Top             =   690
      Width           =   2895
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "MOTO>R REPUBLIC CORP."
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
      Left            =   390
      TabIndex        =   2
      Tag             =   "TLF"
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INVENTORY AUDIT"
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
      Height          =   240
      Index           =   3
      Left            =   10890
      TabIndex        =   4
      Tag             =   "TRF"
      Top             =   720
      Width           =   1755
   End
   Begin VB.Menu mnuADJ 
      Caption         =   "SpareParts"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Row         Alt+A"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "AuditInvntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsParts As New ADODB.Recordset
Dim ActRow As Integer, ActCol As Integer, fRow As Integer, fCol As Integer, LocRow As Integer
Dim AddNew As Boolean, LockFlx As Boolean, DispOnly As Boolean, PrtSrch As Boolean, PrtBln As Boolean, CustSrch As Boolean

Private Sub ShowButtons(MyType As Integer)
    Dim i As Integer
    
    If MyType = 1 Then
        For i = 0 To 5
            cmdTrans(i).Left = cmdTrans(0).Width * i
            cmdTrans(i).Visible = True
        Next i
        For i = 6 To 10
            cmdTrans(i).Visible = False
        Next i
        cmdPrint.Visible = True
        cmdSort.Visible = True
    ElseIf MyType = 2 Then 'Save Cancel
        For i = 0 To 10
            cmdTrans(i).Visible = False
        Next i
        cmdTrans(6).Left = 0
        cmdTrans(7).Left = cmdTrans(6).Width
        cmdTrans(6).Visible = True
        cmdTrans(7).Visible = True
        cmdPrint.Visible = False
        cmdSort.Visible = False
    ElseIf MyType = 3 Then 'Sub Report, Print Preview, Main Report, Summary, Exit
        For i = 0 To 10
            cmdTrans(i).Visible = False
        Next i
        For i = 0 To 2
            cmdTrans(8 + i).Top = 0
            cmdTrans(8 + i).Left = cmdTrans(8).Width * i
            cmdTrans(8 + i).Visible = True
        Next i
        cmdTrans(4).Left = cmdTrans(8).Width * 3
        cmdTrans(4).Visible = True
        cmdTrans(5).Left = cmdTrans(8).Width * 4
        cmdTrans(5).Visible = True
        cmdPrint.Visible = False
        cmdSort.Visible = False
    ElseIf MyType = 4 Then 'Sub Report, Print Preview, Main Report, Summary, Exit
        For i = 0 To 10
            cmdTrans(i).Visible = False
        Next i
        For i = 0 To 1
            cmdTrans(9 + i).Top = 0
            cmdTrans(9 + i).Left = cmdTrans(9).Width * i
            cmdTrans(9 + i).Visible = True
        Next i
        cmdTrans(4).Left = cmdTrans(9).Width * 2
        cmdTrans(4).Visible = True
        cmdTrans(5).Left = cmdTrans(9).Width * 3
        cmdTrans(5).Visible = True
        cmdPrint.Visible = False
        cmdSort.Visible = False
    End If
End Sub

Private Sub ClearEntries()
    txtStatus.Text = ""
    txtSubReport.Text = ""
    txtAuditor.Text = ""
    txtBatchNo.Text = ""
    txtDateStart.Text = ""
    txtDateFin.Text = ""
    txtFind.Text = ""
    txtRemarks.Text = ""
    txtTemp.Text = ""
    txtFinal.Text = ""
    txtTotal.Text = ""
    txtPrnCopy.Text = ""
    txtBatchID.Text = ""
    flxInv.Rows = 1
    flxInv.Rows = 2
    txtPrnQry.Text = ""
    txtHomeID.Text = ""
    txtSeqNo.Text = ""
End Sub

Private Sub LockData(Bln As Boolean)
    'Frame1.Enabled = Not Bln
    Frame2.Enabled = Not Bln
    cmdInsertSP.Enabled = Not Bln
    LockFlx = Bln
End Sub

Private Sub ShowMgrReport(MyID As Integer)
    Dim rsAInv As New ADODB.Recordset
    Dim rsAIdtl As New ADODB.Recordset
    Dim Qry As String, Tmp As Integer
    
    
    'Qry = "SELECT AUDITINVEN.* From AUDITINVEN WHERE (((AUDITINVEN.BATCHID)=" & MyID & "));"
    'mySQL
    Qry = "SELECT AUDITINVEN.* From AUDITINVEN WHERE AUDITINVEN.BATCHID=" & MyID & ";"
    
    With rsAInv
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtStatus.Text = !Status
            txtSubReport.Text = IIf(IsNull(!BatchMember), 0, !BatchMember)
            txtAuditor.Text = !Auditor
            txtBatchNo.Text = !BATCHNO
            txtDateStart.Text = !DateStart
            txtDateFin.Text = IIf(IsNull(!DateEnd), "", Format(!DateEnd, "mm/dd/yyyy"))
            txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
            txtSeqNo.Text = !SEQNO
            txtHomeID.Text = !HOMEID
            txtBatchID.Text = MyID
        End If
        .Close
    End With
    With flxInv
        .Rows = 1
        .Rows = 2
        .TextMatrix(0, 5) = "QTY [PHY]"
        .TextMatrix(0, 6) = "QTY [REC]"
        .TextMatrix(0, 7) = "VARIANCE"
    End With
    'txtPrnQry.Text = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.QTYONPC, AUDITINVDTLS.DISCREPANCY " & _
          "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((AUDITINVDTLS.DISCREPANCY) <> 0) And ((AUDITINVDTLS.BATCHID) = " & txtBatchID.Text & ")) ORDER BY AUDITINVDTLS.ROWNO;"
    'mySQL
    txtPrnQry.Text = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.QTYONPC, AUDITINVDTLS.DISCREPANCY " & _
          "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where AUDITINVDTLS.DISCREPANCY <> 0 And AUDITINVDTLS.BATCHID = " & txtBatchID.Text & " ORDER BY AUDITINVDTLS.ROWNO;"
    
    'txtPrnQry.Tag = "MGR"
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.LASTUPDATE, AUDITINVDTLS.QTYONPC, AUDITINVDTLS.DISCREPANCY, AUDITINVDTLS.SPAREPARTID " & _
          "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((AUDITINVDTLS.DISCREPANCY) <> 0) And ((AUDITINVDTLS.BATCHID) = " & txtBatchID.Text & ")) ORDER BY AUDITINVDTLS.ROWNO;"
    
    'mySQL
    txtPrnQry.Tag = "MGR"
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.LASTUPDATE, AUDITINVDTLS.QTYONPC, AUDITINVDTLS.DISCREPANCY, AUDITINVDTLS.SPAREPARTID " & _
          "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where AUDITINVDTLS.DISCREPANCY <> 0 And AUDITINVDTLS.BATCHID = " & txtBatchID.Text & " ORDER BY AUDITINVDTLS.ROWNO;"
    
    With rsAIdtl
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxInv.Rows = 1
            Do While Not .EOF
                Tmp = Tmp + IIf(IsNull(!LASTUPDATE), 1, 0)
                flxInv.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & Format(!STOREPRICE, "#,##0.00") & vbTab & IIf(IsNull(!Location), "", !Location) & vbTab & Format(!PhyQty, "#,##0") & vbTab & IIf(IsNull(!QTYONPC), "", !QTYONPC) & vbTab & IIf(IsNull(!DISCREPANCY), "", !DISCREPANCY)
                .MoveNext
            Loop
            txtTemp.Text = Format(Tmp, "#,##0") & Space(7)
            txtFinal.Text = Format(.RecordCount - Tmp, "#,##0") & Space(7)
            txtTotal.Text = Format(.RecordCount, "#,##0") & Space(7)
        Else
            txtTemp.Text = "0" & Space(7)
            txtFinal.Text = "0" & Space(7)
            txtTotal.Text = "0" & Space(7)
        End If
        .Close
    End With
End Sub

Public Sub ShowRecord(MyID As Integer)
    Dim rsAInv As New ADODB.Recordset
    Dim rsAIdtl As New ADODB.Recordset
    Dim Qry As String, Tmp As Integer
    
    'Qry = "SELECT AUDITINVEN.* From AUDITINVEN WHERE (((AUDITINVEN.BATCHID)=" & MyID & "));"
    'mySQL
    Qry = "SELECT AUDITINVEN.* From AUDITINVEN WHERE AUDITINVEN.BATCHID=" & MyID & ";"
    
    With rsAInv
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtStatus.Text = !Status
            txtSubReport.Text = IIf(IsNull(!BatchMember), 0, !BatchMember)
            txtAuditor.Text = !Auditor
            txtBatchNo.Text = !BATCHNO
            txtDateStart.Text = !DateStart
            txtDateFin.Text = IIf(IsNull(!DateEnd), "", Format(!DateEnd, "mm/dd/yyyy"))
            txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
            txtSeqNo.Text = !SEQNO
            txtHomeID.Text = !HOMEID
            txtBatchID.Text = MyID
        End If
        .Close
    End With
    With flxInv
        .Rows = 1
        .Rows = 2
        .TextMatrix(0, 5) = "QTY"
        .TextMatrix(0, 6) = "TIME"
        .TextMatrix(0, 7) = "DATE"
    End With
    'txtPrnQry.Text = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.LASTUPDATE " & _
          "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((AUDITINVDTLS.BATCHID) = " & MyID & ")) ORDER BY AUDITINVDTLS.ROWNO;"
    'mySQL
    txtPrnQry.Text = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.LASTUPDATE " & _
          "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where AUDITINVDTLS.BATCHID = " & MyID & " ORDER BY AUDITINVDTLS.ROWNO;"
    
    'txtPrnQry.Tag = ""
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.LASTUPDATE, AUDITINVDTLS.QTYONPC, AUDITINVDTLS.DISCREPANCY, AUDITINVDTLS.SPAREPARTID " & _
          "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((AUDITINVDTLS.BATCHID) = " & MyID & ")) ORDER BY AUDITINVDTLS.ROWNO;"
    'mySQL
    txtPrnQry.Tag = ""
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.LASTUPDATE, AUDITINVDTLS.QTYONPC, AUDITINVDTLS.DISCREPANCY, AUDITINVDTLS.SPAREPARTID " & _
          "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where AUDITINVDTLS.BATCHID = " & MyID & " ORDER BY AUDITINVDTLS.ROWNO;"
    
    
    With rsAIdtl
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxInv.Rows = 1
            pBar.Value = 0
            pBar.Max = .RecordCount
            pBar.Visible = True
            Do While Not .EOF
                Tmp = Tmp + IIf(IsNull(!LASTUPDATE), 1, 0)
                flxInv.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & Format(!STOREPRICE, "#,##0.00") & vbTab & IIf(IsNull(!Location), "", !Location) & vbTab & Format(!PhyQty, "#,###") & vbTab & IIf(IsNull(!LASTUPDATE), "", Format(!LASTUPDATE, "HH:MM ampm")) & vbTab & IIf(IsNull(!LASTUPDATE), "", Format(!LASTUPDATE, "mm/dd/yyyy")) & vbTab & IIf(IsNull(!LASTUPDATE), "", !LASTUPDATE) & vbTab & IIf(IsNull(!QTYONPC), "", !QTYONPC) & vbTab & IIf(IsNull(!DISCREPANCY), "", !DISCREPANCY) & vbTab & !SPAREPARTID
                .MoveNext
                pBar.Value = pBar.Value + 1
                DoEvents
            Loop
            txtTemp.Text = Format(Tmp, "#,##0") & Space(7)
            txtFinal.Text = Format(.RecordCount - Tmp, "#,##0") & Space(7)
            txtTotal.Text = Format(.RecordCount, "#,##0") & Space(7)
        Else
            txtTemp.Text = "0" & Space(7)
            txtFinal.Text = "0" & Space(7)
            txtTotal.Text = "0" & Space(7)
        End If
        .Close
    End With
    pBar.Visible = False
End Sub

Private Sub cmdA_Click()
    Call mnuAdd_Click
End Sub

Private Sub LoadSpareParts()
    Dim Qry As String
        
    Qry = "SELECT  SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.LOCATION " & _
          "From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxParts.DataSource = rsParts
    With flxParts
        .Width = flxInv.ColWidth(4)
        .ColWidth(0) = 0
        .ColWidth(1) = 0    'SPid
        .ColWidth(2) = flxInv.ColWidth(4)   'PartNo
        .ColWidth(3) = 0    'Description
        .ColWidth(4) = 0    'Store Price
        .ColWidth(5) = 0    'Location
    End With
End Sub

Private Sub cmdCloseTab_Click()
    Dim i As Integer
    
    For i = cmdTrans.LBound To cmdTrans.UBound
        cmdTrans(i).Enabled = True
    Next i
    cmdShowDtls.Enabled = True
    If Not LockFlx Then
        cmdInsertSP.Enabled = True
        flxInv.Enabled = True
    End If
    cmdFind.Enabled = True
    txtFind.Locked = False
    With tabDetails
        .Visible = False
        .Top = 12000
    End With
End Sub

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxInv
        For i = fRow To .Rows - 1
            For j = fCol To 7
                If InStr(UCase$(.TextMatrix(i, j)), txtFind.Text) > 0 Then
                    .Col = 1: .row = i
                    .ColSel = 7:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 7, i + 1, i)
                    fCol = IIf(j = 7, 1, j + 1)
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

Private Sub cmdInsertSP_Click()
    SPfind.MySource = "AUDITINVEN"
    SPfind.Show vbModal
End Sub

Private Sub cmdLocCancel_Click()
    cmdTrans(6).Enabled = True
    cmdTrans(7).Enabled = True
    flxInv.Enabled = True
    Frame2.Enabled = True
    txtFind.Locked = False
    cmdFind.Enabled = True
    cmdInsertSP.Enabled = True
    tabLoc.Top = 15000
    tabLoc.Visible = False
    flxInv.SetFocus
    flxInv.row = ActRow
    flxInv.Col = 5
End Sub

Private Sub cmdLocOK_Click()
    Dim i As Integer
        
    With flxLoc
        'Validate if Partnumber already exist
        For i = 1 To flxInv.Rows - 1
            If flxInv.TextMatrix(i, 11) = .TextMatrix(LocRow, 1) Then
                MsgBox .TextMatrix(.row, 2) & " already exist at Row " & i, vbExclamation + vbOKOnly, "Location"
                Exit Sub
            End If
        Next i
        flxInv.TextMatrix(ActRow, 11) = .TextMatrix(LocRow, 1)
        flxInv.TextMatrix(ActRow, 1) = .TextMatrix(LocRow, 2)
        flxInv.TextMatrix(ActRow, 2) = .TextMatrix(LocRow, 3)
        flxInv.TextMatrix(ActRow, 3) = .TextMatrix(LocRow, 4)
        flxInv.TextMatrix(ActRow, 4) = .TextMatrix(LocRow, 5)
    End With
    Call cmdLocCancel_Click
End Sub

Private Sub cmdPrint_Click()
    If txtPrnQry.Text = "" Then
        MsgBox "No Records to Print!", vbExclamation + vbOKOnly, "Print Failed"
        Exit Sub
    End If
    PRNmc.QryPrn = txtPrnQry.Text
    PRNmc.ParamValue1 = txtBatchNo.Text
    PRNmc.ParamValue2 = txtAuditor.Text   'cboType.Text
    PRNmc.ParamValue3 = IIf(txtStatus.Text = "FINAL", "", "Inventory Audit is still ON-PROCESS") & " " & txtRemarks.Text
    PRNmc.MySource = "AUDITPAGE1"
    PRNmc.Show

End Sub

Private Sub cmdShowDtls_Click()
    Dim i As Integer
    
    For i = cmdTrans.LBound To cmdTrans.UBound
        cmdTrans(i).Enabled = False
    Next i
    cmdShowDtls.Enabled = False
    cmdInsertSP.Enabled = False
    flxInv.Enabled = False
    cmdFind.Enabled = False
    txtFind.Locked = True
    With tabDetails
        .Top = (Me.ScaleHeight - .Height) / 2
        .Left = (Me.ScaleWidth - .Width) / 2
        .Visible = True
    End With
End Sub

Private Sub cmdSort_Click()
    Dim rsAIdtl As New ADODB.Recordset
    Dim Qry As String, Tmp As Integer
    
    If Val(txtBatchID.Text) = 0 Then
        MsgBox "No Records to Sort!", vbExclamation + vbOKOnly, "Sort Failed"
        Exit Sub
    End If
    'txtPrnQry.Text = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.LASTUPDATE " & _
          "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((AUDITINVDTLS.BATCHID) = " & Val(txtBatchID.Text) & ")) ORDER BY SPAREPARTS.PARTNUMBER;"
    'mySQL
    txtPrnQry.Text = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.LASTUPDATE " & _
          "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where AUDITINVDTLS.BATCHID = " & Val(txtBatchID.Text) & " ORDER BY SPAREPARTS.PARTNUMBER;"
    
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.LASTUPDATE, AUDITINVDTLS.QTYONPC, AUDITINVDTLS.DISCREPANCY, AUDITINVDTLS.SPAREPARTID " & _
          "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where AUDITINVDTLS.BATCHID = " & Val(txtBatchID.Text) & " ORDER BY SPAREPARTS.PARTNUMBER;"
    With rsAIdtl
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxInv.Rows = 1
            pBar.Value = 0
            pBar.Visible = True
            pBar.Max = .RecordCount
            Do While Not .EOF
                Tmp = Tmp + IIf(IsNull(!LASTUPDATE), 1, 0)
                flxInv.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & Format(!STOREPRICE, "#,##0.00") & vbTab & IIf(IsNull(!Location), "", !Location) & vbTab & Format(!PhyQty, "#,###") & vbTab & IIf(IsNull(!LASTUPDATE), "", Format(!LASTUPDATE, "HH:MM ampm")) & vbTab & IIf(IsNull(!LASTUPDATE), "", Format(!LASTUPDATE, "mm/dd/yyyy")) & vbTab & IIf(IsNull(!LASTUPDATE), "", !LASTUPDATE) & vbTab & IIf(IsNull(!QTYONPC), "", !QTYONPC) & vbTab & IIf(IsNull(!DISCREPANCY), "", !DISCREPANCY) & vbTab & !SPAREPARTID
                .MoveNext
                pBar.Value = pBar.Value + 1
                DoEvents
            Loop
            txtTemp.Text = Format(Tmp, "#,##0") & Space(7)
            txtFinal.Text = Format(.RecordCount - Tmp, "#,##0") & Space(7)
            txtTotal.Text = Format(.RecordCount, "#,##0") & Space(7)
        Else
            txtTemp.Text = "0" & Space(7)
            txtFinal.Text = "0" & Space(7)
            txtTotal.Text = "0" & Space(7)
        End If
        .Close
    End With
    pBar.Visible = False
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim rsAud As New ADODB.Recordset
    Dim ans As Integer, i As Integer, SubRptID As Integer
    Dim myMsg As String, PrnVal(0 To 12) As String, Audi As String, Qry As String, MyAcc As String
    
    Select Case Index
        Case 0  'New
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=94));"
            'mySQL
'            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=94;"
'
'            MyAcc = GetFieldValue(Qry)
'            If MyAcc <> "FULL ACCESS" Then
'                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
'                Exit Sub
'            End If
            
            Call ClearEntries
            Call LockData(False)
            Call ShowButtons(2)
            AddNew = True
            txtBatchNo.Text = GetBatchNo
            txtDateStart.Text = Date
            txtStatus.Text = "ON PROCESS"
            txtSubReport.Text = "0"
            Audi = InputBox("Please type Name of Auditor:", "In-Charge")
            If Audi = "" Then
                MsgBox "No Auditor Specified... Transaction aborted!", vbExclamation + vbOKOnly, "Cancelled"
                Call ClearEntries
                Call LockData(True)
                Call ShowButtons(1)
                Exit Sub
            End If
            txtAuditor.Text = UCase$(Audi)
            'save records to table
            rsAud.Open "AuditInven", LYRcon, adOpenKeyset, adLockOptimistic
            rsAud.AddNew
            rsAud!BATCHNO = txtBatchNo.Text
            rsAud!Auditor = txtAuditor.Text
            rsAud!DateStart = Date
            rsAud!Status = txtStatus.Text
            rsAud!SEQNO = 1
            rsAud!LOCKENTRY = False
            rsAud!UserID = UserID
            rsAud.Update
            txtBatchID.Text = rsAud!BATCHID
            rsAud.Close
        Case 1  'Edit
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=94));"
            'mySQL
'            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=94;"
'
'            MyAcc = GetFieldValue(Qry)
'            If MyAcc <> "FULL ACCESS" Then
'                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
'                Exit Sub
'            End If
            'Validation for Editing
            If txtStatus.Text = "FINAL" Then
                MsgBox "Editing is not allowed because the Report is already Final.", vbInformation + vbOKOnly, "Edit Denied"
                Exit Sub
            End If
            If Val(txtBatchID.Text) = 0 Then
                MsgBox "No Report available for editing.", vbInformation + vbOKOnly, "Edit Denied"
                Exit Sub
            End If
            'Identify if Report is Being used by Other User
            'Qry = "SELECT AUDITINVEN.LOCKENTRY From AUDITINVEN WHERE (((AUDITINVEN.BATCHID)=" & txtBatchID.Text & "));"
            
            'mySQL
            Qry = "SELECT AUDITINVEN.LOCKENTRY From AUDITINVEN WHERE AUDITINVEN.BATCHID=" & txtBatchID.Text & ";"
            
            myMsg = GetFieldValue(Qry)
            If myMsg = "True" Then  '
                'Qry = "SELECT AUDITINVEN.USERID From AUDITINVEN WHERE (((AUDITINVEN.BATCHID)=" & txtBatchID.Text & "));"
                'mySQL
                Qry = "SELECT AUDITINVEN.USERID From AUDITINVEN WHERE AUDITINVEN.BATCHID=" & txtBatchID.Text & ";"
                i = Val(GetFieldValue(Qry))
                'Qry = "SELECT [LASTNAME] & ', ' & [FIRSTNAME] & ' ' & Left$([MIDDLENAME],1) & '.' AS USERNAME, EMPLOYEES.EMPID " & _
                      "From EMPLOYEES WHERE (((EMPLOYEES.EMPID)=" & i & "));"
                'mySQL
                Qry = "SELECT concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.'))) AS USERNAME, EMPLOYEES.EMPID " & _
                      "From EMPLOYEES WHERE EMPLOYEES.EMPID=" & i & ";"
                
                Audi = GetFieldValue(Qry)
                MsgBox "Unavailable to Edit Report. Currently used by " & Audi, vbInformation + vbOKOnly, "Edit Deneid"
                Exit Sub
            End If
            'LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.LOCKENTRY = True, AUDITINVEN.USERID = " & UserID & " WHERE (((AUDITINVEN.BATCHID)=" & txtBatchID.Text & "));"
            AddNew = False
            Call LockData(False)
            Call ShowButtons(2)
            
        Case 2  'Display Discrepancy
            'validate
            If Val(txtBatchID.Text) = 0 Then
                MsgBox "No available report to process!", vbExclamation + vbOKOnly, "Inventory Audit"
                Exit Sub
            End If
            If txtStatus.Text = "FINAL" And Val(txtSubReport.Text) > 0 Then
                'Qry = "SELECT AUDITINVEN.BATCHID From AUDITINVEN " & _
                      "WHERE (((AUDITINVEN.HOMEID)=" & txtHomeID.Text & ") AND ((AUDITINVEN.SEQNO)=" & Val(txtSeqNo.Text) + 1 & "));"
                'mySQL
                Qry = "SELECT AUDITINVEN.BATCHID From AUDITINVEN " & _
                      "WHERE AUDITINVEN.HOMEID=" & txtHomeID.Text & " AND AUDITINVEN.SEQNO=" & Val(txtSeqNo.Text) + 1 & ";"
                
                SubRptID = GetFieldValue(Qry)
                Call ShowRecord(SubRptID)
                Call ShowButtons(1)
                Exit Sub
            End If
            With flxInv
                .Rows = 1
                .Rows = 2
                .TextMatrix(0, 5) = "QTY"
                .TextMatrix(0, 6) = "TIME"
                .TextMatrix(0, 7) = "DATE"
            End With
            '***Access
            'txtPrnQry.Text = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.LASTUPDATE " & _
                  "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((AUDITINVDTLS.DISCREPANCY) <> 0) And ((AUDITINVDTLS.BATCHID) = " & txtBatchID.Text & ")) ORDER BY AUDITINVDTLS.ROWNO;"
            'mySQL
            txtPrnQry.Text = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.LASTUPDATE " & _
                  "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where AUDITINVDTLS.DISCREPANCY <> 0 And AUDITINVDTLS.BATCHID = " & txtBatchID.Text & " ORDER BY AUDITINVDTLS.ROWNO;"
            
            '***Access
            'txtPrnQry.Tag = ""
            'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.LASTUPDATE, AUDITINVDTLS.QTYONPC, AUDITINVDTLS.DISCREPANCY, AUDITINVDTLS.SPAREPARTID " & _
                  "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((AUDITINVDTLS.DISCREPANCY) <> 0) And ((AUDITINVDTLS.BATCHID) = " & txtBatchID.Text & ")) ORDER BY AUDITINVDTLS.ROWNO;"
            
            txtPrnQry.Tag = ""
            Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.LASTUPDATE, AUDITINVDTLS.QTYONPC, AUDITINVDTLS.DISCREPANCY, AUDITINVDTLS.SPAREPARTID " & _
                  "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where AUDITINVDTLS.DISCREPANCY <> 0 And AUDITINVDTLS.BATCHID = " & txtBatchID.Text & " ORDER BY AUDITINVDTLS.ROWNO;"
            
            With rsAud
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    flxInv.Rows = 1
                    .MoveFirst
                    Do While Not .EOF
                        flxInv.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & Format(!STOREPRICE, "#,##0.00") & vbTab & IIf(IsNull(!Location), "", !Location) & vbTab & Format(!PhyQty, "#,##0") & vbTab & IIf(IsNull(!LASTUPDATE), "", Format(!LASTUPDATE, "HH:MM ampm")) & vbTab & IIf(IsNull(!LASTUPDATE), "", Format(!LASTUPDATE, "mm/dd/yyyy")) & vbTab & IIf(IsNull(!LASTUPDATE), "", !LASTUPDATE) & vbTab & IIf(IsNull(!QTYONPC), "", !QTYONPC) & vbTab & IIf(IsNull(!DISCREPANCY), "", !DISCREPANCY) & vbTab & !SPAREPARTID
                        .MoveNext
                    Loop
                    txtTemp.Text = "0" & Space(7)
                    txtFinal.Text = Format(.RecordCount, "#,###") & Space(7)
                    txtTotal.Text = Format(.RecordCount, "#,###") & Space(7)
                Else
                    flxInv.Rows = 1
                    flxInv.Rows = 2
                    txtTemp.Text = ""
                    txtFinal.Text = ""
                    txtTotal.Text = ""
                End If
                .Close
            End With
            Call ShowButtons(3)
        Case 3  'Managers Report
            'validate
            If Val(txtBatchID.Text) = 0 Then
                MsgBox "No available report to process!", vbExclamation + vbOKOnly, "Inventory Audit"
                Exit Sub
            End If
            'Prompt for Password
            Password.MySource = "AUDITINV"
            Password.Show vbModal
            If cmdTrans(3).Tag <> "" Then
                cmdTrans(3).Tag = ""
                Exit Sub
            End If
            
            'Qry = "SELECT AUDITINVEN.BATCHID From AuditInven Where (((AuditInven.HOMEID) = " & txtHomeID.Text & ")) ORDER BY AUDITINVEN.SEQNO DESC;"
            'mySQL
            Qry = "SELECT AUDITINVEN.BATCHID From AuditInven Where AuditInven.HOMEID = " & txtHomeID.Text & " ORDER BY AUDITINVEN.SEQNO DESC;"
            
            SubRptID = Val(GetFieldValue(Qry))
            Call ShowMgrReport(SubRptID)
            Call ShowButtons(4)
            
        Case 4  'Summary
            AuditInvSum.Show
            Unload Me
        
        Case 5  'Exit
            'MotoRep.tbMain.Visible = True
            ans = MsgBox("Exit Audit Tools!", vbQuestion + vbYesNo, "Exit")
            If ans = vbNo Then Exit Sub
            Unload Me
            
        Case 6 'Save
            myMsg = ValidateInvAdj
            If myMsg <> "" Then
                MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            ans = MsgBox("Do you want to Save Adjustment?", vbQuestion + vbYesNo, "Save")
            If ans = vbNo Then Exit Sub
            Call SaveAuditInv
            MsgBox "Inventory report Successfully Saved!", vbInformation + vbOKOnly, "Save"
            
            Call LockData(True)
            Call ShowButtons(1)
        Case 7  'Cancel
            ans = MsgBox("Do you want to Cancel transaction?", vbQuestion + vbYesNo, "Cancel")
            If ans = vbNo Then Exit Sub
            If AddNew Then
                'remove record
                '***Access
                'LYRcon.Execute "DELETE AUDITINVEN.* From AUDITINVEN WHERE (((AUDITINVEN.BATCHID)=" & Val(txtBatchID.Text) & "));"
                'mySQL
                LYRcon.Execute "DELETE From AUDITINVEN WHERE AUDITINVEN.BATCHID=" & Val(txtBatchID.Text) & ";"
                '****Access
                LYRcon.Execute "DELETE AUDITINVDTLS.* From AUDITINVDTLS WHERE (((AUDITINVDTLS.BATCHID)=" & Val(txtBatchID.Text) & "));"
                'mySQL
                LYRcon.Execute "DELETE From AUDITINVDTLS WHERE AUDITINVDTLS.BATCHID=" & Val(txtBatchID.Text) & ";"
                Call ClearEntries
            Else
                Call ShowRecord(txtBatchID.Text)
                '***Access
                'LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.LOCKENTRY = False WHERE (((AUDITINVEN.BATCHID)=" & txtBatchID.Text & "));"
                'mySQL
                LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.LOCKENTRY = False WHERE AUDITINVEN.BATCHID=" & txtBatchID.Text & ";"
                
            End If
            Call ShowButtons(1)
            Call LockData(True)
        Case 8  'Sub Report
            '****Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=94));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=94;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            'Check if there are existing sub-reports and display if none create new
            If flxInv.Rows = 2 And flxInv.TextMatrix(1, 11) = "" Then
                MsgBox "No Discrepancies available!", vbInformation + vbOKOnly, "Sub-Report"
                Exit Sub
            End If
            ans = MsgBox("Do you want to Create Sub-Report?", vbQuestion + vbYesNo, "Sub-Report")
            If ans = vbNo Then Exit Sub
            'Update Home's Batch Members
            
            '***Access
            'LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.BATCHMEMBER = [AUDITINVEN].[BATCHMEMBER]+1 WHERE (((AUDITINVEN.BATCHID)=" & txtHomeID.Text & "));"
            'mySQL
            LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.BATCHMEMBER = AUDITINVEN.BATCHMEMBER+1 WHERE AUDITINVEN.BATCHID=" & txtHomeID.Text & ";"
            
            If Val(txtBatchID.Text) <> Val(txtHomeID.Text) Then LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.BATCHMEMBER = AUDITINVEN.BATCHMEMBER+1 WHERE AUDITINVEN.BATCHID=" & txtBatchID.Text & ";"
            'Finalize Previous Batch
            '***Access
            'Qry = "SELECT AUDITINVEN.STATUS From AUDITINVEN WHERE (((AUDITINVEN.STATUS)='FINAL') AND ((AUDITINVEN.BATCHID)=" & txtBatchID.Text & "));"
            'mySQL
            Qry = "SELECT AUDITINVEN.STATUS From AUDITINVEN WHERE AUDITINVEN.STATUS='FINAL' AND AUDITINVEN.BATCHID=" & txtBatchID.Text & ";"
            
            
            If CtrlNumExist(Qry) Then   'if Prev Batch already final update only Discrepancy else include Status and Date Finish
                LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.DISCREPANCY = " & flxInv.Rows - 1 & " WHERE AUDITINVEN.BATCHID=" & txtBatchID.Text & ";"
            Else
                'LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.STATUS = 'FINAL', AUDITINVEN.DATEEND = #" & Date & "#, AUDITINVEN.DISCREPANCY = " & flxInv.Rows - 1 & _
                " WHERE (((AUDITINVEN.BATCHID)=" & txtBatchID.Text & "));"
                'mySQL
                LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.STATUS = 'FINAL', AUDITINVEN.DATEEND = '" & Format(Date, "yyyy-mm-dd") & "', AUDITINVEN.DISCREPANCY = " & flxInv.Rows - 1 & _
                " WHERE AUDITINVEN.BATCHID=" & txtBatchID.Text & ";"
            
            End If
            
            '***Access
            'Qry = "SELECT AUDITINVEN.BATCHNO From AUDITINVEN WHERE (((AUDITINVEN.BATCHID)=" & txtHomeID.Text & "));"
            'mySQL
            Qry = "SELECT AUDITINVEN.BATCHNO From AUDITINVEN WHERE AUDITINVEN.BATCHID=" & txtHomeID.Text & ";"
            
            myMsg = GetFieldValue(Qry)
            txtBatchNo.Text = myMsg & "-" & ConvertToRoman(txtSeqNo.Text)
            txtDateStart.Text = Date
            txtStatus.Text = "ON PROCESS"
            txtSubReport.Text = "0"
            txtSeqNo.Text = Val(txtSeqNo.Text) + 1
            'Create New Batch
            
            '***Access
            'rsAud.Open "AuditInven", LYRcon, adOpenKeyset, adLockOptimistic
            'rsAud.AddNew
            'rsAud!BatchNo = txtBatchNo.Text
            'rsAud!Auditor = txtAuditor.Text
            'rsAud!DateStart = Date
            'rsAud!Status = txtStatus.Text
            'rsAud!SeqNo = Val(txtSeqNo.Text)
            'rsAud!HOMEID = txtHomeID.Text
            'rsAud!LockEntry = True
            'rsAud!UserID = UserID
            'rsAud.Update
            'txtBatchID.Text = rsAud!BATCHID
            'rsAud.Close
            
            'mySQL
            LYRcon.Execute "INSERT INTO auditinven(BATCHNO,AUDITOR,DATESTART,`STATUS`,SEQNO,HOMEID,LOCKENTRY,USERID)" & _
                "VALUES('" & txtBatchNo.Text & "','" & txtAuditor.Text & "','" & Format(Date, "yyyy-mm-dd") & "'," & _
                "'" & txtStatus.Text & "'," & Val(txtSeqNo.Text) & "," & txtHomeID.Text & ",1," & UserID & ");"
            Qry = "SELECT auditinven.BATCHID,auditinven.BATCHNO From auditinven Where auditinven.BATCHNO = '" & txtBatchNo.Text & "';"
            txtBatchID.Text = GetFieldValue(Qry)
            
            Call LockData(False)
            Call ShowButtons(2)
            AddNew = False
        Case 9  'Print Preview
            If txtPrnQry.Tag = "" Then
                PRNmc.QryPrn = txtPrnQry.Text
                PRNmc.ParamValue1 = txtBatchNo.Text
                PRNmc.ParamValue2 = txtAuditor.Text   'cboType.Text
                PRNmc.ParamValue3 = IIf(txtStatus.Text = "FINAL", "", "Inventory Audit is still ON-PROCESS") & " " & txtRemarks.Text
                PRNmc.MySource = "AUDITINVEN"
            Else
                PRNmc.QryPrn = txtPrnQry.Text
                PRNmc.ParamValue1 = txtBatchNo.Text
                PRNmc.ParamValue2 = txtAuditor.Text   'cboType.Text
                
                'Qry = "SELECT SETTINGS.VALUE From SETTINGS WHERE (((SETTINGS.REMARKS)='MANAGER') AND ((SETTINGS.DESCRIPTION)='EMP DEFAULT VALUE'));"
                'mySQL
                Qry = "SELECT SETTINGS.VALUE From SETTINGS WHERE SETTINGS.REMARKS='MANAGER' AND SETTINGS.DESCRIPTION='EMP DEFAULT VALUE';"
                
                'Audi = GetFieldValue(Qry)
                PRNmc.ParamValue3 = Audi
                PRNmc.ParamValue4 = IIf(txtStatus.Text = "FINAL", "", "Inventory Audit is still ON-PROCESS") & " " & txtRemarks.Text
                ans = MsgBox("Do you want to print report with Cost?", vbQuestion + vbYesNo, "Sub-Report")
                If ans = vbYes Then
                    PRNmc.MySource = "AUDINVMGRCOST"
                Else
                    PRNmc.MySource = "AUDINVMGR"
                End If
            End If
            PRNmc.Show
            
        Case 10 'Main Report
            'Display Main Report
            Call ShowButtons(1)
            Call ShowRecord(txtHomeID.Text)
    End Select
End Sub
Private Sub SaveAuditInv()
    Dim i As Integer, Tmp As Integer
    Dim eMsg As String
    
                
    On Error GoTo LocalErr
    With flxInv
        If AddNew Then
            Tmp = 0
            For i = 1 To .Rows - 1
                If Val(flxInv.TextMatrix(i, 11)) > 0 Then
                    If flxInv.TextMatrix(i, 8) = "" Then    'Record Not yet updated
                        eMsg = "Error 589: Adding New Record in Table Audit Detailed Inventory"
                        LYRcon.Execute "INSERT INTO AUDITINVDTLS ( BATCHID, SPAREPARTID, PHYQTY, ROWNO ) " & _
                            "VALUES (" & txtBatchID.Text & ", " & flxInv.TextMatrix(i, 11) & ", " & Val(Replace(flxInv.TextMatrix(i, 5), ",", "")) & ", " & i & ");"
                        Tmp = Tmp + 1
                    End If
                End If
            Next i
            eMsg = "Error 596: Updating Table Audit Inventory Batch ID"
            'LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.HOMEID = " & txtBatchID.Text & " WHERE (((AUDITINVEN.BATCHID)=" & txtBatchID.Text & "));"
            'mySQL
            LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.HOMEID = " & txtBatchID.Text & " WHERE AUDITINVEN.BATCHID=" & txtBatchID.Text & ";"
            
        Else
            eMsg = "Error 599: Removing Record in Table Audit Detailed Inventory"
            'LYRcon.Execute "DELETE AUDITINVDTLS.* From AUDITINVDTLS WHERE (((AUDITINVDTLS.BATCHID)=" & txtBatchID.Text & "));"
            'mySQL
            LYRcon.Execute "DELETE From AUDITINVDTLS WHERE AUDITINVDTLS.BATCHID=" & txtBatchID.Text & ";"
            
            For i = 1 To .Rows - 1
                If Val(flxInv.TextMatrix(i, 11)) > 0 Then
                    If flxInv.TextMatrix(i, 8) = "" Then    'Record Not yet updated
                        eMsg = "Error 604: Adding New Record in Table Audit Detailed Inventory"
                        LYRcon.Execute "INSERT INTO AUDITINVDTLS ( BATCHID, SPAREPARTID, PHYQTY, QTYONPC, DISCREPANCY, ROWNO ) " & _
                                "VALUES (" & txtBatchID.Text & ", " & flxInv.TextMatrix(i, 11) & ", " & Val(Replace(flxInv.TextMatrix(i, 5), ",", "")) & ", " & Val(flxInv.TextMatrix(i, 9)) & ", " & Val(flxInv.TextMatrix(i, 10)) & ", " & i & ");"
                        Tmp = Tmp + 1
                    Else
                        eMsg = "Error 609: Adding New Record in Table Audit Detailed Inventory"
                        'LYRcon.Execute "INSERT INTO AUDITINVDTLS ( BATCHID, SPAREPARTID, PHYQTY, LASTUPDATE, QTYONPC, DISCREPANCY, ROWNO ) " & _
                                "VALUES (" & txtBatchID.Text & ", " & flxInv.TextMatrix(i, 11) & ", " & Val(Replace(flxInv.TextMatrix(i, 5), ",", "")) & ", #" & flxInv.TextMatrix(i, 8) & "#, " & Val(flxInv.TextMatrix(i, 9)) & ", " & Val(flxInv.TextMatrix(i, 10)) & ", " & i & ");"
                        'mySQL
                        LYRcon.Execute "INSERT INTO AUDITINVDTLS ( BATCHID, SPAREPARTID, PHYQTY, LASTUPDATE, QTYONPC, DISCREPANCY, ROWNO ) " & _
                                "VALUES (" & txtBatchID.Text & ", " & flxInv.TextMatrix(i, 11) & ", " & Val(Replace(flxInv.TextMatrix(i, 5), ",", "")) & ", '" & Format(flxInv.TextMatrix(i, 8), "yyyy-mm-dd hh:mm:ss") & "', " & Val(flxInv.TextMatrix(i, 9)) & ", " & Val(flxInv.TextMatrix(i, 10)) & ", " & i & ");"
                        
                    End If
                End If
            Next i
        End If
        txtTemp.Text = Format(Tmp, "#,##0") & Space(7)
        txtFinal.Text = Format((.Rows - 1) - Tmp, "#,##0") & Space(7)
        txtTotal.Text = Format((.Rows - 1), "#,##0") & Space(7)
    End With
    eMsg = "Error 620: Updating Table Audit Inventory Lock Entry"
    
    'LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.LOCKENTRY = False WHERE (((AUDITINVEN.BATCHID)=" & txtBatchID.Text & "));"
    'mySQL
    LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.LOCKENTRY = False WHERE AUDITINVEN.BATCHID=" & txtBatchID.Text & ";"
    
    Exit Sub
LocalErr:
    MsgBox eMsg & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
    
    
End Sub

Private Sub flxInv_Click()
    Dim ans As Integer, MyQty As Integer, MyDesc As String
    Dim Qry As String, eMsg As String
    ActRow = flxInv.row
    ActCol = flxInv.Col
    
    On Error GoTo LocalErr
    If ActCol = 0 Or ActRow = 0 Or LockFlx Then Exit Sub
    Select Case ActCol
        Case 1 'PartNumber
            With txtPartNo
                If flxInv.TextMatrix(ActRow, 1) <> "" Then flxInv.TextMatrix(ActRow, 12) = flxInv.TextMatrix(ActRow, 1) & "|" & flxInv.TextMatrix(ActRow, 2) & "|" & flxInv.TextMatrix(ActRow, 3) & "|" & flxInv.TextMatrix(ActRow, 4) & "|" & flxInv.TextMatrix(ActRow, 5) & "|" & flxInv.TextMatrix(ActRow, 6) & "|" & flxInv.TextMatrix(ActRow, 7) & "|" & flxInv.TextMatrix(ActRow, 8) & "|" & flxInv.TextMatrix(ActRow, 9) & "|" & flxInv.TextMatrix(ActRow, 10) & "|" & flxInv.TextMatrix(ActRow, 11)
                .Top = flxInv.CellTop + flxInv.Top
                .Left = flxInv.CellLeft + flxInv.Left
                .Width = flxInv.CellWidth
                .Text = flxInv.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 2  'Description show Applic. Model tool tip
            MyDesc = flxInv.TextMatrix(ActRow, 2)
            If Len(MyDesc) > 0 Then
                With txtAppModel
                    .Top = flxInv.CellTop + flxInv.Top + 300
                    .Left = flxInv.CellLeft + flxInv.Left
                    .Text = MyDesc
                    .Visible = True
                    .ZOrder
                End With
            End If
        Case 4  'Location
            With txtLocation
                If flxInv.TextMatrix(ActRow, 4) <> "" Then .Tag = flxInv.TextMatrix(ActRow, 4)
                .Top = flxInv.CellTop + flxInv.Top
                .Left = flxInv.CellLeft + flxInv.Left
                .Width = flxInv.CellWidth
                .Text = flxInv.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 5    'qty
            With txtQty
                If flxInv.TextMatrix(ActRow, 5) <> "" Then .Tag = flxInv.TextMatrix(ActRow, 5) & "|" & flxInv.TextMatrix(ActRow, 6) & "|" & flxInv.TextMatrix(ActRow, 7) & "|" & flxInv.TextMatrix(ActRow, 8) & "|" & flxInv.TextMatrix(ActRow, 9) & "|" & flxInv.TextMatrix(ActRow, 10)
                .Top = flxInv.CellTop + flxInv.Top
                .Left = flxInv.CellLeft + flxInv.Left
                .Width = flxInv.CellWidth
                .Text = flxInv.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 6  'Time
            If flxInv.TextMatrix(ActRow, 5) = "" Or Val(flxInv.TextMatrix(ActRow, 11)) = 0 Then Exit Sub
            ans = MsgBox("Do you want to Finalize update on Physical Quantity?", vbQuestion + vbYesNo, "Final")
            If ans = vbNo Then Exit Sub
            'Qry = "SELECT SPAREPARTS.QTYONHAND From SPAREPARTS WHERE (((SPAREPARTS.SPAREPARTID)=" & flxInv.TextMatrix(ActRow, 11) & "));"
            'mySQL
            Qry = "SELECT SPAREPARTS.QTYONHAND From SPAREPARTS WHERE SPAREPARTS.SPAREPARTID=" & flxInv.TextMatrix(ActRow, 11) & ";"
            
            eMsg = "Retrieving Qty. on  hand"
            MyQty = Val(GetFieldValue(Qry))
            flxInv.TextMatrix(ActRow, 9) = MyQty
            flxInv.TextMatrix(ActRow, 10) = Val(Replace(flxInv.TextMatrix(ActRow, 5), ",", "")) - MyQty
            flxInv.TextMatrix(ActRow, 8) = Now
            flxInv.TextMatrix(ActRow, 7) = Format(Now, "mm/dd/yyyy")
            flxInv.TextMatrix(ActRow, 6) = Format(Now, "hh:mm ampm")
            'Qry = "SELECT AUDITINVDTLS.BATCHID From AUDITINVDTLS WHERE (((AUDITINVDTLS.BATCHID)=" & txtBatchID.Text & ") AND ((AUDITINVDTLS.SPAREPARTID)=" & flxInv.TextMatrix(ActRow, 11) & "));"
            'mySQL
            Qry = "SELECT AUDITINVDTLS.BATCHID From AUDITINVDTLS WHERE AUDITINVDTLS.BATCHID=" & txtBatchID.Text & " AND AUDITINVDTLS.SPAREPARTID=" & flxInv.TextMatrix(ActRow, 11) & ";"
            
            eMsg = "Searching for existing Batch ID"
            If CtrlNumExist(Qry) Then
                eMsg = "Updating Table Audit Detailed Inventory for Physical Qty."
                'LYRcon.Execute "UPDATE AUDITINVDTLS SET AUDITINVDTLS.PHYQTY = " & Val(Replace(flxInv.TextMatrix(ActRow, 5), ",", "")) & ", AUDITINVDTLS.LASTUPDATE = #" & flxInv.TextMatrix(ActRow, 8) & "#, AUDITINVDTLS.QTYONPC = " & flxInv.TextMatrix(ActRow, 9) & ", AUDITINVDTLS.DISCREPANCY = " & flxInv.TextMatrix(ActRow, 10) & ", AUDITINVDTLS.ROWNO = " & ActRow & _
                " WHERE (((AUDITINVDTLS.BATCHID)=" & txtBatchID.Text & ") AND ((AUDITINVDTLS.SPAREPARTID)=" & flxInv.TextMatrix(ActRow, 11) & "));"
                'mySQL
                LYRcon.Execute "UPDATE AUDITINVDTLS SET AUDITINVDTLS.PHYQTY = " & Val(Replace(flxInv.TextMatrix(ActRow, 5), ",", "")) & ", AUDITINVDTLS.LASTUPDATE = '" & Format(flxInv.TextMatrix(ActRow, 8), "yyyy-mm-dd hh:mm:ss") & "', AUDITINVDTLS.QTYONPC = " & flxInv.TextMatrix(ActRow, 9) & ", AUDITINVDTLS.DISCREPANCY = " & flxInv.TextMatrix(ActRow, 10) & ", AUDITINVDTLS.ROWNO = " & ActRow & _
                " WHERE AUDITINVDTLS.BATCHID=" & txtBatchID.Text & " AND AUDITINVDTLS.SPAREPARTID=" & flxInv.TextMatrix(ActRow, 11) & ";"
            
            Else
                eMsg = "Adding New Record in Table Audit Detailed Inventory"
                'LYRcon.Execute "INSERT INTO AUDITINVDTLS ( BATCHID, SPAREPARTID, PHYQTY, LASTUPDATE, QTYONPC, DISCREPANCY, ROWNO ) " & _
                              "VALUES (" & txtBatchID.Text & ", " & flxInv.TextMatrix(ActRow, 11) & ", " & Val(Replace(flxInv.TextMatrix(ActRow, 5), ",", "")) & ", #" & flxInv.TextMatrix(ActRow, 8) & "#, " & Val(flxInv.TextMatrix(ActRow, 9)) & ", " & Val(flxInv.TextMatrix(ActRow, 10)) & ", " & ActRow & ");"
                'mySQL
                LYRcon.Execute "INSERT INTO AUDITINVDTLS ( BATCHID, SPAREPARTID, PHYQTY, LASTUPDATE, QTYONPC, DISCREPANCY, ROWNO ) " & _
                              "VALUES (" & txtBatchID.Text & ", " & flxInv.TextMatrix(ActRow, 11) & ", " & Val(Replace(flxInv.TextMatrix(ActRow, 5), ",", "")) & ", '" & Format(flxInv.TextMatrix(ActRow, 8), "yyyy-mm-dd hh:mm:ss") & "', " & Val(flxInv.TextMatrix(ActRow, 9)) & ", " & Val(flxInv.TextMatrix(ActRow, 10)) & ", " & ActRow & ");"
            End If
            MsgBox "Record successfully updated!", vbOKOnly + vbInformation, flxInv.TextMatrix(ActRow, 1)
            flxInv.SetFocus
            flxInv.Col = 6
            If ActRow < flxInv.Rows - 1 Then
                flxInv.row = ActRow + 1
            Else
                flxInv.row = ActRow
            End If
    End Select
    Exit Sub
LocalErr:
    MsgBox eMsg & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
    
End Sub

Private Sub flxInv_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxInv_Click
End Sub

Private Sub flxInv_LeaveCell()
    txtAppModel.Visible = False
End Sub

Private Sub flxInv_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If LockFlx Then Exit Sub
    If Button = vbRightButton And ActCol = 0 Then
        Call PopupMenu(mnuADJ)
    End If
End Sub

Private Sub flxInv_Scroll()
    txtPartNo.Visible = False
    txtLocation.Visible = False
    txtQty.Visible = False
    txtQtyTip.Text = ""
    txtQtyTip.Visible = False
End Sub

Private Sub flxLoc_DblClick()
    Call cmdLocOK_Click
End Sub

Private Sub flxLoc_EnterCell()
    LocRow = flxLoc.row
End Sub

Private Sub flxLoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdLocOK_Click
End Sub

Private Sub flxParts_Click()
    Dim TtlPrice As Currency
    
    With flxParts
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.row, 2)
        'Insert Additional retrieval here
        flxInv.TextMatrix(ActRow, 2) = .TextMatrix(.row, 3) 'Description
        flxInv.TextMatrix(ActRow, 3) = Format$(Val(.TextMatrix(.row, 4)), "#,##0.00") 'Price
        flxInv.TextMatrix(ActRow, 4) = .TextMatrix(.row, 5) 'Location
        flxInv.TextMatrix(ActRow, 11) = .TextMatrix(.row, 1) 'SPid
    End With
End Sub

Private Sub flxParts_EnterCell()
    Dim TtlPrice As Currency
    
    With flxParts
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.row, 2)
        'Insert Additional retrieval here
        flxInv.TextMatrix(ActRow, 2) = .TextMatrix(.row, 3) 'Description
        flxInv.TextMatrix(ActRow, 3) = Format$(Val(.TextMatrix(.row, 4)), "#,##0.00") 'Price
        flxInv.TextMatrix(ActRow, 4) = .TextMatrix(.row, 5) 'Location
        flxInv.TextMatrix(ActRow, 11) = .TextMatrix(.row, 1) 'SPid
    End With
End Sub

Private Sub flxParts_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Tmp() As String
    Dim i As Integer
    
    If KeyCode = vbKeyEscape Then
        Tmp = Split(flxInv.TextMatrix(ActRow, 12), "|")
        For i = LBound(Tmp) To UBound(Tmp)
            flxInv.TextMatrix(ActRow, i + 1) = Tmp(i)
        Next i
        txtPartNo.Visible = False
        flxParts.Visible = False
        flxInv.SetFocus
        flxInv.Col = 1
    End If

End Sub

Private Sub flxParts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtPartNo_LostFocus
        flxInv.SetFocus
        flxInv.row = ActRow
        flxInv.Col = 5
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    Call LockData(True)
    Call ShowButtons(1)
    cmdA.Top = 15000
    Timer1.Enabled = True
    'MotoRep.sBar.Panels(6).Text = "Spare Parts Inventory Auditing"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    MotoRep.sBar.Panels(6).Text = ""
    If rsParts.State = adStateOpen Then rsParts.Close
End Sub


Private Sub mnuAdd_Click()
    If flxInv.TextMatrix(ActRow, 11) = "" Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        flxInv.SetFocus
        flxInv.Col = ActCol
        Exit Sub
    Else
        If ActRow = flxInv.Rows - 1 Then flxInv.AddItem ""
        flxInv.SetFocus
        flxInv.Col = 1
        flxInv.row = flxInv.Rows - 1
    End If
End Sub

Private Sub mnuDelete_Click()
    Dim ans As Integer, i As Integer
    
    ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If ans = vbYes Then
        If ActRow = 1 And flxInv.Rows = 2 Then
            flxInv.Rows = 1
            flxInv.Rows = 2
        Else
            flxInv.RemoveItem ActRow
        End If
    End If
End Sub

Private Sub RXResizer1_ResizeComplete()
    With flxInv
        .ColWidth(0) = 350
        .ColWidth(1) = (.Width - 380) * 0.18
        .ColWidth(2) = (.Width - 380) * 0.21
        .ColWidth(3) = (.Width - 380) * 0.11
        .ColWidth(4) = (.Width - 380) * 0.15
        .ColWidth(5) = (.Width - 380) * 0.11
        .ColWidth(6) = (.Width - 380) * 0.11
        .ColWidth(7) = (.Width - 380) * 0.11
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .ColWidth(12) = 0
    End With
    With flxLoc
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = .Width - 200
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .ColAlignment(2) = 4
    End With
End Sub

Function GetBatchNo() As String
    Dim rsAdjCntr As New ADODB.Recordset
    Dim Qry As String
    Dim MyNo As Long
    
    'Qry = "SELECT Max(Val([BATCHNO])) AS MyBatch From AUDITINVEN WHERE (((AUDITINVEN.SEQNO)=1));"
    
    'mySQL
     Qry = "SELECT Max(round(AUDITINVEN.BATCHNO)) As MyBatch From AUDITINVEN Where AUDITINVEN.SEQNO = 1;"
    
    With rsAdjCntr
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MyNo = IIf(IsNull(.Fields(0)), 1, .Fields(0) + 1)
        Else
            MyNo = 1
        End If
        .Close
    End With
    GetBatchNo = MyNo
End Function

Function ValidateInvAdj() As String
    Dim vMsg As String, Qry As String
    Dim i As Integer, Cntr As Integer
    
    vMsg = ""

    If flxInv.TextMatrix(flxInv.Rows - 1, 11) = "" Then
        If flxInv.Rows <= 2 Then
            vMsg = vMsg & vbCrLf & "Incomplete entry on last row of Adjusted Spare Parts"
        Else
            flxInv.RemoveItem flxInv.Rows - 1
        End If
    End If
    With flxInv
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 11)) = 0 Then
                vMsg = vMsg & vbCrLf & "Invalid Part Number at row " & i
            End If
        Next i
    End With
    ValidateInvAdj = vMsg
End Function

Private Sub Timer1_Timer()
    Call LoadSpareParts
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

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxInv.SetFocus
        flxInv.Col = 5
        flxInv.row = ActRow
    End If
End Sub

Private Sub txtLocation_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtLocation.Text = Empty
        flxInv.TextMatrix(ActRow, 4) = txtLocation.Tag
        txtLocation.Tag = "skiplostfocus"
        flxInv.SetFocus
        flxInv.Col = 4
        txtLocation.Visible = False
    End If
End Sub

Private Sub txtLocation_LostFocus()
    Dim rsLoc As New ADODB.Recordset
    Dim Qry As String
    
    If txtLocation.Tag = "skiplostfocus" Then
        txtLocation.Tag = ""
        Exit Sub
    End If
    If txtLocation.Text = "" Then
        MsgBox "No Location specified!", vbExclamation + vbOKOnly, "Location"
        txtLocation.Visible = False
        Exit Sub
    End If
    'Find Location in Database
    '***Access
    'Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.LOCATION " & _
          "From SPAREPARTS Where (((SPAREPARTS.Location) Like '%" & txtLocation.Text & "%')) ORDER BY SPAREPARTS.PARTNUMBER;"
    
    Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.LOCATION " & _
          "From SPAREPARTS Where (SPAREPARTS.Location) Like '%" & txtLocation.Text & "%' ORDER BY SPAREPARTS.PARTNUMBER;"
    
    With rsLoc
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount = 0 Then
            MsgBox "No Spare Part assigned on this location!", vbExclamation + vbOKOnly, "Location"
            
        ElseIf .RecordCount = 1 Then
            .MoveFirst
            flxInv.TextMatrix(ActRow, 1) = !PartNumber
            flxInv.TextMatrix(ActRow, 2) = !Description
            flxInv.TextMatrix(ActRow, 3) = IIf(IsNull(!STOREPRICE), "", Format(!STOREPRICE, "#,##0.00"))
            flxInv.TextMatrix(ActRow, 4) = txtLocation.Text
            flxInv.TextMatrix(ActRow, 11) = !SPAREPARTID
        ElseIf .RecordCount > 1 Then
            .MoveFirst
            cmdTrans(6).Enabled = False
            cmdTrans(7).Enabled = False
            flxInv.Enabled = False
            Frame2.Enabled = False
            txtFind.Locked = True
            cmdFind.Enabled = False
            cmdInsertSP.Enabled = False
            flxLoc.Rows = 0
            tabLoc.Top = flxInv.Top + 600
            tabLoc.Left = (Me.Width - tabLoc.Width) / 2
            tabLoc.Caption = txtLocation.Text
            tabLoc.Visible = True
            tabLoc.ZOrder
            Do While Not .EOF
                flxLoc.AddItem vbTab & !SPAREPARTID & vbTab & !PartNumber & vbTab & !Description & vbTab & IIf(IsNull(!STOREPRICE), "", Format(!STOREPRICE, "#,##0.00")) & vbTab & txtLocation.Text
                .MoveNext
            Loop
            LocRow = 1
            flxLoc.SetFocus
        End If
        .Close
    End With
    txtLocation.Visible = False
End Sub

Private Sub txtPartNo_Change()
    Dim TtlPrice As Currency, DiscPrice As Currency
    Dim PctDisc As Single, i As Integer
    
    With flxParts
        If PrtSrch Then
            PrtSrch = False
            Exit Sub
        End If
        .Top = txtPartNo.Height + txtPartNo.Top
        .Left = txtPartNo.Left
        .Width = txtPartNo.Width
        .Visible = True
        .ZOrder
        If txtPartNo.Text = "" Then .Visible = False: Exit Sub
        rsParts.Find "[PARTNUMBER] LIKE '" + txtPartNo.Text + "%'", , adSearchForward, 1
        If Not rsParts.EOF Then
            If Not PrtBln Then
                txtPartNo.Tag = Len(txtPartNo.Text)
                PrtBln = True
            End If
            .TopRow = rsParts.AbsolutePosition - 1
            .row = rsParts.AbsolutePosition - 1
            .Col = 0: .ColSel = 5
            txtPartNo.Text = rsParts!PartNumber
            'Insert Additional retrieval here
            flxInv.TextMatrix(ActRow, 2) = .TextMatrix(.row, 3) 'Description
            flxInv.TextMatrix(ActRow, 3) = Format$(Val(.TextMatrix(.row, 4)), "#,##0.00") 'Price
            flxInv.TextMatrix(ActRow, 4) = .TextMatrix(.row, 5) 'Location
            flxInv.TextMatrix(ActRow, 11) = .TextMatrix(.row, 1) 'SPid
            For i = 5 To 10
                flxInv.TextMatrix(ActRow, i) = ""   'Clear Previous Entry
            Next i
            txtPartNo.SelStart = Val(txtPartNo.Tag)
            txtPartNo.SelLength = Len(txtPartNo.Text)
            PrtBln = False
        Else
            flxInv.TextMatrix(ActRow, 2) = ""
            flxInv.TextMatrix(ActRow, 3) = ""
            flxInv.TextMatrix(ActRow, 4) = ""
            flxInv.TextMatrix(ActRow, 11) = ""
            .Visible = False
        End If

    End With
End Sub

Private Sub txtPartNo_GotFocus()
    With flxParts
        .Top = txtPartNo.Height + txtPartNo.Top
        .Left = txtPartNo.Left
        .Width = txtPartNo.Width
    End With
End Sub

Private Sub txtPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim Tmp() As String
    
    Select Case KeyCode
        Case vbKeyDown
            flxParts.Tag = "skiplostfocus"
            flxParts.Visible = True
            flxParts.SetFocus
        Case vbKeyEscape
            txtPartNo.Text = Empty
            If flxInv.TextMatrix(ActRow, 12) <> "" Then
                Tmp = Split(flxInv.TextMatrix(ActRow, 12), "|")
                For i = LBound(Tmp) To UBound(Tmp)
                    flxInv.TextMatrix(ActRow, i + 1) = Tmp(i)
                Next i
            Else
                For i = 1 To 12
                    flxInv.TextMatrix(ActRow, i) = ""
                Next i
            End If
            flxParts.Tag = "skiplostfocus"
            flxInv.SetFocus
            flxInv.Col = 1
            txtPartNo.Visible = False
            flxParts.Visible = False
        Case vbKeyDelete
            If rsParts.State = adStateOpen Then rsParts.MoveFirst
            PrtSrch = True
        Case vbKeyBack
            If rsParts.State = adStateOpen Then rsParts.MoveFirst
            PrtSrch = True
    End Select
End Sub

Private Sub txtPartNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxInv.SetFocus
        flxInv.row = ActRow
        flxInv.Col = 5
    End If
End Sub

Private Sub txtPartNo_LostFocus()
    If flxParts.Tag = "skiplostfocus" Then
        flxParts.Tag = ""
        Exit Sub
    End If
    If flxInv.TextMatrix(ActRow, 11) = "" Then
        MsgBox "Unrecognized Part Number!" & vbCrLf & "Please select a Part Number from list.", vbExclamation + vbOKOnly, "Warning"
        txtPartNo.Visible = False
        flxParts.Visible = False
        Exit Sub
    End If
    flxInv.TextMatrix(ActRow, 1) = txtPartNo.Text
    txtPartNo.Visible = False
    flxParts.Visible = False
End Sub
Private Sub txtQty_GotFocus()
    Dim Qry As String, Dq As Integer
    
    If txtDispRef.Text = "" Or Val(flxInv.TextMatrix(ActRow, 11)) = 0 Then Exit Sub
    
    'Qry = "SELECT AUDITINVDTLS.PHYQTY FROM AUDITINVEN INNER JOIN AUDITINVDTLS ON AUDITINVEN.BATCHID = AUDITINVDTLS.BATCHID " & _
          "WHERE (((AUDITINVEN.BATCHNO)='" & txtDispRef.Text & "') AND ((AUDITINVDTLS.SPAREPARTID)=" & Val(flxInv.TextMatrix(ActRow, 11)) & "));"
    
    'mySQL
    Qry = "SELECT AUDITINVDTLS.PHYQTY FROM AUDITINVEN INNER JOIN AUDITINVDTLS ON AUDITINVEN.BATCHID = AUDITINVDTLS.BATCHID " & _
          "WHERE AUDITINVEN.BATCHNO='" & txtDispRef.Text & "' AND AUDITINVDTLS.SPAREPARTID=" & Val(flxInv.TextMatrix(ActRow, 11)) & ";"
    Dq = Val(GetFieldValue(Qry))
    
    With txtQtyTip
        .Text = IIf(Dq <> 0, "Qty. on Display: " & Dq, "No Items on Display")
        .Top = txtQty.Top + 40
        .Left = flxInv.Left + flxInv.ColWidth(0) + flxInv.ColWidth(1) - 400
        .ZOrder
        .Visible = True
    End With

End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim Tmp() As String
    
    If KeyCode = vbKeyEscape Then
        txtQty.Text = Empty
        If txtQty.Tag <> "" Then
            Tmp = Split(txtQty.Tag, "|")
            For i = LBound(Tmp) To UBound(Tmp)
                flxInv.TextMatrix(ActRow, i + 5) = Tmp(i)
            Next i
        Else
            For i = 5 To 10
                flxInv.TextMatrix(ActRow, i) = ""
            Next i
        End If
        txtQty.Tag = "skiplostfocus"
        flxInv.SetFocus
        flxInv.Col = 5
        txtQty.Visible = False
    End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        With flxInv
            .SetFocus
            If .TextMatrix(ActRow, 11) <> "" Then
                If ActRow = .Rows - 1 Then .AddItem ""
                .row = .Rows - 1
                .Col = 1
            Else
                .row = ActRow
                .Col = 5
            End If
        End With
    End If
End Sub

Private Sub txtQty_LostFocus()
    Dim i As Integer
    
    If txtQty.Tag = "skiplostfocus" Then
        txtQty.Tag = ""
        txtQtyTip.Text = ""
        txtQtyTip.Visible = False
        Exit Sub
    End If
    If txtQty.Text <> txtQty.Tag Then
        For i = 5 To 10
            flxInv.TextMatrix(ActRow, i) = ""
        Next i
    End If
    flxInv.TextMatrix(ActRow, 5) = txtQty.Text
    txtQty.Visible = False
    txtQtyTip.Text = ""
    txtQtyTip.Visible = False
End Sub

Private Function ConvertToRoman(pstrDecimalNumber As String) As String
'--------------------------------------------------------------------------------------

    Const strPOS_VAL    As String = "IXCM"
    Const strFIVE_VAL   As String = "VLD"
    
    Dim strRoman        As String
    Dim strCurrRomanPos As String
    Dim strLetter1      As String
    Dim strLetter2      As String
    Dim intCurrPos      As Integer
    Dim intDigit        As Integer
    Dim intDigitPos     As Integer
    
    intCurrPos = 1
    strRoman = ""
    
    For intDigitPos = Len(pstrDecimalNumber) To 1 Step -1
        intDigit = Val(Mid$(pstrDecimalNumber, intDigitPos, 1))
        strCurrRomanPos = Mid$(strPOS_VAL, intCurrPos, 1)
        Select Case intDigit
            Case 9
                strLetter1 = strCurrRomanPos
                strLetter2 = Mid$(strPOS_VAL, intCurrPos + 1, 1)
            Case Is > 4
                strLetter1 = Mid$(strFIVE_VAL, intCurrPos, 1)
                strLetter2 = String$(intDigit - 5, strCurrRomanPos)
            Case 4
                strLetter1 = strCurrRomanPos
                strLetter2 = Mid$(strFIVE_VAL, intCurrPos, 1)
            Case Else
                strLetter1 = String$(intDigit, strCurrRomanPos)
                strLetter2 = ""
        End Select
        strRoman = strLetter1 & strLetter2 & strRoman
        intCurrPos = intCurrPos + 1
    Next

    ConvertToRoman = strRoman

End Function

