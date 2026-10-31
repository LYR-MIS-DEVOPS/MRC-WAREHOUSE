VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMonthlyPhysicalInvSumm 
   Caption         =   "Monthly Physical Inventory Summary"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   13905
   WindowState     =   2  'Maximized
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   495
      Left            =   5420
      TabIndex        =   26
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
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
      Image           =   "frmMonthlyPhysicalInvSumm.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdShowDtls 
      Height          =   495
      Left            =   3550
      TabIndex        =   11
      Top             =   0
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   873
      Caption         =   "&View Details"
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
      Image           =   "frmMonthlyPhysicalInvSumm.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdDiscrepancy 
      Height          =   495
      Left            =   1930
      TabIndex        =   13
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
      Image           =   "frmMonthlyPhysicalInvSumm.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdXLS 
      Height          =   495
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Import P.O. from Excel"
      Top             =   0
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      Caption         =   "&Upload Report"
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
      Image           =   "frmMonthlyPhysicalInvSumm.frx":10CE
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame2 
      Height          =   6855
      Left            =   120
      TabIndex        =   9
      Tag             =   "XY"
      Top             =   1080
      Width           =   13695
      Begin TabDlg.SSTab ssTViewDetails 
         Height          =   6255
         Left            =   480
         TabIndex        =   14
         Top             =   3120
         Visible         =   0   'False
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   11033
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         TabCaption(0)   =   "VIEW DETAILS"
         TabPicture(0)   =   "frmMonthlyPhysicalInvSumm.frx":1668
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblViewDetailsClose"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame3"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lvButtons_H1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Text1"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Text3"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   1080
            TabIndex        =   24
            Tag             =   "XYF"
            Top             =   645
            Width           =   2295
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   13200
            TabIndex        =   17
            Tag             =   "XYF"
            Top             =   645
            Width           =   1935
         End
         Begin lvButton.lvButtons_H lvButtons_H1 
            Height          =   375
            Left            =   11880
            TabIndex        =   18
            Tag             =   "XY"
            Top             =   600
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
            Image           =   "frmMonthlyPhysicalInvSumm.frx":1684
            cBack           =   -2147483633
         End
         Begin VB.Frame Frame3 
            Height          =   5175
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   15015
            Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
               Height          =   4815
               Left            =   120
               TabIndex        =   16
               Top             =   240
               Width           =   14775
               _ExtentX        =   26061
               _ExtentY        =   8493
               _Version        =   393216
               Cols            =   10
               FormatString    =   $"frmMonthlyPhysicalInvSumm.frx":1C1E
            End
         End
         Begin VB.Label lblViewDetailsClose 
            Alignment       =   2  'Center
            Caption         =   " ( X )"
            Height          =   255
            Left            =   14760
            MousePointer    =   99  'Custom
            TabIndex        =   27
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "Batch No:"
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
            Left            =   240
            TabIndex        =   25
            Top             =   645
            Width           =   1335
         End
      End
      Begin TabDlg.SSTab ssTDiscrepancy 
         Height          =   6015
         Left            =   480
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   18255
         _ExtentX        =   32200
         _ExtentY        =   10610
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         TabCaption(0)   =   "DISCREPANCY REPORT"
         TabPicture(0)   =   "frmMonthlyPhysicalInvSumm.frx":1CF4
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblDiscrepancyClose"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lvButtons_H2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame4"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Text2"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cmdTrans(1)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         Begin lvButton.lvButtons_H cmdTrans 
            Height          =   375
            Index           =   1
            Left            =   16560
            TabIndex        =   29
            Top             =   480
            Width           =   1560
            _ExtentX        =   2752
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
            Image           =   "frmMonthlyPhysicalInvSumm.frx":1D10
            cBack           =   -2147483633
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1560
            TabIndex        =   22
            Tag             =   "XYF"
            Top             =   405
            Width           =   1935
         End
         Begin VB.Frame Frame4 
            Height          =   5055
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Width           =   18015
            Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
               Height          =   4695
               Left            =   120
               TabIndex        =   21
               Top             =   240
               Width           =   17775
               _ExtentX        =   31353
               _ExtentY        =   8281
               _Version        =   393216
               Cols            =   11
               FormatString    =   $"frmMonthlyPhysicalInvSumm.frx":22AA
            End
         End
         Begin lvButton.lvButtons_H lvButtons_H2 
            Height          =   375
            Left            =   240
            TabIndex        =   23
            Tag             =   "XY"
            Top             =   360
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
            Image           =   "frmMonthlyPhysicalInvSumm.frx":23B6
            cBack           =   -2147483633
         End
         Begin VB.Label lblDiscrepancyClose 
            Alignment       =   2  'Center
            Caption         =   " ( X )"
            Height          =   255
            Left            =   17760
            MousePointer    =   99  'Custom
            TabIndex        =   28
            Top             =   120
            Width           =   375
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   6495
         Left            =   120
         TabIndex        =   10
         Tag             =   "XY"
         Top             =   240
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   11456
         _Version        =   393216
         Cols            =   9
         FormatString    =   $"frmMonthlyPhysicalInvSumm.frx":2950
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Tag             =   "XY"
      Top             =   480
      Width           =   13695
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   11640
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   195
         Width           =   1935
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   10320
         TabIndex        =   2
         Tag             =   "XY"
         Top             =   150
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
         Image           =   "frmMonthlyPhysicalInvSumm.frx":2A2E
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dFrom 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   180
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   103874561
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   375
         Left            =   3240
         TabIndex        =   4
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
         Image           =   "frmMonthlyPhysicalInvSumm.frx":2FC8
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   375
         Left            =   5400
         TabIndex        =   5
         ToolTipText     =   "Execute"
         Top             =   120
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
         Image           =   "frmMonthlyPhysicalInvSumm.frx":3562
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dTo 
         Height          =   315
         Left            =   3720
         TabIndex        =   6
         Top             =   180
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   103874561
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdPrev 
         Height          =   375
         Left            =   5880
         TabIndex        =   7
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
         Image           =   "frmMonthlyPhysicalInvSumm.frx":3AFC
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
         Left            =   120
         TabIndex        =   8
         Top             =   210
         Width           =   1335
      End
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   7200
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14145
      m_MDI_Height    =   8595
   End
   Begin VB.Menu mnuMain 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuApproved 
         Caption         =   "Approved"
      End
   End
   Begin VB.Menu mnuDiscrepancy 
      Caption         =   "Discrepancy"
      Visible         =   0   'False
      Begin VB.Menu mnuGood 
         Caption         =   "Good"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExcess 
         Caption         =   "Excess"
      End
      Begin VB.Menu mnuLacking 
         Caption         =   "Lacking"
      End
   End
End
Attribute VB_Name = "frmMonthlyPhysicalInvSumm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDiscrepancy_Click()
    frmDiscrepancyByBatchNo.Show vbModal
'    ssTDiscrepancy.Visible = True
'    ssTDiscrepancy.Left = (Me.Width - ssTDiscrepancy.Width) / 2
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdShowDtls_Click()
    ssTViewDetails.Visible = True
    ssTViewDetails.Left = (Me.Width - ssTDiscrepancy.Width) / 2
End Sub

Private Sub cmdXLS_Click()
    frmMonthlyPhysicalInv.Show vbModal
End Sub

Private Sub lblDiscrepancyClose_Click()
    ssTDiscrepancy.Visible = False
End Sub

Private Sub lblViewDetailsClose_Click()
    ssTViewDetails.Visible = False
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuMain, vbPopupMenuRightButton
    End If
End Sub

Private Sub MSFlexGrid3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuDiscrepancy, vbPopupMenuRightButton
    End If
End Sub
