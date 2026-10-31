VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPriceBull 
   Caption         =   "PRICE BULLETIN"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13950
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   13950
   StartUpPosition =   2  'CenterScreen
   Tag             =   "XYF"
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Tag             =   "XYZ"
      Top             =   480
      Width           =   13695
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   375
         Left            =   3240
         TabIndex        =   1
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
         Image           =   "frmPriceBull.frx":0000
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   375
         Left            =   5400
         TabIndex        =   2
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
         Image           =   "frmPriceBull.frx":059A
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dTo 
         Height          =   315
         Left            =   3720
         TabIndex        =   3
         Top             =   180
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   109314049
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdPrev 
         Height          =   375
         Left            =   5880
         TabIndex        =   4
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
         Image           =   "frmPriceBull.frx":0B34
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dFrom 
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   180
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   109248513
         CurrentDate     =   39322
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
         TabIndex        =   5
         Top             =   210
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5295
      Left            =   120
      TabIndex        =   6
      Tag             =   "XYF"
      Top             =   1080
      Width           =   13695
      Begin MSFlexGridLib.MSFlexGrid flxSRP 
         Height          =   4935
         Left            =   120
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   240
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   8705
         _Version        =   393216
         Rows            =   3
         Cols            =   11
         FixedRows       =   2
         RowHeightMin    =   315
         BackColor       =   16777215
         BackColorBkg    =   12632256
         WordWrap        =   -1  'True
         HighLight       =   2
         GridLines       =   2
         AllowUserResizing=   3
         FormatString    =   $"frmPriceBull.frx":10CE
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
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MRC PRICE BULLETIN"
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
      Index           =   2
      Left            =   9240
      TabIndex        =   8
      Tag             =   "XYF"
      Top             =   120
      Width           =   4500
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   6840
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14190
      m_MDI_Height    =   7095
      OriginalBorder  =   10
      OriginalCaption =   31
   End
End
Attribute VB_Name = "frmPriceBull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub flxSRPColsChanger()
On Error Resume Next
Dim i As Integer
With flxSRP
    
    'MERGE COL
    .TextMatrix(0, 0) = " "
    .TextMatrix(0, 1) = "PART NUMBER"
    .TextMatrix(0, 2) = "DESCRIPTION"
    .TextMatrix(0, 3) = "DATE APPROVED"
    .TextMatrix(0, 4) = "RETAIL"
    .TextMatrix(0, 5) = "RETAIL"
    .TextMatrix(0, 6) = "RETAIL"
    .TextMatrix(0, 7) = "WHOLESALE"
    .TextMatrix(0, 8) = "WHOLESALE"
    .TextMatrix(0, 9) = "WHOLESALE"
    .TextMatrix(0, 10) = "STATUS"
        
    .TextMatrix(1, 0) = " "
    .TextMatrix(1, 1) = "PART NUMBER"
    .TextMatrix(1, 2) = "DESCRIPTION"
    .TextMatrix(1, 3) = "DATE APPROVED"
    .TextMatrix(1, 4) = "OLD SRP"
    .TextMatrix(1, 5) = "NEW SRP"
    .TextMatrix(1, 6) = "SRP DIFF."
    .TextMatrix(1, 7) = "OLD PRICE"
    .TextMatrix(1, 8) = "NEW PRICE"
    .TextMatrix(1, 9) = "PRICE DIFF."
    
    .TextMatrix(1, 10) = "STATUS"
        
    'SET MERGING
    .MergeCells = 1
    .MergeRow(0) = True
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = True
    .MergeCol(10) = True
    
    For i = 4 To 9
    .MergeCol(i) = False
    Next i
    
    .MergeRow(0) = True
           
End With

Call flxSRPSetColsWidth

End Sub

Sub flxSRPSetColsWidth()
On Error Resume Next
    With flxSRP
            
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.1 'PART NUMBER
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.1 'DESCRIPTION
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.08  'DATE APPROVED
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.09  'DIPSTICK(cm)
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.09  'CALIBRATION
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.09  'Liters(Calibration)
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.09  'Liters(truck)
        .ColWidth(8) = (.Width - .ColWidth(0)) * 0.09  'DIPSTICK(cm)
        .ColWidth(9) = (.Width - .ColWidth(0)) * 0.09  'CALIBRATION
        .ColWidth(10) = (.Width - .ColWidth(0)) * 0.15  'AMOUNT
        
    End With
End Sub

Private Sub Form_Load()
    Call flxSRPColsChanger
End Sub

Private Sub RXResizer1_ResizeComplete()
    Call flxSRPColsChanger
End Sub
