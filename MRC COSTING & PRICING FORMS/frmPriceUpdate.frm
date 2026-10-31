VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPriceUpdate 
   Caption         =   "PRICE FACTOR"
   ClientHeight    =   9390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14565
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9390
   ScaleWidth      =   14565
   WindowState     =   2  'Maximized
   Begin lvButton.lvButtons_H cmdRefresh 
      Height          =   375
      Left            =   3000
      TabIndex        =   54
      Top             =   120
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   661
      Caption         =   "&REFRESH"
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
      Image           =   "frmPriceUpdate.frx":0000
      cBack           =   -2147483633
   End
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7560
      Top             =   0
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   15000
      Locked          =   -1  'True
      TabIndex        =   47
      Text            =   "txtHoldGrayBackColor"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   15000
      TabIndex        =   46
      Text            =   "HoldBackColor"
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   120
      TabIndex        =   41
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   " E&XIT FORM"
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
      Image           =   "frmPriceUpdate.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdShow 
      Height          =   375
      Left            =   1560
      TabIndex        =   42
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "SRP &FORM"
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
      Image           =   "frmPriceUpdate.frx":5D8C
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Tag             =   "XYF"
      Top             =   480
      Width           =   14295
      Begin VB.ComboBox cboSupplier 
         Height          =   315
         Left            =   11040
         TabIndex        =   39
         Tag             =   "XYF"
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtFindParts 
         Height          =   315
         Left            =   2520
         TabIndex        =   3
         Top             =   240
         Width           =   3135
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   250
         Left            =   2210
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxFind 
         Height          =   3735
         Left            =   120
         TabIndex        =   51
         Tag             =   "XYF"
         Top             =   600
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   6588
         _Version        =   393216
         Cols            =   10
         RowHeightMin    =   300
         AllowBigSelection=   0   'False
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
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
         _Band(0).Cols   =   10
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "SELECT A SUPPLIER:"
         Height          =   255
         Left            =   8880
         TabIndex        =   40
         Tag             =   "XYF"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblFindParts 
         Caption         =   "FIND PART DESCRIPTION:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4335
      Left            =   9000
      TabIndex        =   5
      Tag             =   "XYF"
      Top             =   4920
      Width           =   5415
      Begin MSFlexGridLib.MSFlexGrid flxSRPPriceHis 
         Height          =   1935
         Left            =   120
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   480
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   3413
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "^DATE                |^REF. NO.              |^PREV. COST    |^PREV. SRP   "
      End
      Begin MSFlexGridLib.MSFlexGrid flxWSPriceHis 
         Height          =   1455
         Left            =   120
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   2760
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   2566
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "^DATE                |^REF. NO.              |^PREV. COST    |^PREV. SRP   "
      End
      Begin VB.Label Label3 
         Caption         =   "WHOLESALE PRICE HISTORY"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   2520
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "SRP PRICE HISTORY"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4335
      Left            =   120
      TabIndex        =   4
      Tag             =   "XYF"
      Top             =   4920
      Width           =   8655
      Begin VB.TextBox txtNetofVat 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   7080
         TabIndex        =   17
         Tag             =   "XYF"
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtRefDate 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   18
         Tag             =   "XYF"
         Top             =   540
         Width           =   1455
      End
      Begin lvButton.lvButtons_H cmdTrans 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   661
         Caption         =   "&EDIT"
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
         Image           =   "frmPriceUpdate.frx":6326
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   2655
         Left            =   120
         TabIndex        =   20
         Tag             =   "XYF"
         Top             =   1560
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4683
         _Version        =   393216
         TabOrientation  =   1
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         TabCaption(0)   =   "     ."
         TabPicture(0)   =   "frmPriceUpdate.frx":68C0
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame5"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         Begin VB.Frame Frame5 
            Height          =   2415
            Left            =   4200
            TabIndex        =   31
            Tag             =   "XYF"
            Top             =   120
            Width           =   4095
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "   PROFIT RATIO"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   53
               Tag             =   "XYF"
               Top             =   720
               Width           =   1575
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "PROFIT MARGIN"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   37
               Tag             =   "XYF"
               Top             =   1160
               Width           =   1575
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "           FIX PRICE"
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   36
               Tag             =   "XYF"
               Top             =   1600
               Width           =   1575
            End
            Begin VB.TextBox txtWS 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   1730
               Locked          =   -1  'True
               TabIndex        =   35
               Tag             =   "XYF"
               Top             =   240
               Width           =   1935
            End
            Begin VB.TextBox txtWSProfRatio 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   1730
               Locked          =   -1  'True
               TabIndex        =   34
               Tag             =   "XYF"
               Top             =   720
               Width           =   1935
            End
            Begin VB.TextBox txtWSProfMar 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   1730
               Locked          =   -1  'True
               TabIndex        =   33
               Tag             =   "XYF"
               Top             =   1160
               Width           =   1935
            End
            Begin VB.TextBox txtWSFixPrice 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   1730
               Locked          =   -1  'True
               TabIndex        =   32
               Tag             =   "XYF"
               Top             =   1600
               Width           =   1935
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               Caption         =   "%"
               Height          =   255
               Left            =   3680
               TabIndex        =   50
               Tag             =   "XYF"
               Top             =   780
               Width           =   135
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               Caption         =   "WHOLESALE:"
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
               Left            =   240
               TabIndex        =   38
               Tag             =   "XYF"
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.Frame Frame4 
            Height          =   2415
            Left            =   120
            TabIndex        =   21
            Tag             =   "XYF"
            Top             =   120
            Width           =   3975
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "   PROFIT RATIO"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   52
               Tag             =   "XYF"
               Top             =   720
               Width           =   1575
            End
            Begin VB.TextBox txtSRPMaxDis 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   1725
               Locked          =   -1  'True
               TabIndex        =   30
               Tag             =   "XYF"
               Top             =   2000
               Width           =   1815
            End
            Begin VB.TextBox txtSRPFixPrice 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   1730
               Locked          =   -1  'True
               TabIndex        =   28
               Tag             =   "XYF"
               Top             =   1560
               Width           =   1815
            End
            Begin VB.TextBox txtSRPProfMar 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   1730
               Locked          =   -1  'True
               TabIndex        =   27
               Tag             =   "XYF"
               Top             =   1160
               Width           =   1815
            End
            Begin VB.TextBox txtSRPProfRatio 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   1730
               Locked          =   -1  'True
               TabIndex        =   26
               Tag             =   "XYF"
               Top             =   720
               Width           =   1815
            End
            Begin VB.TextBox txtSRP 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   1730
               Locked          =   -1  'True
               TabIndex        =   25
               Tag             =   "XYF"
               Top             =   240
               Width           =   1815
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "           FIX PRICE"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   24
               Tag             =   "XYF"
               Top             =   1600
               Width           =   1575
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "PROFIT MARGIN"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   23
               Tag             =   "XYF"
               Top             =   1160
               Width           =   1575
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               Caption         =   "%"
               Height          =   255
               Left            =   3550
               TabIndex        =   49
               Tag             =   "XYF"
               Top             =   780
               Width           =   135
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Caption         =   "%"
               Height          =   255
               Left            =   3550
               TabIndex        =   48
               Tag             =   "XYF"
               Top             =   2040
               Width           =   135
            End
            Begin VB.Label Label11 
               Caption         =   " MAX DISCOUNT:"
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Tag             =   "XYF"
               Top             =   1995
               Width           =   1335
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               Caption         =   "SRP:"
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
               Left            =   840
               TabIndex        =   22
               Tag             =   "XYF"
               Top             =   240
               Width           =   735
            End
         End
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   16
         Tag             =   "XYF"
         Top             =   1200
         Width           =   7095
      End
      Begin VB.TextBox txtRefNo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   14
         Tag             =   "XYF"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtPartNo 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   840
         Width           =   3135
      End
      Begin lvButton.lvButtons_H cmdTrans 
         Height          =   375
         Index           =   2
         Left            =   1320
         TabIndex        =   44
         Top             =   240
         Visible         =   0   'False
         Width           =   1080
         _ExtentX        =   1905
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
         Image           =   "frmPriceUpdate.frx":68DC
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdTrans 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Visible         =   0   'False
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   661
         Caption         =   "&SAVE"
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
         Image           =   "frmPriceUpdate.frx":6E76
         cBack           =   -2147483633
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "REF. DATE:"
         Height          =   255
         Left            =   5400
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   555
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "REF. NO.:"
         Height          =   255
         Left            =   5400
         TabIndex        =   15
         Tag             =   "XYF"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "NET OF VAT:"
         Height          =   255
         Left            =   5400
         TabIndex        =   13
         Tag             =   "XYF"
         Top             =   870
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "DESCRIPTION:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "PART NUMBER:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   870
         Width           =   1455
      End
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   11280
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14805
      m_MDI_Height    =   9960
      OriginalBorder  =   10
      OriginalCaption =   31
   End
   Begin VB.Menu mnuPriceUpdater 
      Caption         =   "PRICEUPDATER"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmPriceUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ans As VbMsgBoxResult
Dim AllArrSupplier() As String
Dim arrSupplier() As String

Dim ActRow As Integer, fRow As Integer, fCol As Integer, TCol As Integer, AccID As Integer, ColPrn As Integer

Dim SearchcboSup As Boolean
Dim SupID As Integer
Dim pSupID As Integer
Dim didRecordEntry As Boolean

Dim SPid As Integer


Private Sub cboSupplier_Change()
On Error GoTo er
    Dim i As Integer, j As Integer
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    
    If SearchcboSup Then
        SearchcboSup = False
        Exit Sub
    End If


    With cboSupplier
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters

        If i <> -1 And i <= .ListCount Then
            strTotal = .List(i)
            
            SupID = arrSupplier(i, 0)
    
            DoEvents
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
               SearchcboSup = True
               .SelText = Right$(strTotal, j)
               .SelStart = Len(strPartial)
               .SelLength = j
            End If

        End If

    End With
        
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error ->cboSupplier_Change()")
End Sub

Private Sub cboSupplier_Click()
    
    SupID = arrSupplier(cboSupplier.ListIndex, 0)

    
    If SupID <> pSupID Then
        pSupID = SupID
        
        Screen.MousePointer = vbHourglass
        tmrRefresh.Enabled = True
        DoEvents
        frm_OnProcess2.OwnerForm = "SRP Updater"
        Call Load(frm_OnProcess2)
        frm_OnProcess2.Show vbModal
        
    End If
    
End Sub

Private Sub cboSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Or KeyCode = 46 Then
        SearchcboSup = True
    End If
End Sub

Private Sub cboSupplier_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboSupplier, KeyAscii, ValidChars)
    
    If KeyAscii = 13 Then
        If SupID <> pSupID Then
            pSupID = SupID
            
            Screen.MousePointer = vbHourglass
            tmrRefresh.Enabled = True
            DoEvents
            frm_OnProcess2.OwnerForm = "SRP Updater"
            Call Load(frm_OnProcess2)
            frm_OnProcess2.Show vbModal
        End If
    End If
End Sub


Private Sub Check1_Click(Index As Integer)
On Error GoTo er
    
    If didRecordEntry = False Then
        Check1(0).Value = 0
        Check1(1).Value = 0
        Check1(2).Value = 0
        Check1(3).Value = 0
        Check1(4).Value = 0
        Check1(5).Value = 0
        Exit Sub
    End If

    If Index = 0 Then ' Check box SRP Prof Ratio
        
        If Check1(0).Value = 1 Then
            txtSRPProfRatio.Locked = False
            txtSRPProfRatio.BackColor = txtHoldWhiteBackColor.BackColor
            txtSRPProfRatio.SetFocus
            Check1(1).Value = 0
            Check1(2).Value = 0
        Else
            txtSRPProfRatio.Locked = True
            txtSRPProfRatio.BackColor = txtHoldGrayBackColor.BackColor
        End If

    ElseIf Index = 1 Then ' Check box SRP Prof Margin
    
        If Check1(1).Value = 1 Then
            txtSRPProfMar.Locked = False
            txtSRPProfMar.BackColor = txtHoldWhiteBackColor.BackColor
            txtSRPProfMar.SetFocus
            Check1(0).Value = 0
            Check1(2).Value = 0
        Else
            txtSRPProfMar.Locked = True
            txtSRPProfMar.BackColor = txtHoldGrayBackColor.BackColor
        End If
        
    ElseIf Index = 2 Then ' Check box SRP Fix Price
        
        If Check1(2).Value = 1 Then
            txtSRPFixPrice.Locked = False
            txtSRPFixPrice.BackColor = txtHoldWhiteBackColor.BackColor
            txtSRPFixPrice.SetFocus
            Check1(0).Value = 0
            Check1(1).Value = 0
        Else
            txtSRPFixPrice.Locked = True
            txtSRPFixPrice.BackColor = txtHoldGrayBackColor.BackColor
        End If
    
    ElseIf Index = 3 Then ' Check box WS Prof Ratio
        
        If Check1(3).Value = 1 Then
            txtWSProfRatio.Locked = False
            txtWSProfRatio.BackColor = txtHoldWhiteBackColor.BackColor
            txtWSProfRatio.SetFocus
            Check1(4).Value = 0
            Check1(5).Value = 0
        Else
            txtWSProfRatio.Locked = True
            txtWSProfRatio.BackColor = txtHoldGrayBackColor.BackColor
        End If

    ElseIf Index = 4 Then ' Check box WS Prof Margin
        
        If Check1(4).Value = 1 Then
            txtWSProfMar.Locked = False
            txtWSProfMar.BackColor = txtHoldWhiteBackColor.BackColor
            txtWSProfMar.SetFocus
            Check1(3).Value = 0
            Check1(5).Value = 0
        Else
            txtWSProfMar.Locked = True
            txtWSProfMar.BackColor = txtHoldGrayBackColor.BackColor
        End If
        
    ElseIf Index = 5 Then ' Check box WS Fix Price
         
        If Check1(5).Value = 1 Then
            txtWSFixPrice.Locked = False
            txtWSFixPrice.BackColor = txtHoldWhiteBackColor.BackColor
            txtWSFixPrice.SetFocus
            Check1(3).Value = 0
            Check1(4).Value = 0
        Else
            txtWSFixPrice.Locked = True
            txtWSFixPrice.BackColor = txtHoldGrayBackColor.BackColor
        End If
        
    End If

Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error ->Check1_Click")
End Sub

Private Sub cmdExit_Click()
    
    Ans = MsgBox("ARE YOU SURE YOU WANT TO EXIT?", vbQuestion + vbYesNo, Me.Caption)
    If Ans = vbNo Then Exit Sub
    
    Unload Me
End Sub

Private Sub cmdRefresh_Click()

    If cboSupplier.Text = "" Then SupID = 0
    
    Screen.MousePointer = vbHourglass
    tmrRefresh.Enabled = True
    DoEvents
    frm_OnProcess2.OwnerForm = "SRP Updater"
    Call Load(frm_OnProcess2)
    frm_OnProcess2.Show vbModal
End Sub

Private Sub cmdShow_Click()
    frmSRP.Show
End Sub

Private Sub cmdTrans_Click(Index As Integer)
On Error GoTo er
'cmdTrans Legend
'0 ->EDIT
'1 ->SAVE
'2 ->CANCEL

    Dim Level As String
    Dim message As String
    
    If Index = 0 Then 'Edit
        Ans = MsgBox("Do you want to edit price?", vbYesNo + vbInformation, Me.Caption)
        If Ans = vbNo Then Exit Sub
        
        cmdTrans(0).Visible = False
        cmdTrans(1).Visible = True
        cmdTrans(2).Visible = True
                
        txtRefNo.Locked = False
        txtRefNo.BackColor = txtHoldWhiteBackColor.BackColor
                
        txtRefDate.Locked = False
        txtRefDate.Text = Format(Now(), "mm / dd / yyyy")
        txtRefDate.Locked = True
        
        txtSRPMaxDis.Locked = False
        txtSRPMaxDis.BackColor = txtHoldWhiteBackColor.BackColor
        
        didRecordEntry = True
        
    ElseIf Index = 1 Then 'SAVE
        Dim maxdis As Integer
        
        If ValidateEntry() <> "" Then
            MsgBox ValidateEntry, vbExclamation, Me.Caption
            Exit Sub
        End If
        
        Ans = MsgBox("DO YOU WANT TO SAVE CHANGE?", vbQuestion + vbYesNo, Me.Caption)
        If Ans = vbNo Then Exit Sub
            
       Screen.MousePointer = vbHourglass
            
            If txtSRPMaxDis.Text <> "" Then
                maxdis = txtSRPMaxDis.Text
            End If
            
            Call modSRP.sp_priceupdate_save(Trim(SPid), Trim(txtRefNo.Text), Format(txtNetofVat.Text, "#,##0.00") _
                                            , Format(txtSRPFixPrice.Text, "#,##0"), Format(txtSRPProfMar.Text, "#,##0") _
                                            , txtSRPProfRatio.Text & "%", Format(txtWSFixPrice.Text, "#,##0") _
                                            , Format(txtWSProfMar.Text, "#,##0"), txtWSProfRatio.Text & "%", maxdis, Level, message)
    
       
        
       Screen.MousePointer = vbDefault
        
        If Trim(Level) = "" Then
            Exit Sub
        ElseIf Level = "Granted" Or Level = "Warning" Then
            If Level = "Granted" Then
                MsgBox message, vbInformation, "SAVE."
            Else
                MsgBox message, vbExclamation, "SAVE!"
                Exit Sub
            End If
        Else
            MsgBox message, vbCritical, "Error!"
            Exit Sub
        End If


        cmdTrans(0).Visible = True
        cmdTrans(1).Visible = False
        cmdTrans(2).Visible = False
        cmdTrans(0).Enabled = False
        
        txtSRP.Text = Format(txtSRPFixPrice.Text, "#,##0")
        txtWS.Text = Format(txtWSFixPrice.Text, "#,##0")
        
        Call LockedEntryObject(False)
        
        DoEvents
        Call LoadWSHist(SPid)
        DoEvents
        Call LoadSRPHist(SPid)
        
        didRecordEntry = False
    
    ElseIf Index = 2 Then 'CANCEL
        Ans = MsgBox("Cancel the currrent operation?", vbYesNo + vbQuestion, Me.Caption)
        If Ans = vbNo Then Exit Sub

        cmdTrans(0).Visible = True
        cmdTrans(1).Visible = False
        cmdTrans(2).Visible = False
        cmdTrans(0).Enabled = False
  
        Call ClearEntryForm
        Call LockedEntryObject(False)
        
        didRecordEntry = False
        
    End If

Exit Sub
er:
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdTrans_Click()")
End Sub

Function ValidateEntry() As String
    Dim vMsg As String, Qry As String
    Dim i As Integer, Cntr As Integer
    
    vMsg = ""
    
    If txtRefNo.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field ref. no."
    If txtSRPFixPrice.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field SRP fix price."
    If txtWSFixPrice.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field WS fix price."
              
    ValidateEntry = vMsg
    
End Function

Private Sub flxFind_Click()
    
    flxFind.HighLight = flexHighlightAlways
    
    If didRecordEntry = True Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    txtRefNo.Text = ""
    txtRefDate.Text = Format(Now(), "mm/dd/YYYY")
    txtPartNo.Text = flxFind.TextMatrix(flxFind.Row, 1)
    txtDesc.Text = flxFind.TextMatrix(flxFind.Row, 2)
    txtNetofVat.Text = flxFind.TextMatrix(flxFind.Row, 3)
    txtSRP.Text = flxFind.TextMatrix(flxFind.Row, 4)
    txtWS.Text = flxFind.TextMatrix(flxFind.Row, 5)
    SPid = flxFind.TextMatrix(flxFind.Row, 10)
    
    txtSRPProfRatio.Text = ""
    txtSRPProfMar.Text = ""
    txtSRPFixPrice.Text = ""
    txtSRPMaxDis.Text = ""
    
    txtWSProfRatio.Text = ""
    txtWSProfMar.Text = ""
    txtWSFixPrice.Text = ""
    
    fRow = 1
    fCol = 1
    
    DoEvents
    Call LoadSRPHist(SPid)
    DoEvents
    Call LoadWSHist(SPid)
    
    If IsNumeric(flxSRPPriceHis.TextMatrix(1, 4)) = False And flxSRPPriceHis.TextMatrix(1, 4) <> "" Then
        txtSRPProfRatio.Text = Left(Trim(flxSRPPriceHis.TextMatrix(1, 4)), Len((flxSRPPriceHis.TextMatrix(1, 4))) - 1)
    Else
        txtSRPProfRatio.Text = "0"
    End If
        
    txtSRPProfMar.Text = flxSRPPriceHis.TextMatrix(1, 5)

    If IsNumeric(flxWSPriceHis.TextMatrix(1, 4)) = False And flxWSPriceHis.TextMatrix(1, 4) <> "" Then
        txtWSProfRatio.Text = Left(Trim(flxWSPriceHis.TextMatrix(1, 4)), Len((flxWSPriceHis.TextMatrix(1, 4))) - 1)
    Else
        txtWSProfRatio.Text = "0"
    End If
    
    txtWSProfMar.Text = flxWSPriceHis.TextMatrix(1, 5)
    
    txtSRPMaxDis.Text = flxFind.TextMatrix(flxFind.Row, 11)
    
    Check1(0).Value = 0
    Check1(1).Value = 0
    Check1(2).Value = 0
    Check1(3).Value = 0
    Check1(4).Value = 0
    Check1(5).Value = 0
        
    cmdTrans(0).Enabled = True
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub LoadSRPHist(pSPID As Integer)
On Error GoTo er
    Dim tmpRs As New ADODB.Recordset
    Dim SPid As Integer
    
    
    Call modSRP.sp_srp_history(pSPID, tmpRs)
    If tmpRs.State = 0 Then Exit Sub
    
    If tmpRs.RecordCount <= 0 Then
        tmpRs.Close
        flxSRPPriceHis.Rows = 1
        flxSRPPriceHis.Rows = 2
        Exit Sub
    End If
    
'   0       1      2     3      4       5
'DATE, REFERENCE, COST, SRP, RATIO, MARGIN

    flxSRPPriceHis.Rows = 1
    flxSRPPriceHis.Rows = 2

    tmpRs.MoveFirst
           
        While Not tmpRs.EOF
            
            If flxSRPPriceHis.TextMatrix(1, 0) = "" Then
                flxSRPPriceHis.TextMatrix(1, 0) = tmpRs!Date
                flxSRPPriceHis.TextMatrix(1, 1) = tmpRs!REFERENCE
                flxSRPPriceHis.TextMatrix(1, 2) = Format(tmpRs!Cost, "#,##0.00")
                flxSRPPriceHis.TextMatrix(1, 3) = Format(tmpRs!SRP, "#,##0.00")
                flxSRPPriceHis.TextMatrix(1, 4) = tmpRs!PRORATIO
                flxSRPPriceHis.TextMatrix(1, 5) = Format(tmpRs!ProMar, "#,##0.00")
            
            Else
            
                flxSRPPriceHis.AddItem _
                tmpRs!Date & _
                vbTab & tmpRs!REFERENCE & _
                vbTab & Format(tmpRs!Cost, "#,##0.00") & _
                vbTab & Format(tmpRs!SRP, "#,##0.00") & _
                vbTab & tmpRs!PRORATIO & _
                vbTab & Format(tmpRs!ProMar, "#,##0.00")
              
            End If
            
        tmpRs.MoveNext
        
    Wend
    
    If tmpRs.State <> 0 Then tmpRs.Close
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->LoadSRPHist()")
End Sub

Private Sub LoadWSHist(pSPID As Integer)
On Error GoTo er
    Dim tmpRs As New ADODB.Recordset
    
    Call modSRP.sp_ws_history(pSPID, tmpRs)
    If tmpRs.State = 0 Then Exit Sub
    
    If tmpRs.RecordCount <= 0 Then
        tmpRs.Close
        flxWSPriceHis.Rows = 1
        flxWSPriceHis.Rows = 2
        Exit Sub
    End If
    
'   0       1            2          3           4       5
'^DATE |^REF. NO. |^PREV. COST |^PREV. SRP |^RATIO |^MARGIN

    flxWSPriceHis.Rows = 1
    flxWSPriceHis.Rows = 2
      
    tmpRs.MoveFirst
           
        While Not tmpRs.EOF
            
            If flxWSPriceHis.TextMatrix(1, 0) = "" Then
                flxWSPriceHis.TextMatrix(1, 0) = tmpRs!FDate
                flxWSPriceHis.TextMatrix(1, 1) = tmpRs!RefNo
                flxWSPriceHis.TextMatrix(1, 2) = Format(tmpRs!Cost, "#,##0.00")
                flxWSPriceHis.TextMatrix(1, 3) = Format(tmpRs!DP, "#,##0.00")
                flxWSPriceHis.TextMatrix(1, 4) = tmpRs!PRORATIO
                flxWSPriceHis.TextMatrix(1, 5) = Format(tmpRs!ProMar, "#,##0.00")

            Else
            
                flxWSPriceHis.AddItem _
                tmpRs!FDate & _
                vbTab & tmpRs!RefNo & _
                vbTab & Format(tmpRs!Cost, "#,##0.00") & _
                vbTab & Format(tmpRs!DP, "#,##0.00") & _
                vbTab & tmpRs!PRORATIO & _
                vbTab & Format(tmpRs!ProMar, "#,##0.00")
              
            End If
            
        tmpRs.MoveNext
        
    Wend
    
    If tmpRs.State <> 0 Then tmpRs.Close
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->LoadWSHist()")
End Sub

Private Sub flxFind_EnterCell()
    Call flxFind_Click
End Sub

Private Sub Form_Load()
On Error GoTo er
Dim tmpRs As ADODB.Recordset

    DoEvents
    Call modSRP.sp_supplier_list(tmpRs)
    If tmpRs.State = 1 Then
        If tmpRs.RecordCount > 0 Then
            Call modObjTools.fillArray(AllArrSupplier, tmpRs)
        End If
    End If
    If tmpRs.State <> 0 Then tmpRs.Close
    
    DoEvents
    Call modObjTools.addToListObjectUsingArrayWithCols(cboSupplier, arrSupplier, AllArrSupplier, 1, 1)
    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->Form_Load()")
End Sub



Private Sub RXResizer1_ResizeComplete()
On Error Resume Next

        DoEvents
        Call flxFindColsChanger
        DoEvents
        Call flxHistColsChanger
        
'        Call LoadRecord
        
        Screen.MousePointer = vbHourglass
        tmrRefresh.Enabled = True
        DoEvents
        frm_OnProcess2.OwnerForm = "SRP Updater"
        Call Load(frm_OnProcess2)
        frm_OnProcess2.Show vbModal
    
End Sub



Private Sub tmrRefresh_Timer()
On Error GoTo er

    tmrRefresh.Enabled = False
    Screen.MousePointer = vbHourglass
    frmPriceUpdate.Enabled = False
    
    DoEvents
    DoEvents
    Call LoadRecordBySup(SupID)
    DoEvents
    DoEvents
    
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    
    DoEvents
    frmPriceUpdate.Enabled = True
    Screen.MousePointer = vbDefault
    
Exit Sub
er:
    frmPriceUpdate.Enabled = True
    Screen.MousePointer = vbDefault
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    Call modObjTools.ErrorMessage(Err.Description, "Error->tmrRefresh_Timer()")
End Sub


Private Sub txtFindParts_Change()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFindParts_Click()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFindParts_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
    Dim i As Integer, j As Integer
    
        If lblFindParts.Caption = "FIND PART DESCRIPTION:" Then

            If txtFindParts.Text = "" Then Exit Sub
            With flxFind
                For i = fRow To .Rows - 1
                    For j = fCol To 2
                        If InStr(.TextMatrix(i, j), txtFindParts.Text) > 0 Then
                            .Col = 1: .Row = i
                            .ColSel = 9: .RowSel = i
                            .TopRow = i
                            .HighLight = flexHighlightAlways
                            fRow = IIf(j = TCol - 1, i + 1, i)
                            fCol = IIf(j = TCol - 1, 1, j + 1)
                            ActRow = i
                            
                            Call flxFind_Click
                    
                            Exit Sub
                        End If
                    
                    Next j
                    fCol = 1
                Next i
            End With
    
            MsgBox "REACH END OF ROW!", vbInformation + vbOKOnly, "Find"
            fRow = 1
            fCol = 1
    
        Else
        
            If txtFindParts.Text = "" Then Exit Sub
            With flxFind
                For i = fRow To .Rows - 1
                    For j = fCol To 1
                        If .TextMatrix(i, j) = txtFindParts.Text Then
                            .Col = 1: .Row = i
                            .ColSel = 9: .RowSel = i
                            .TopRow = i
                            .HighLight = flexHighlightAlways
                            fRow = IIf(j = TCol - 1, i + 1, i)
                            fCol = IIf(j = TCol - 1, 1, j + 1)
                            ActRow = i
                                
                            Call flxFind_Click
                                                
                            Exit Sub
                        End If
                    
                    Next j
                    fCol = 1
                Next i
            End With
            MsgBox "REACH END OF ROW!", vbInformation + vbOKOnly, "Find"
            fRow = 1
            fCol = 1
            
        End If
        
    End If
    
End Sub

Private Sub txtSRPFixPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtSRPFixPrice, KeyAscii, NumbersOnly)
    
    If KeyAscii = 13 Then
            Dim ProMar As Double
            Dim NewCost As Double
            Dim NewSRP As Double
            Dim Ratio As Double
            
            NewSRP = Val(Replace(Trim(txtSRPFixPrice.Text), ",", ""))
            NewCost = Format(txtNetofVat.Text, "#,#00")
            NewCost = Val(Replace(NewCost, ",", ""))
            NewCost = NewCost * 1.12
            NewCost = Format(NewCost, "#,#00")
            
            If NewSRP < NewCost Then
                MsgBox "New SRP is less than new cost!", vbExclamation, Me.Caption
                Exit Sub
            Else
                ProMar = NewSRP - NewCost
            End If
            
            Ratio = ProMar / NewSRP
            Ratio = Ratio * 100
            
            txtSRPFixPrice.Text = Format(NewSRP, "#,#00")
            txtSRPProfMar.Text = Format(ProMar, "#,#00")
            
            If Ratio > 100 Then
                Ratio = 100
            End If
            
            txtSRPProfRatio.Text = Format(Ratio, "#,#00.00")
    
    End If
    
End Sub



Private Sub txtSRPMaxDis_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtSRPMaxDis, KeyAscii, NumbersOnly)
End Sub

Private Sub txtSRPProfMar_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtSRPProfMar, KeyAscii, NumbersOnly)
    
    If KeyAscii = 13 Then
            Dim NewCost As Double
            Dim NewSRP As Double
            Dim Ratio As Double
            
            NewCost = Format(txtNetofVat.Text, "#,#00")
            NewCost = Val(Replace(NewCost, ",", ""))
            NewCost = NewCost * 1.12
            NewCost = Format(NewCost, "#,#00")
            
            NewSRP = NewCost + txtSRPProfMar.Text
            
            Ratio = txtSRPProfMar.Text / NewSRP
            Ratio = Ratio * 100
            
            txtSRPFixPrice.Text = Format(NewSRP, "#,#00")
            txtSRPProfMar.Text = Format(txtSRPProfMar.Text, "#,#00")
            
            If Ratio > 100 Then
                Ratio = 100
            End If
            
            txtSRPProfRatio.Text = Format(Ratio, "#,#00.00")
    End If
    
End Sub

Private Sub txtSRPProfRatio_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtSRPProfRatio, KeyAscii, NumbersWithDecimalPoint)
    
    If KeyAscii = 13 Then
            Dim NewCost As Double
            Dim NewSRP As Double
            Dim Ratio As Double
            Dim ProfMar As Double
            

            NewCost = Val(Replace(txtNetofVat, ",", ""))
            NewCost = NewCost * 1.12
            NewCost = Format(NewCost, "#,#00")
            
            Ratio = txtSRPProfRatio.Text
            
            If Ratio > 100 Then
                Ratio = 100
            End If
            
            Ratio = (100 - Ratio) / 100
            NewSRP = NewCost / Ratio
            ProfMar = NewSRP - NewCost
            
            txtSRPFixPrice.Text = Format(NewSRP, "#,#00")
            txtSRPProfMar.Text = Format(ProfMar, "#,#00")
            txtSRPProfRatio.Text = Format(txtSRPProfRatio.Text, "#,#00")

    End If
End Sub


Private Sub txtWSFixPrice_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtWSFixPrice, KeyAscii, NumbersOnly)
    
    If KeyAscii = 13 Then
            Dim ProMar As Double
            Dim NewCost As Double
            Dim NewWSP As Double
            Dim Ratio As Double
            
            NewWSP = Val(Replace(Trim(txtWSFixPrice.Text), ",", ""))
            NewCost = Val(Replace(Trim(txtNetofVat.Text), ",", ""))
            NewCost = NewCost * 1.12
            NewCost = Format(NewCost, "#,#00")
            
            If NewWSP < NewCost Then
                MsgBox "New WSP is less than new cost!", vbExclamation, Me.Caption
                Exit Sub
            Else
                ProMar = NewWSP - NewCost
            End If
            
            Ratio = ProMar / NewWSP
            Ratio = Ratio * 100
            
            txtWSFixPrice.Text = Format(NewWSP, "#,#00")
            txtWSProfMar.Text = Format(ProMar, "#,#00")
            
            If Ratio > 100 Then
                Ratio = 100
            End If
            
            txtWSProfRatio.Text = Format(Ratio, "#,#00.00")
    
    End If
End Sub

Private Sub txtWSProfMar_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtWSProfMar, KeyAscii, NumbersOnly)
    
    If KeyAscii = 13 Then
            Dim NewCost As Double
            Dim NewWSP As Double
            Dim Ratio As Double
            
            NewCost = Format(txtNetofVat.Text, "#,#00")
            NewCost = Val(Replace(NewCost, ",", ""))
            NewCost = NewCost * 1.12
            NewCost = Format(NewCost, "#,#00")
            
            NewWSP = NewCost + txtWSProfMar.Text
            
            Ratio = txtWSProfMar.Text / NewWSP
            Ratio = Ratio * 100
            
            txtWSFixPrice.Text = Format(NewWSP, "#,#00")
            txtWSProfMar.Text = Format(txtWSProfMar, "#,#00")
            
            If Ratio > 100 Then
                Ratio = 100
            End If
            
            txtWSProfRatio.Text = Format(Ratio, "#,#00.00")
    End If
    
End Sub

Private Sub txtWSProfRatio_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtWSProfRatio, KeyAscii, NumbersWithDecimalPoint)
        
    If KeyAscii = 13 Then
        Dim NewCost As Double
        Dim NewWSP As Double
        Dim Ratio As Double
        Dim ProfMar As Double
        
        NewCost = Val(Replace(txtNetofVat, ",", ""))
        NewCost = NewCost * 1.12
        NewCost = Format(NewCost, "#,#00")
        
        Ratio = txtWSProfRatio.Text
        
        If Ratio > 100 Then
            Ratio = 100
        End If
        
        Ratio = (100 - Ratio) / 100
        NewWSP = NewCost / Ratio
        ProfMar = NewWSP - NewCost
        
        txtWSFixPrice.Text = Format(NewWSP, "#,#00")
        txtWSProfMar.Text = Format(ProfMar, "#,#00")
        txtWSProfRatio.Text = Format(txtWSProfRatio.Text, "#,#00.00")
    End If
End Sub

Private Sub VScroll1_Change()
    VScroll1.Max = 2
    
    If VScroll1.Value = 0 Then
        lblFindParts.Caption = "FIND PART DESCRIPTION:"
    ElseIf VScroll1.Value = 1 Then
        lblFindParts.Caption = "FIND PART NUMBER:"
    ElseIf VScroll1.Value = 2 Then
        VScroll1.Value = 0
    End If
    
End Sub


Sub flxFindColsChanger()
On Error Resume Next
    With flxFind
        .Clear
        .Rows = 2
        .Cols = 12
        .FixedCols = 1
        
        '               0       1             2            3           4      5       6      7        8              9
        .FormatString = "|<PART NUMBER |<DESCRIPTION  |^NET OF VAT  |^SRP  |^WSP  |^QTY. |^6MSQ |^BACKORDER |<APPLICABLE MODEL"
        
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.15  'PART NUMBER
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.27  'DESCRIPTION
        
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.08  'NET OF VAT
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.08  'SRP
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.08  'WSP
        
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.07  'QTY.
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.07  '6MSQ
        .ColWidth(8) = (.Width - .ColWidth(0)) * 0.07  'BACKORDER
        .ColWidth(9) = (.Width - .ColWidth(0)) * 0.1   'APPLICABLE MODEL
        .ColWidth(10) = 0 'SPAREPARTID
        .ColWidth(11) = 0 'MAX DISCOUNT
    End With
    
End Sub


Sub flxHistColsChanger()
On Error Resume Next
    With flxSRPPriceHis
         .Clear
         .Rows = 2
         .Cols = 6
         '                  0       1           2            3          4       5
         .FormatString = "^DATE |^REF. NO. |^PREV. COST |^PREV. SRP "
         
         
         .ColWidth(0) = (.Width - .ColWidth(0)) * 0.26 'DATE
         .ColWidth(1) = (.Width - .ColWidth(0)) * 0.4  'REF. NO.
         .ColWidth(2) = (.Width - .ColWidth(0)) * 0.26 'PREV. COST
         .ColWidth(3) = (.Width - .ColWidth(0)) * 0.26 'PREV. SRP
         .ColWidth(4) = 0 'RATIO
         .ColWidth(5) = 0 'MARGIN
    
    End With
    
    
    With flxWSPriceHis
         .Clear
         .Rows = 2
         .Cols = 6
         '                  0       1           2            3          4       5
         .FormatString = "^DATE |^REF. NO. |^PREV. COST |^PREV. SRP "
         
         .ColWidth(0) = (.Width - .ColWidth(0)) * 0.26  'DATE
         .ColWidth(1) = (.Width - .ColWidth(0)) * 0.4  'REF. NO.
         .ColWidth(2) = (.Width - .ColWidth(0)) * 0.26 'PREV. COST
         .ColWidth(3) = (.Width - .ColWidth(0)) * 0.26 'PREV. SRP
         .ColWidth(4) = 0 'RATIO
         .ColWidth(5) = 0 'MARGIN
    
    End With

End Sub

Public Sub LoadRecord()
On Error GoTo er
    Dim tmpRs As New ADODB.Recordset
    
    Call modSRP.sp_load_record(tmpRs)
    If tmpRs.State = 0 Then Exit Sub
    
    If tmpRs.RecordCount <= 0 Then
        tmpRs.Close
        Exit Sub
    End If
    
    Call flxFindColsChanger
    
'       1             2            3          4      5       6      7        8              9
'|^PART NUMBER |^DESCRIPTION  |^NET OF VAT  |^SRP  |^WSP  |^QTY. |^6MSQ |^BACKORDER |^APPLICABLE MODEL "
'SPAREPARTID, PARTNUMBER, DESCRIPTION, COST, STOREPRICE, DLRPRICE, QTYONHAND, BACKORDER, APPLICMODEL
    tmpRs.MoveFirst
           
        While Not tmpRs.EOF
        
            If flxFind.TextMatrix(1, 1) = "" Then
                flxFind.TextMatrix(1, 1) = tmpRs!PartNumber
                flxFind.TextMatrix(1, 2) = tmpRs!Description
                flxFind.TextMatrix(1, 3) = Format(tmpRs!Cost, "#,##0.00")
                flxFind.TextMatrix(1, 4) = Format(tmpRs!StorePrice, "#,##0")
                flxFind.TextMatrix(1, 5) = Format(tmpRs!DlrPrice, "#,##0")
                flxFind.TextMatrix(1, 6) = tmpRs!QtyOnHand
                flxFind.TextMatrix(1, 7) = tmpRs!SixMSQ
                flxFind.TextMatrix(1, 8) = tmpRs!BACKORDER
                flxFind.TextMatrix(1, 9) = tmpRs!ApplicModel
                flxFind.TextMatrix(1, 10) = tmpRs!SPAREPARTID
                flxFind.TextMatrix(1, 11) = tmpRs!MaxDiscount

            Else
            
                flxFind.AddItem _
                vbTab & tmpRs!PartNumber & _
                vbTab & tmpRs!Description & _
                vbTab & Format(tmpRs!Cost, "#,##0.00") & _
                vbTab & Format(tmpRs!StorePrice, "#,##0") & _
                vbTab & Format(tmpRs!DlrPrice, "#,##0") & _
                vbTab & tmpRs!QtyOnHand & _
                vbTab & tmpRs!SixMSQ & _
                vbTab & tmpRs!BACKORDER & _
                vbTab & tmpRs!ApplicModel & _
                vbTab & tmpRs!SPAREPARTID & _
                vbTab & tmpRs!MaxDiscount
                
            End If
            
        tmpRs.MoveNext
        
    Wend
    
    If tmpRs.State <> 0 Then tmpRs.Close
    
    flxFind.HighLight = flexHighlightNever
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->LoadRecord()")
End Sub

Public Sub LoadRecordBySup(pSupID As Integer)
On Error GoTo er
    Dim tmpRs As New ADODB.Recordset
    
    Call modSRP.sp_load_record(pSupID, tmpRs)
    If tmpRs.State = 0 Then Exit Sub
    
    If tmpRs.RecordCount <= 0 Then
        tmpRs.Close
        Exit Sub
    End If
    
    Call flxFindColsChanger
    
    frm_OnProcess2.ProgLoading.Value = 0
    
'       1             2            3          4      5       6      7        8              9
'|^PART NUMBER |^DESCRIPTION  |^NET OF VAT  |^SRP  |^WSP  |^QTY. |^6MSQ |^BACKORDER |^APPLICABLE MODEL "
'SPAREPARTID, PARTNUMBER, DESCRIPTION, COST, STOREPRICE, DLRPRICE, QTYONHAND, BACKORDER, APPLICMODEL
    tmpRs.MoveFirst
           
        While Not tmpRs.EOF
        
            frm_OnProcess2.ProgLoading.Max = tmpRs.RecordCount
        
            If frm_OnProcess2.ProgLoading.Value < frm_OnProcess2.ProgLoading.Max Then
                frm_OnProcess2.ProgLoading.Value = frm_OnProcess2.ProgLoading.Value + 1
            End If
        
            If flxFind.TextMatrix(1, 1) = "" Then
                flxFind.TextMatrix(1, 1) = tmpRs!PartNumber
                flxFind.TextMatrix(1, 2) = tmpRs!Description
                flxFind.TextMatrix(1, 3) = Format(tmpRs!Cost, "#,##0.00")
                flxFind.TextMatrix(1, 4) = Format(tmpRs!StorePrice, "#,##0")
                flxFind.TextMatrix(1, 5) = Format(tmpRs!DlrPrice, "#,##0")
                flxFind.TextMatrix(1, 6) = tmpRs!QtyOnHand
                flxFind.TextMatrix(1, 7) = tmpRs!SixMSQ
                flxFind.TextMatrix(1, 8) = tmpRs!BACKORDER
                flxFind.TextMatrix(1, 9) = tmpRs!ApplicModel
                flxFind.TextMatrix(1, 10) = tmpRs!SPAREPARTID
                flxFind.TextMatrix(1, 11) = tmpRs!MaxDiscount

            Else
            
                flxFind.AddItem _
                vbTab & tmpRs!PartNumber & _
                vbTab & tmpRs!Description & _
                vbTab & Format(tmpRs!Cost, "#,##0.00") & _
                vbTab & Format(tmpRs!StorePrice, "#,##0") & _
                vbTab & Format(tmpRs!DlrPrice, "#,##0") & _
                vbTab & tmpRs!QtyOnHand & _
                vbTab & tmpRs!SixMSQ & _
                vbTab & tmpRs!BACKORDER & _
                vbTab & tmpRs!ApplicModel & _
                vbTab & tmpRs!SPAREPARTID & _
                vbTab & tmpRs!MaxDiscount
                
            End If
            
        tmpRs.MoveNext
        
    Wend
    
    If tmpRs.State <> 0 Then tmpRs.Close
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->LoadRecordBySup()")
End Sub


Sub ClearEntryForm()
   
    txtRefDate.Text = ""
    txtNetofVat.Text = ""
   
    txtPartNo.Text = ""
    txtDesc.Text = ""
    
    txtSRP.Text = ""
    txtSRPProfRatio.Text = ""
    txtSRPProfMar.Text = ""
    txtSRPFixPrice.Text = ""
    txtSRPMaxDis.Text = ""
    
    txtWSProfRatio.Text = ""
    txtWSProfMar.Text = ""
    txtWS.Text = ""
    txtWSFixPrice.Text = ""
    
    flxSRPPriceHis.Rows = 1
    flxSRPPriceHis.Rows = 2
    
    flxWSPriceHis.Rows = 1
    flxWSPriceHis.Rows = 2
    
    Check1(0).Value = 0
    Check1(1).Value = 0
    Check1(2).Value = 0
    Check1(3).Value = 0
    Check1(4).Value = 0
    Check1(5).Value = 0
                              
End Sub


Sub LockedEntryObject(DataEntry As Boolean)

    DataEntry = Not DataEntry
        
        txtRefNo.Locked = DataEntry
        txtSRPMaxDis.Locked = DataEntry
        txtRefDate.Locked = DataEntry
        txtNetofVat.Locked = DataEntry
        txtPartNo.Locked = DataEntry
        txtDesc.Locked = DataEntry
        txtSRP.Locked = DataEntry
        txtSRPProfMar.Locked = DataEntry
        txtSRPProfRatio.Locked = DataEntry
        txtSRPFixPrice.Locked = DataEntry
        txtSRPMaxDis.Locked = DataEntry
        txtWSProfMar.Locked = DataEntry
        txtWSProfRatio.Locked = DataEntry
        txtWSFixPrice.Locked = DataEntry
                                                                                                                                                                                                                                 
    If txtRefNo.Locked Then
        
        txtRefNo.BackColor = txtHoldGrayBackColor.BackColor
        txtSRPMaxDis.BackColor = txtHoldGrayBackColor.BackColor
        txtRefDate.BackColor = txtHoldGrayBackColor.BackColor
        txtNetofVat.BackColor = txtHoldGrayBackColor.BackColor
        txtPartNo.BackColor = txtHoldGrayBackColor.BackColor
        txtDesc.BackColor = txtHoldGrayBackColor.BackColor
        txtSRP.BackColor = txtHoldGrayBackColor.BackColor
        txtSRPProfMar.BackColor = txtHoldGrayBackColor.BackColor
        txtSRPProfRatio.BackColor = txtHoldGrayBackColor.BackColor
        txtSRPFixPrice.BackColor = txtHoldGrayBackColor.BackColor
        txtSRPMaxDis.BackColor = txtHoldGrayBackColor.BackColor
        txtWSProfMar.BackColor = txtHoldGrayBackColor.BackColor
        txtWSProfRatio.BackColor = txtHoldGrayBackColor.BackColor
        txtWSFixPrice.BackColor = txtHoldGrayBackColor.BackColor
        
    Else
        
        txtRefNo.BackColor = txtHoldWhiteBackColor.BackColor
        txtSRPMaxDis.BackColor = txtHoldWhiteBackColor.BackColor
        txtRefDate.BackColor = txtHoldWhiteBackColor.BackColor
        txtNetofVat.BackColor = txtHoldWhiteBackColor.BackColor
        txtPartNo.BackColor = txtHoldWhiteBackColor.BackColor
        txtDesc.BackColor = txtHoldWhiteBackColor.BackColor
        txtSRP.BackColor = txtHoldWhiteBackColor.BackColor
        txtSRPProfMar.BackColor = txtHoldWhiteBackColor.BackColor
        txtSRPProfRatio.BackColor = txtHoldWhiteBackColor.BackColor
        txtSRPFixPrice.BackColor = txtHoldWhiteBackColor.BackColor
        txtSRPMaxDis.BackColor = txtHoldWhiteBackColor.BackColor
        txtWSProfMar.BackColor = txtHoldWhiteBackColor.BackColor
        txtWSProfRatio.BackColor = txtHoldWhiteBackColor.BackColor
        txtWSFixPrice.BackColor = txtHoldWhiteBackColor.BackColor
        
    End If

End Sub
