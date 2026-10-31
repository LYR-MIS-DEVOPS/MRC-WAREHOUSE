VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmDataEntry 
   Caption         =   "Parts Catalog Data Entry"
   ClientHeight    =   9150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   10680
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtCategoryID 
      Height          =   285
      Left            =   12000
      TabIndex        =   30
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtCatalogID 
      Height          =   285
      Left            =   11520
      TabIndex        =   29
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtRefNo 
      Height          =   285
      Left            =   11400
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPartNo 
      Height          =   1815
      Left            =   11040
      TabIndex        =   26
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   3201
      _Version        =   393216
      FixedRows       =   0
      GridLines       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   11040
      Top             =   3600
   End
   Begin VB.TextBox txtPiece 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   11760
      TabIndex        =   11
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtPartNo 
      Height          =   285
      Left            =   11040
      TabIndex        =   9
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   6735
      Left            =   120
      TabIndex        =   22
      Tag             =   "XY"
      Top             =   2280
      Width           =   10455
      Begin MSFlexGridLib.MSFlexGrid flxParts 
         Height          =   6015
         Left            =   120
         TabIndex        =   8
         Tag             =   "XY"
         Top             =   600
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   10610
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         FormatString    =   "NO. |<PARTNUMBER |<DESCRIPTION |^REF NO.  |<PRICE |^PIECE"
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
      Begin VB.CommandButton cmdAltD 
         Caption         =   "&D"
         Height          =   375
         Left            =   1920
         TabIndex        =   31
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton cmdAltA 
         Caption         =   "&A"
         Height          =   375
         Left            =   1920
         TabIndex        =   32
         Top             =   1200
         Width           =   375
      End
      Begin lvButton.lvButtons_H cmdInsX 
         Height          =   375
         Left            =   9600
         TabIndex        =   35
         Tag             =   "XY"
         ToolTipText     =   "Insert Existing Parts [F12]"
         Top             =   150
         Width           =   495
         _ExtentX        =   873
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
         ImgAlign        =   4
         Image           =   "frmDataEntry.frx":0000
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   14
      Tag             =   "XY"
      Top             =   600
      Width           =   10455
      Begin VB.ComboBox txtVersion 
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
         Left            =   1320
         TabIndex        =   34
         Tag             =   "XY"
         Top             =   1100
         Width           =   3375
      End
      Begin VB.ComboBox txtModel 
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
         Left            =   1320
         TabIndex        =   33
         Tag             =   "XY"
         Top             =   645
         Width           =   3375
      End
      Begin VB.ComboBox cboBlock 
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
         Left            =   7680
         TabIndex        =   7
         Tag             =   "XY"
         Top             =   1100
         Width           =   2535
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   375
         Left            =   7680
         TabIndex        =   21
         Tag             =   "XY"
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
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
      Begin VB.ComboBox cboGroup 
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
         Left            =   7680
         TabIndex        =   6
         Tag             =   "XY"
         Top             =   645
         Width           =   2535
      End
      Begin VB.ComboBox cboBrand 
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
         Left            =   1320
         TabIndex        =   5
         Tag             =   "XY"
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label lblDate 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE:"
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
         Left            =   6720
         TabIndex        =   20
         Tag             =   "XY"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblBlock 
         Alignment       =   1  'Right Justify
         Caption         =   "BLOCK:"
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
         Left            =   6720
         TabIndex        =   19
         Tag             =   "XY"
         Top             =   1155
         Width           =   855
      End
      Begin VB.Label lblGroup 
         Alignment       =   1  'Right Justify
         Caption         =   "GROUP:"
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
         Left            =   6600
         TabIndex        =   18
         Tag             =   "XY"
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblVersion 
         Caption         =   "VERSION:"
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
         Tag             =   "XY"
         Top             =   1155
         Width           =   975
      End
      Begin VB.Label lblModel 
         Caption         =   "MODEL:"
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
         TabIndex        =   16
         Tag             =   "XY"
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblBrand 
         Caption         =   "BRAND:"
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
         TabIndex        =   15
         Tag             =   "XY"
         Top             =   255
         Width           =   855
      End
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   11040
      TabIndex        =   13
      Text            =   "HoldBackColor"
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   11040
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
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
      Image           =   "frmDataEntry.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   2
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "Su&mmary"
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
      Image           =   "frmDataEntry.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4620
      TabIndex        =   3
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
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
      Image           =   "frmDataEntry.frx":10CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6120
      TabIndex        =   4
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "E&xit"
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
      Image           =   "frmDataEntry.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   7620
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
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
      Image           =   "frmDataEntry.frx":1C02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   9120
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
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
      Image           =   "frmDataEntry.frx":219C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1620
      TabIndex        =   1
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
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
      Image           =   "frmDataEntry.frx":2736
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   23
      Tag             =   "XY"
      Top             =   9240
      Width           =   10455
      Begin VB.TextBox txtUpdatedBy 
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
         Left            =   1680
         TabIndex        =   25
         Tag             =   "XY"
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label lblUpdatedBy 
         Caption         =   "UPDATED BY:"
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
         TabIndex        =   24
         Tag             =   "XY"
         Top             =   300
         Width           =   1455
      End
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   10200
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   10920
      m_MDI_Height    =   9720
   End
   Begin VB.Menu mnuStatus 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add Row"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete Row"
      End
   End
End
Attribute VB_Name = "frmDataEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ActRow As Integer, ActCol As Integer, catID As Integer, blockID As Integer
Dim rsMain As New ADODB.Recordset
Dim rsParts As New ADODB.Recordset
Dim Qry As String, formAction As String
Dim txtstr As String
Dim ans As String, MyErr As String, myMsg As String
Dim arrGROUP() As String
Dim arrBRAND() As String
Dim arrBLOCK() As String
Dim arrMODEL() As String
Dim arrVERSION() As String

Dim strPartial As String, strTotal As String, brandName As String
Dim StatSrch As Boolean, ProperExit As Boolean, PrtSrch As Boolean, PrtBln As Boolean
Dim i As Long, j As Long
Dim intLoopCount As Integer, chkdupCount As Integer

Public LockFlx As Boolean

Public Sub ShowRecord(catID As Long, categoryID As Long)
Dim rsHeader As New ADODB.Recordset
Dim rsDetails As New ADODB.Recordset

On Error GoTo er

    txtCatalogID.Text = catID 'Used for Edit Verification
    txtCategoryID.Text = categoryID 'Used for Edit Verification
    
    Screen.MousePointer = vbHourglass
    
    DoEvents
    Qry = "call partscatalog.sp_partscatalog_getHeader('" & catID & "', '" & categoryID & "')"
    
    With rsHeader
        .Open Qry, conPartsCat, adOpenStatic, adLockReadOnly
        If .State <> 1 Then Exit Sub
            
        If .RecordCount <= 0 Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If .RecordCount > 0 Then
            .MoveFirst
            txtDate.Text = Format$(!DateCreated, "mm/dd/YYYY")
            cboBRAND.Text = IIf(IsNull(!BRAND), "", !BRAND)
            txtModel.Text = IIf(IsNull(!MODEL), "", !MODEL)
            txtVersion.Text = IIf(IsNull(!Version), "", !Version)
            cboGROUP.Text = IIf(IsNull(!Group), "", !Group)
            cboBlock.Text = IIf(IsNull(!BLOCK), "", !BLOCK)
            txtUpdatedBy.Text = IIf(IsNull(!UPDATEDBY), "", !UPDATEDBY)
        End If
    
        If .State <> 0 Then
            Screen.MousePointer = vbDefault
            .Close
        End If
    End With


   DoEvents
    Qry = "call partscatalog.sp_partscatalog_getDetails('" & categoryID & "')"
    
    With rsHeader
        .Open Qry, conPartsCat, adOpenStatic, adLockReadOnly
        If .State <> 1 Then Exit Sub
    
        If .RecordCount <= 0 Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        If .RecordCount > 0 Then
            .MoveFirst
            flxParts.Rows = .RecordCount + 1
            
            For i = 1 To flxParts.Rows - 1
                flxParts.TextMatrix(i, 0) = i
                flxParts.TextMatrix(i, 1) = IIf(IsNull(!PartNumber), "", !PartNumber)
                flxParts.TextMatrix(i, 2) = IIf(IsNull(!Description), "", !Description)
                flxParts.TextMatrix(i, 3) = IIf(IsNull(!RefNo), "", !RefNo)
                flxParts.TextMatrix(i, 4) = IIf(IsNull(Format(!Price, "##,#00.00")), "", Format(!Price, "##,#00.00"))
                flxParts.TextMatrix(i, 5) = IIf(IsNull(!PIECE), "", !PIECE)
                flxParts.TextMatrix(i, 6) = IIf(IsNull(!SPAREPARTID), "", !SPAREPARTID)
                .MoveNext
            Next i
        End If
    
        If .State <> 0 Then
            Screen.MousePointer = vbDefault
            .Close
        End If
    End With
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->dataEntryShowRecord()")
End Sub


'Public Sub CheckDuplicate()
'On Error GoTo er
''Verify the Part Number if already exist
'chkdupCount = 0
'If ActRow <> 1 Then
'    With flxParts
'
'        For intLoopCount = 1 To .Rows - 1
'        txtstr = flxParts.TextMatrix(intLoopCount, 1)
'
'        If txtstr = txtPartNo.Text Then 'flxPARTS.TextMatrix(actrow, 1) Then
'            MsgBox "Part Number is already exist in row " & " " & intLoopCount & "" & Err.Description, vbExclamation + vbOKOnly, "Warning"
'            .TextMatrix(.Row, 1) = ""
'            .TextMatrix(.Row, 2) = ""
'            .TextMatrix(.Row, 3) = ""
'            .TextMatrix(.Row, 4) = ""
'            .TextMatrix(.Row, 5) = ""
'            .TextMatrix(.Row, 6) = ""
'            txtPartNo.Text = ""
'            chkdupCount = chkdupCount + 1
'            Exit For
'        End If
'
'        Next intLoopCount
'
'    End With
'End If
'Exit Sub
'er:
'    DoEvents
'    Call modObjTools.ErrorMessage(Err.Description, "Error->CheckDuplicate()")
'End Sub


Public Sub textVisible(textBoxVisible As Boolean)

    txtPartNo.Visible = textBoxVisible
    flxPartNo.Visible = textBoxVisible
    txtRefNo.Visible = textBoxVisible
    txtPiece.Visible = textBoxVisible
    
    textBoxVisible = True

End Sub

Private Sub SetButtonVisibility(DataEntry As Boolean)
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->VIEW SUMM
'3 ->EXIT
'4 ->SAVE
'5 ->CANCEL

    Dim i As Integer

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

Function ValidationPartsCatalog() As String
On Error GoTo er
   
    MyErr = ""
    If flxParts.TextMatrix(flxParts.Rows - 1, 1) = "" Then
        If flxParts.Rows = 2 Then
            flxParts.Rows = 1
            flxParts.Rows = 2
        Else
            flxParts.RemoveItem flxParts.Rows - 1
        End If
    End If
    
    If flxParts.TextMatrix(flxParts.Rows - 1, 1) = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on Last Row of Parts Matrix"

    For i = 1 To flxParts.Rows - 1
        If flxParts.TextMatrix(i, 3) = "" Then MyErr = MyErr & vbCrLf & "There is no REF NO. @ Row " & i & "!"
        If flxParts.TextMatrix(i, 5) = "" Then MyErr = MyErr & vbCrLf & "There is no PIECE @ Row " & i & "!"
    Next i
    
    ValidationPartsCatalog = MyErr

er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->ValidationPartsCatalog()")
End Function

Public Sub CreateTempTable()
On Error GoTo er
        conPartsCat.Execute "Drop Temporary Table IF  Exists partscatalog.TMP_PARTSCATALOG_DTLS;"

         DoEvents
       Qry = "Create Temporary Table IF Not Exists partscatalog.TMP_PARTSCATALOG_DTLS( " & _
        "   REFNO Int(11),              " & _
        "   PARTNUMBER Varchar(50),     " & _
        "   DESCRIPTION Varchar(100),   " & _
        "   PRICE Float,                " & _
        "   PIECE INT(11),              " & _
        "   SPAREPARTID INT(11)         " & _
        ");"

        conPartsCat.Execute Qry

      ' "   PMATRIXID Int(11),           "
        Qry = "insert into partscatalog.TMP_PARTSCATALOG_DTLS(REFNO , PARTNUMBER, DESCRIPTION, PRICE, PIECE, SPAREPARTID) VALUES "
        
        For i = 1 To flxParts.Rows - 1
            If i > 1 Then Qry = Qry & ","
            Qry = Qry & "('" & Val(flxParts.TextMatrix(i, 3)) & "', '" & flxParts.TextMatrix(i, 1) & "', '" & flxParts.TextMatrix(i, 2) & "', '" & Format$(flxParts.TextMatrix(i, 4), "##0.00") & "', '" & Val(flxParts.TextMatrix(i, 5)) & "', '" & Val(flxParts.TextMatrix(i, 6)) & "')"
        Next
        'InputBox Qry, Qry, Qry
        Call conPartsCat.Execute(Qry)

Exit Sub
er:
 Call modObjTools.ErrorMessage(Err.Description, "Error->CreateTempTable()")
End Sub

Function ValidateEntry() As Boolean
  ValidateEntry = False
    
    If Trim(cboBRAND.Text) = "" Then
        MsgBox "BRAND IS REQUIRED!", vbExclamation, "Validation"
        cboBRAND.SetFocus
        Exit Function
    End If
    
    If Trim(txtModel.Text) = "" Then
        MsgBox "MODEL IS REQUIRED!", vbExclamation, "Validation"
        txtModel.SetFocus
        Exit Function
    End If
        
    If Trim(txtVersion.Text) = "" Then
        MsgBox "VERSION IS REQUIRED!", vbExclamation, "Validation"
        txtVersion.SetFocus
        Exit Function
    End If
    
    If Trim(cboGROUP.Text) = "" Then
        MsgBox "GROUP IS REQUIRED!", vbExclamation, "Validation"
        cboGROUP.SetFocus
        Exit Function
    End If
    
    If Trim(cboBlock.Text) = "" Then
        MsgBox "BLOCK IS REQUIRED!", vbExclamation, "Validation"
        cboBlock.SetFocus
        Exit Function
    End If
    
    If Trim(txtUpdatedBy.Text) = "" Then
        MsgBox "UPDATED BY IS REQUIRED!", vbExclamation, "Validation"
        txtUpdatedBy.SetFocus
        Exit Function
    End If
    
    ValidateEntry = True

End Function


Sub Clearform()
    flxParts.Rows = 1
    flxParts.Rows = 2
   
    cboBRAND.Text = ""
    txtModel.Text = ""
    txtVersion.Text = ""
    cboGROUP.Text = ""
    cboBlock.Text = ""
    txtDate.Text = Format(Date, "mm/dd/yyyy") 'date
    txtUpdatedBy.Text = ""
    txtCatalogID.Text = ""
    txtCategoryID.Text = ""
End Sub

Public Sub CboDisplay()
On Error GoTo er
    If formAction = "add" Then
        'Get Data from database
        DoEvents
        Qry = "Call partscatalog.sp_brand_list()"
        
        With rsMain
            .Open Qry, conPartsCat, adOpenStatic, adLockReadOnly
            If .State = 1 Then
               Call modObjTools.addToListObjectWithArrayGivenField(cboBRAND, rsMain, arrBRAND, "brand_ID", "BRANDNAME")
            End If
            If .State <> 0 Then rsMain.Close
        End With
    End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub

Public Sub CboDisplayModel()
On Error GoTo er
    If formAction = "add" Then
        'Get Data from database
        DoEvents
        Qry = "Call partscatalog.sp_mc_getModel('" & cboBRAND.Text & "')"
        
        With rsMain
            .Open Qry, conPartsCat, adOpenStatic, adLockReadOnly
            If .State = 1 Then
               Call modObjTools.addToListObjectWithArrayGivenField(txtModel, rsMain, arrMODEL, "cat_ID", "cat_Model")
            End If
            If .State <> 0 Then rsMain.Close
        End With
    End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplayModel()")
End Sub

Public Sub CboDisplayVersion()
On Error GoTo er
    If formAction = "add" Then
        'Get Data from database
        DoEvents
        Qry = "Call partscatalog.sp_mc_getVersion('" & txtModel.Text & "')"
        
        With rsMain
            .Open Qry, conPartsCat, adOpenStatic, adLockReadOnly
            If .State = 1 Then
               Call modObjTools.addToListObjectWithArrayGivenField(txtVersion, rsMain, arrVERSION, "cat_ID", "cat_Version")
            End If
            If .State <> 0 Then rsMain.Close
        End With
    End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplayVersion()")
End Sub

Public Sub CboDisplayGroup()
On Error GoTo er
    If formAction = "add" Then
        'Get Data from database
        DoEvents
        Qry = "Call partscatalog.sp_mc_group_list()"
        
        With rsMain
            .Open Qry, conPartsCat, adOpenStatic, adLockReadOnly
            If .State = 1 Then
               Call modObjTools.addToListObjectWithArrayGivenField(cboGROUP, rsMain, arrGROUP, "ID", "GROUPNAME")
            End If
            If .State <> 0 Then rsMain.Close
        End With
    End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplayGroup()")
End Sub

Public Sub CboDisplayBlock()
On Error GoTo er
    If formAction = "add" Then
        'Get Data from database
        DoEvents
        Qry = "Call partscatalog.sp_block_list('" & cboGROUP.Text & "','" & catID & "')"
        
        With rsMain
            .Open Qry, conPartsCat, adOpenStatic, adLockReadOnly
            If .State = 1 Then
               Call modObjTools.addToListObjectWithArrayGivenField(cboBlock, rsMain, arrBLOCK, "ID", "BLOCK")
            End If
            If .State <> 0 Then rsMain.Close
        End With
    End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplayGroup()")
End Sub

Public Sub flxChnger()

    With flxParts
        '.Clear
        .Cols = 7
        .FormatString = "NO. |<PARTS NUMBER|<DESCRIPTION |^REF NO. |^PRICE |^PIECE"
        .ColWidth(0) = 600
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.2  'PARTS NUMBER
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.5 'DESCRIPTION
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.1 'REF NO.
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.11 'PRICE
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.07 'PIECE
        .ColWidth(6) = 0 '(.Width - .ColWidth(0)) * 0.07 'SPID
        flxPartNo.ColWidth(1) = .ColWidth(1)
    End With
    
    With flxPartNo
        .Top = txtPartNo.Height + txtPartNo.Top
        .Left = txtPartNo.Left
        .Width = txtPartNo.Width
    End With
    
End Sub


Sub LockData(DataEntry As Boolean)
    DataEntry = Not DataEntry
    cboBRAND.Locked = DataEntry
    txtModel.Locked = DataEntry
    txtVersion.Locked = DataEntry
    cboGROUP.Locked = DataEntry
    cboBlock.Locked = DataEntry
    txtUpdatedBy.Locked = DataEntry
    LockFlx = DataEntry
    If cboBRAND.Locked Then
        'Frame 1
        cboBRAND.BackColor = txtHoldGrayBackColor.BackColor
        txtModel.BackColor = txtHoldGrayBackColor.BackColor
        txtVersion.BackColor = txtHoldGrayBackColor.BackColor
        cboGROUP.BackColor = txtHoldGrayBackColor.BackColor
        cboBlock.BackColor = txtHoldGrayBackColor.BackColor
        txtUpdatedBy.BackColor = txtHoldGrayBackColor.BackColor
        txtDate.BackColor = txtHoldGrayBackColor.BackColor
    Else
        cboBRAND.BackColor = txtHoldWhiteBackColor.BackColor
        txtModel.BackColor = txtHoldWhiteBackColor.BackColor
        txtVersion.BackColor = txtHoldWhiteBackColor.BackColor
        cboGROUP.BackColor = txtHoldWhiteBackColor.BackColor
        cboBlock.BackColor = txtHoldWhiteBackColor.BackColor
        txtUpdatedBy.Locked = True
    End If
End Sub


Private Sub cboBlock_Change()
On Error GoTo er
    If StatSrch Then
            StatSrch = False
            Exit Sub
    End If
    blockID = 0
    With cboBlock
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
            
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboBlock-Change()")
End Sub

Private Sub cboBlock_Click()
If cboGROUP.Text <> "" Then
    flxParts.SetFocus
    flxParts.row = ActRow
    flxParts.Col = 1
    blockID = arrBLOCK(cboBlock.ListIndex, 0)
Else
    MsgBox "GROUP IS REQUIRED!", vbExclamation, "Validation"
    Exit Sub
End If
End Sub

Private Sub cboBlock_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboBlock_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboBlock, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        flxParts.SetFocus
        flxParts.row = ActRow
        flxParts.Col = 1
    End If
End Sub

Private Sub cboBRAND_Change()
On Error GoTo er
    If StatSrch Then
            StatSrch = False
            Exit Sub
    End If
    catID = 0
    With cboBRAND
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
           
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboBrand-Change()")
End Sub

Private Sub cboBRAND_Click()
    If rsParts.State = adStateOpen Then rsParts.Close
    Call LoadSpareParts
    Call CboDisplayModel
    Call CboDisplayGroup
    txtModel.SetFocus
End Sub

Private Sub cboBRAND_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboBRAND_KeyPress(KeyAscii As Integer)
Dim cbolistIndex As Integer

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboBRAND, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Call cboBRAND_Click
    End If
End Sub

Private Sub cboGROUP_Change()
On Error GoTo er
    If StatSrch Then
            StatSrch = False
            Exit Sub
    End If
    With cboGROUP
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
           
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboGroup-Change()")
End Sub

Private Sub cboGROUP_Click()
If cboBRAND.Text <> "" And txtModel <> "" And txtVersion <> "" Then
    Call CboDisplayBlock
    cboBlock.SetFocus
Else
    If cboBRAND.Text = "" Then
        MsgBox "BRAND IS REQUIRED!", vbExclamation, "Validation"
        cboBRAND.SetFocus
        Exit Sub
    ElseIf txtModel.Text = "" Then
        MsgBox "MODEL IS REQUIRED!", vbExclamation, "Validation"
        txtModel.SetFocus
        Exit Sub
    ElseIf txtVersion.Text = "" Then
        MsgBox "VERSION IS REQUIRED!", vbExclamation, "Validation"
        txtVersion.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub cboGROUP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboGROUP_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboGROUP, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Call cboGROUP_Click
    End If
End Sub

Private Sub cmdAltA_Click()
 Call mnuAdd_Click
End Sub

Private Sub cmdAltD_Click()
Call mnuDelete_Click
End Sub

Private Sub cmdInsX_Click()
    SPInsXstPartsCat.MySource = "Item Entry"
    SPInsXstPartsCat.Show vbModal
End Sub

Private Sub cmdTrans_Click(Index As Integer)
'Index Legend
' 0 -> add
' 1-> edit
' 2 -> view
' 3 -> exit
' 4 - > save
' 5 - > cancel

    If Index = 0 Then 'add
        formAction = "add"
        If rsMain.State = adStateOpen Then rsMain.Close
        Call Clearform
        Call CboDisplay
        Call LockData(True)
        Call SetButtonVisibility(True)
        txtUpdatedBy.Text = EmployeesName
    ElseIf Index = 1 Then 'edit
        If txtCatalogID.Text = "" And txtCategoryID.Text = "" Then
            MsgBox "No Available Parts Catalog to edit", vbOKOnly + vbExclamation, "Edit"
        Else
            If rsMain.State = adStateOpen Then rsMain.Close
            Timer1.Enabled = True
            formAction = "edit"
            Call LockData(True)
            Call SetButtonVisibility(True)
            txtUpdatedBy.Text = EmployeesName
        End If
        
    ElseIf Index = 2 Then 'view summary
       FrmSummaryParts.mySrc = "Item Entry"
       FrmSummaryParts.Show
       Unload Me
    ElseIf Index = 3 Then 'print preview
        If txtCategoryID <> "" Then
            Qry = "call partscatalog.sp_partscatalog_getDetails('" & txtCategoryID.Text & "')"
            frmPrint.QryPrn = Qry
                
            frmPrint.ParamValue1 = cboBRAND.Text
            frmPrint.ParamValue2 = txtModel.Text
            frmPrint.ParamValue3 = txtVersion.Text
            frmPrint.ParamValue4 = cboGROUP.Text
            frmPrint.ParamValue5 = cboBlock.Text
            frmPrint.ParamValue6 = txtDate.Text
            frmPrint.ParamValue7 = txtUpdatedBy.Text
            frmPrint.MySource = "Parts Catalog"
            frmPrint.Show
        Else
            MsgBox "No Available Parts Catalog to Print", vbInformation + vbOKOnly, "Print"
        End If
    ElseIf Index = 4 Then 'exit
        formAction = "EXIT"
        ProperExit = True
        Unload Me
    ElseIf Index = 5 Then 'save
        If ValidateEntry() = False Then Exit Sub
        myMsg = ValidationPartsCatalog
        If myMsg <> "" Then
            MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
            Exit Sub
        End If
        ans = MsgBox("Do you want to save?", vbQuestion + vbYesNo, "Save")
        If ans = vbNo Then Exit Sub
        If formAction = "add" Then
            Call CreateTempTable
            Qry = "Call partscatalog.sp_partscatalog_create('" & cboBRAND.Text & "','" & txtModel.Text & "','" & txtVersion.Text & "','" & cboGROUP.Text & "','" & cboBlock.Text & "','" & UserID & "', '" & UserID & "','" & Format(txtDate.Text, "YYYY-mm-dd") & "', '" & catID & "', '" & blockID & "');"
            'InputBox " " & Qry & " ", vbInformation + vbOKOnly, "SAMPLE"
            'InputBox Qry, Qry, Qry
            rsMain.Open Qry, conPartsCat, adOpenKeyset, adLockOptimistic
            If rsMain.RecordCount <= 0 Then
                rsMain.Close
                Exit Sub
            End If
            If rsMain.RecordCount > 0 Then
                txtCatalogID.Text = rsMain!CATALOGID
                txtCategoryID.Text = rsMain!categoryID
                catID = rsMain!CATALOGID
            End If
        ElseIf formAction = "edit" Then
            Call CreateTempTable
            Qry = "Call partscatalog.sp_partscatalog_update('" & cboBRAND.Text & "','" & txtModel.Text & "','" & txtVersion.Text & "','" & cboGROUP.Text & "','" & cboBlock.Text & "', '" & UserID & "','" & Val(txtCatalogID.Text) & "', '" & Val(txtCategoryID.Text) & "');"
            rsMain.Open Qry, conPartsCat, adOpenKeyset, adLockOptimistic
        End If
        
        ans = MsgBox("Successfully Saved!" & vbCrLf & vbCrLf & "Process Another Entry?", vbQuestion + vbYesNo, "Save")
        
        If ans = vbNo Then
            Call SetButtonVisibility(False)
            Call LockData(False)
            formAction = ""
            Timer1.Enabled = True
        Else
            Call SetButtonVisibility(True)
            Call Clearform
            formAction = "add"
            ActRow = 1
            If rsMain.State = adStateOpen Then rsMain.Close
            txtUpdatedBy.Text = EmployeesName
        End If
        
    ElseIf Index = 6 Then 'cancel
        ans = MsgBox("Do you want to cancel?", vbQuestion + vbYesNo)
        If ans = vbNo Then Exit Sub
        Call SetButtonVisibility(False)
        Call LockData(False)
        Call textVisible(False)
        formAction = ""
        If rsMain.State = adStateOpen Then rsMain.Close
    End If

End Sub

Private Sub flxPartNo_Click()
    Dim PctDisc As Single, myVat As Single, i As Integer
    Dim DiscPrice As Currency, TtlPrice As Currency, MyCost As Currency, CostVat As Currency
    
    With flxPartNo
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.row, 1)
        'Insert Additional retrieval here
        flxParts.TextMatrix(ActRow, 2) = .TextMatrix(.row, 2) 'Description
        flxParts.TextMatrix(ActRow, 4) = Format(.TextMatrix(.row, 3), "##,#00.00") 'Price
        flxParts.TextMatrix(ActRow, 6) = .TextMatrix(.row, 4) 'SPID
    End With
    flxParts.SetFocus
    flxParts.row = ActRow
    flxParts.Col = 3
    txtPartNo.Visible = False
End Sub

Private Sub flxPartNo_EnterCell()
On Error GoTo er
    With flxPartNo
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.row, 1) 'Partnumber
        flxParts.TextMatrix(ActRow, 2) = .TextMatrix(.row, 2) 'Description
        flxParts.TextMatrix(ActRow, 4) = Format(.TextMatrix(.row, 3), "##,#00.00") 'Price
        flxParts.TextMatrix(ActRow, 6) = .TextMatrix(.row, 4) 'SPID
        txtPartNo.SelStart = Val(txtPartNo.Tag)
        txtPartNo.SelLength = Len(txtPartNo.Text)
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->FlxPartNumber-EnterCell()")
End Sub

Private Sub flxPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Tmp() As String
    Dim i As Integer
    If KeyCode = vbKeyEscape Then
       Tmp = Split(txtPartNo.Tag, "|")
       For i = LBound(Tmp) To UBound(Tmp)
           flxPartNo.TextMatrix(ActRow, i + 1) = Tmp(i)
       Next i
       txtPartNo.Visible = False
       flxPartNo.Visible = False
       flxParts.SetFocus
       flxParts.Col = 1
    End If
End Sub

Private Sub flxPartNo_KeyPress(KeyAscii As Integer)
On Error GoTo er
    If KeyAscii = 13 Then
        'Call CheckDuplicate
        If flxParts.TextMatrix(ActRow, 6) = "" Then   'No SPID
            flxParts.SetFocus
            flxParts.row = ActRow
            flxParts.Col = 1
        Else
            flxParts.SetFocus
            flxParts.row = ActRow
            flxParts.Col = 3
        End If
        Call txtPartNo_LostFocus
  End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->FlxPartNumber-Keypress()")
End Sub


Private Sub flxPartNo_LostFocus()
flxPartNo.Visible = False
End Sub

Private Sub flxParts_Click()
On Error GoTo er

    ActRow = flxParts.row
    ActCol = flxParts.Col
    If formAction = "add" Or formAction = "edit" Then
        If ActCol = 0 Or ActRow = 0 Then Exit Sub
            Select Case ActCol
                Case 1
                    With txtPartNo
                        If flxParts.TextMatrix(ActRow, 1) <> "" Then .Tag = flxParts.TextMatrix(ActRow, 1)
                        .Top = Frame2.Top + flxParts.CellTop + flxParts.Top
                        .Left = Frame2.Left + flxParts.CellLeft + flxParts.Left
                        .Height = flxParts.CellHeight
                        .Width = flxParts.CellWidth
                        .Text = flxParts.Text
                        .Visible = True
                        .SelStart = 0
                        .SelLength = Len(.Text)
                        .ZOrder
                        .SetFocus
                    End With
                Case 3
                    With txtRefNo
                        If flxParts.TextMatrix(ActRow, 3) <> "" Then .Tag = flxParts.TextMatrix(ActRow, 3)
                        .Top = Frame2.Top + flxParts.CellTop + flxParts.Top
                        .Left = Frame2.Left + flxParts.CellLeft + flxParts.Left
                        .Height = flxParts.CellHeight
                        .Width = flxParts.CellWidth
                        .Text = flxParts.Text
                        .Visible = True
                        .SelStart = 0
                        .SelLength = Len(.Text)
                        .ZOrder
                        .SetFocus
                    End With
                Case 5
                    With txtPiece
                        If flxParts.TextMatrix(ActRow, 5) <> "" Then .Tag = flxParts.TextMatrix(ActRow, 4)
                        .Top = Frame2.Top + flxParts.CellTop + flxParts.Top
                        .Left = Frame2.Left + flxParts.CellLeft + flxParts.Left
                        .Height = flxParts.CellHeight
                        .Width = flxParts.CellWidth
                        .Text = flxParts.Text
                        .Visible = True
                        .SelStart = 0
                        .SelLength = Len(.Text)
                        .ZOrder
                        .SetFocus
                    End With
            End Select
    End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->flxSPARTS-click()")
End Sub

Private Sub flxParts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call flxParts_Click
    End If
End Sub

Private Sub flxParts_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If formAction = "add" Then
            PopupMenu mnuStatus, vbPopupMenuRightButton
        ElseIf formAction = "edit" Then
            PopupMenu mnuStatus, vbPopupMenuRightButton
        ElseIf formAction = "summary" Then
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Load()
    Call ConnectPartsCatalog
    Call LockData(False)
    txtDate.Text = Format$(Now, "mm/dd/yyyy")
    ActRow = 1
    
    '===================BUTTON POSITION===================
    If cmdTrans(0).Visible = False Then
        'Align Save Button
        cmdTrans(5).Left = cmdTrans(0).Left
        cmdTrans(5).Top = cmdTrans(0).Top
        'Align Cancel Button
        cmdTrans(6).Left = cmdTrans(1).Left
        cmdTrans(6).Top = cmdTrans(1).Top
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ProperExit Then
        If rsMain.State = adStateOpen Then rsMain.Close
        If rsParts.State = adStateOpen Then rsParts.Close
        Call cmdTrans_Click(4)
        ProperExit = False
    Else
        Cancel = True
    End If
End Sub

Private Sub mnuAdd_Click()
On Error GoTo er
    ActRow = flxParts.Rows - 1
    If flxParts.TextMatrix(ActRow, 1) <> flxParts.TextMatrix(flxParts.Rows - 1, 1) Then
        flxParts.SetFocus
        flxParts.Col = 1
        Exit Sub
    End If

    If flxParts.TextMatrix(flxParts.Rows - 1, 6) = "" Or flxParts.TextMatrix(flxParts.Rows - 1, 3) = "" Or flxParts.TextMatrix(flxParts.Rows - 1, 5) = "" Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        'Call textVisible(False)
        If flxParts.TextMatrix(flxParts.Rows - 1, 1) = "" Then
            flxParts.SetFocus
            flxParts.Col = 1
        ElseIf flxParts.TextMatrix(flxParts.Rows - 1, 3) = "" Then
            flxParts.SetFocus
            flxParts.Col = 3
        ElseIf flxParts.TextMatrix(flxParts.Rows - 1, 5) = "" Then
            flxParts.SetFocus
            flxParts.Col = 5
        End If
        Exit Sub
    End If
    flxParts.AddItem ""
    flxParts.SetFocus
    flxParts.row = flxParts.Rows - 1
    flxParts.Col = 1
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->mnuAdd()")
    
End Sub

Private Sub mnuDelete_Click()
On Error GoTo er

    'actrow = flxPARTS.Rows - 1
    
    ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If ans = vbYes Then
        If flxParts.Rows = 2 Then
            flxParts.Rows = 1
            flxParts.Rows = 2
        Else
            flxParts.RemoveItem ActRow
        End If
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->mnuDelete()")
End Sub


Private Sub RXResizer1_ResizeComplete()
    Call flxChnger
End Sub

Private Sub Timer1_Timer()
    Call LoadSpareParts
    Timer1.Enabled = False
End Sub

Public Sub LoadSpareParts()
On Error GoTo er
    DoEvents
    
    If rsParts.State = adStateOpen Then rsParts.Close
    Qry = "call partscatalog.sp_spareparts_list('" & cboBRAND.Text & "')"
    rsParts.Open Qry, conPartsCat, adOpenKeyset, adLockOptimistic
    Set flxPartNo.DataSource = rsParts
    flxPartNo.Refresh
    With flxPartNo
        .Width = flxParts.ColWidth(1)
        .ColWidth(0) = 0
        .ColWidth(1) = flxParts.ColWidth(1)
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .ColWidth(6) = 0
    End With
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->LoadSpareparts()")
End Sub

Private Sub txtModel_Change()
On Error GoTo er
    If StatSrch Then
            StatSrch = False
            Exit Sub
    End If
    With txtModel
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
           
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtModel-Change()")
End Sub

Private Sub txtModel_Click()
If cboBRAND.Text <> "" Then
    Call CboDisplayVersion
    txtVersion.SetFocus
Else
   MsgBox "BRAND IS REQUIRED!", vbExclamation, "Validation"
   cboBRAND.SetFocus
   Exit Sub
End If
End Sub

Private Sub txtModel_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(txtBrand, KeyAscii, ValidChars)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = 13 Then Call txtModel_Click
End Sub

Private Sub txtPartNo_Change()

On Error GoTo LocalErr
    With flxPartNo
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
            .Col = 0: .ColSel = 1
            txtPartNo.Text = rsParts!PartNumber
            flxParts.TextMatrix(flxParts.row, 2) = rsParts!Description
            flxParts.TextMatrix(flxParts.row, 4) = Format(rsParts!STOREPRICE, "##,#00.00")
            flxParts.TextMatrix(flxParts.row, 6) = rsParts!SPAREPARTID
            txtPartNo.SelStart = Val(txtPartNo.Tag)
            txtPartNo.SelLength = Len(txtPartNo.Text)
            PrtBln = False
                 
                 
        Else
    
            flxParts.TextMatrix(ActRow, 1) = ""
            flxParts.TextMatrix(ActRow, 2) = ""
            flxParts.TextMatrix(ActRow, 3) = ""
            flxParts.TextMatrix(ActRow, 4) = ""
            flxParts.TextMatrix(ActRow, 5) = ""
            flxParts.TextMatrix(ActRow, 6) = ""
            .Visible = False
        End If
   End With
   
   Exit Sub

LocalErr:
    If Err.Number = 30009 Then
        flxPartNo.Visible = False
        txtPartNo.Visible = False
        If rsParts.State = adStateOpen Then rsParts.Close
        Timer1.Enabled = True
        MsgBox "Program is reloading Part Number" & vbCrLf & "Please wait for a few seconds...", vbInformation + vbOKOnly, "Reloading"
    Else
        flxPartNo.Visible = False
        txtPartNo.Visible = False
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
    End If
End Sub

Private Sub txtPartNo_GotFocus()
    If rsParts.State = adStateClosed Then Timer1.Enabled = True
    With flxPartNo
        .Top = txtPartNo.Height + txtPartNo.Top
        .Left = txtPartNo.Left
        .Width = txtPartNo.Width
    End With
    
End Sub

Private Sub txtPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        txtPartNo.Tag = "skiplostfocus"
        flxPartNo.Visible = True
        flxPartNo.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        flxParts.SetFocus
        flxParts.Col = 1
        txtPartNo.Visible = False
        flxPartNo.Visible = False
    ElseIf KeyCode = vbKeyDelete Then
        PrtSrch = True
        rsParts.MoveFirst
        flxParts.TextMatrix(flxParts.row, 2) = ""
        flxParts.TextMatrix(flxParts.row, 4) = ""
        flxParts.TextMatrix(flxParts.row, 6) = ""
    ElseIf KeyCode = vbKeyBack Then
        PrtSrch = True
        rsParts.MoveFirst
        flxParts.TextMatrix(flxParts.row, 2) = ""
        flxParts.TextMatrix(flxParts.row, 4) = ""
        flxParts.TextMatrix(flxParts.row, 6) = ""
        flxPartNo.Visible = False
    End If
End Sub

Private Sub txtPartNo_KeyPress(KeyAscii As Integer)
On Error GoTo er
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(txtPARTNO, KeyAscii, ValidChars)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If rsParts.State = adStateClosed Then Timer1.Enabled = True
    
    If KeyAscii = 13 Then
        'Call CheckDuplicate
        'If chkdupCount <= 0 Then
            If flxParts.TextMatrix(ActRow, 6) = "" Then   'No SPID
                flxParts.SetFocus
                flxParts.row = ActRow
                flxParts.Col = 1
            Else
                flxParts.SetFocus
                flxParts.row = ActRow
                flxParts.Col = 3
            End If
        'Else
            'flxParts.SetFocus
            'flxParts.Row = ActRow
            'flxParts.Col = 1
        'End If
        flxPartNo.Visible = False
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtPARTNO_KeyPress()")
End Sub

Private Sub txtPartNo_LostFocus()
On Error GoTo er
    If txtPartNo.Tag = "skiplostfocus" Then
        txtPartNo.Tag = ""
        Exit Sub
    End If

    If txtPartNo.Text <> "" Then
        If flxParts.TextMatrix(ActRow, 6) = "" Or txtPartNo.Text = "" Then
            MsgBox "Unrecognized Part Number!" & vbCrLf & "Please select a Part Number from list.", vbExclamation + vbOKOnly, "Warning"
            flxParts.TextMatrix(ActRow, 1) = ""
            flxParts.TextMatrix(ActRow, 2) = ""
            flxParts.TextMatrix(ActRow, 3) = ""
            flxParts.TextMatrix(ActRow, 4) = ""
            flxParts.TextMatrix(ActRow, 5) = ""
            flxParts.TextMatrix(ActRow, 6) = ""
            flxParts.SetFocus
            flxParts.Col = 1
            txtPartNo.Visible = False
            flxPartNo.Visible = False
        Exit Sub
        End If
    End If
    flxParts.TextMatrix(ActRow, 1) = txtPartNo.Text
    txtPartNo.Visible = False
    flxPartNo.Visible = False
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->PartNumber-LostFocus()")
End Sub

Private Sub txtPiece_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(txtBrand, KeyAscii, ValidChars)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = 13 Then
        Call txtPiece_LostFocus
        If flxParts.TextMatrix(ActRow, 5) <> "" Then
            Call mnuAdd_Click
        End If
        txtPiece.Visible = False
    End If
End Sub

Private Sub txtPiece_LostFocus()
    If txtPiece.Tag = "skiplostfocus" Then
        txtPiece.Tag = ""
        Exit Sub
    End If
    flxParts.TextMatrix(ActRow, 5) = txtPiece.Text
    txtPiece.Visible = False
End Sub

Private Sub txtrefno_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(txtBrand, KeyAscii, ValidChars)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = 13 Then
        If flxParts.TextMatrix(ActRow, 1) <> "" Then
            Call txtrefno_LostFocus
        End If
    End If
End Sub

Private Sub txtrefno_LostFocus()
    If txtRefNo.Tag = "skiplostfocus" Then
        txtRefNo.Tag = ""
        Exit Sub
    End If
    flxParts.SetFocus
    flxParts.row = ActRow
    flxParts.Col = 5
    flxParts.TextMatrix(ActRow, 3) = txtRefNo.Text
    txtRefNo.Visible = False
End Sub

Private Sub txtVersion_Change()
On Error GoTo er
    If StatSrch Then
            StatSrch = False
            Exit Sub
    End If
    catID = 0
    With txtVersion
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtVersion-Change()")
End Sub

Private Sub txtVersion_Click()
If txtModel.Text <> "" Then
    cboGROUP.SetFocus
    catID = arrVERSION(txtVersion.ListIndex, 0)
Else
    MsgBox "MODEL IS REQUIRED!", vbExclamation, "Validation"
    txtModel.SetFocus
    Exit Sub
End If
End Sub

Private Sub txtVersion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(txtBrand, KeyAscii, ValidChars)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = 13 Then Call txtVersion_Click
End Sub
