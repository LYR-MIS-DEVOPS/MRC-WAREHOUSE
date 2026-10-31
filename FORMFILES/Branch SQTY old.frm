VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BranchSQTY 
   Caption         =   "Branch Quantity"
   ClientHeight    =   10050
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Branch SQTY.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10050
   ScaleWidth      =   19200
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flxPG 
      Height          =   1695
      Left            =   19320
      TabIndex        =   37
      Top             =   6720
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2990
      _Version        =   393216
   End
   Begin VB.TextBox txtPMID 
      Height          =   375
      Left            =   19320
      TabIndex        =   28
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   19320
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   19320
      TabIndex        =   26
      Text            =   "HoldBackColor"
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   21
      Tag             =   "XY"
      Top             =   9240
      Width           =   18975
      Begin VB.TextBox txtUpdated 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Tag             =   "XY"
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label lblUP 
         Caption         =   "UPDATED BY:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Tag             =   "XY"
         Top             =   245
         Width           =   1095
      End
   End
   Begin VB.TextBox txtSTDQ 
      Height          =   285
      Left            =   19320
      TabIndex        =   20
      Tag             =   "YX"
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   19320
      Top             =   720
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   9
      Tag             =   "XY"
      Top             =   480
      Width           =   18975
      Begin VB.ComboBox cmbBranch 
         Height          =   315
         Left            =   1920
         TabIndex        =   35
         Tag             =   "XY"
         Top             =   240
         Width           =   5175
      End
      Begin VB.TextBox txtEmpty 
         Height          =   285
         Left            =   17400
         TabIndex        =   6
         Tag             =   "XY"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtLacking 
         Height          =   285
         Left            =   17400
         TabIndex        =   7
         Tag             =   "XY"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtExcess 
         Height          =   285
         Left            =   17400
         TabIndex        =   8
         Tag             =   "XY"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtPG 
         Height          =   285
         Left            =   15000
         TabIndex        =   33
         Tag             =   "XY"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtTTSTDQ 
         Height          =   285
         Left            =   15000
         TabIndex        =   32
         Tag             =   "XY"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtTTPN 
         Height          =   285
         Left            =   15000
         TabIndex        =   31
         Tag             =   "XY"
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cmbModel 
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Tag             =   "XY"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.ComboBox cmbBrand 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Tag             =   "XY"
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "New Branch"
         Height          =   255
         Left            =   5280
         TabIndex        =   39
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblBranch 
         Caption         =   "SELECT BRANCH:"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Tag             =   "XY"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblEmpty 
         Alignment       =   1  'Right Justify
         Caption         =   "EMPTY:"
         Height          =   255
         Left            =   16560
         TabIndex        =   19
         Tag             =   "XY"
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblLacking 
         Alignment       =   1  'Right Justify
         Caption         =   "LACKING:"
         Height          =   255
         Left            =   16560
         TabIndex        =   18
         Tag             =   "XY"
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblExcess 
         Alignment       =   1  'Right Justify
         Caption         =   "EXCESS:"
         Height          =   255
         Left            =   16560
         TabIndex        =   17
         Tag             =   "XY"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblTotalPG 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL PARTS GROUP:"
         Height          =   255
         Left            =   12960
         TabIndex        =   16
         Tag             =   "XY"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblTotalSTDQ 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL STDQ:"
         Height          =   255
         Left            =   13560
         TabIndex        =   15
         Tag             =   "XY"
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblTotalPN 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL PN:"
         Height          =   255
         Left            =   14040
         TabIndex        =   14
         Tag             =   "XY"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblModel 
         Caption         =   "SELECT MC MODELS:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Tag             =   "XY"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblBrand 
         Caption         =   "SELECT MC BRAND:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Tag             =   "XY"
         Top             =   720
         Width           =   1695
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   19320
      TabIndex        =   24
      Top             =   2280
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "Save"
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
      Image           =   "Branch SQTY.frx":57E2
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   19320
      TabIndex        =   25
      Top             =   2760
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
      Image           =   "Branch SQTY.frx":5D7C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Tag             =   "YX"
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "&View Summary"
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
      Image           =   "Branch SQTY.frx":6316
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   2040
      TabIndex        =   1
      Tag             =   "YX"
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
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
      Image           =   "Branch SQTY.frx":68B0
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame3 
      Height          =   7095
      Left            =   120
      TabIndex        =   23
      Tag             =   "XY"
      Top             =   2160
      Width           =   18975
      Begin MSComctlLib.ProgressBar pBar 
         Height          =   255
         Left            =   0
         TabIndex        =   38
         Tag             =   "XY"
         Top             =   6840
         Visible         =   0   'False
         Width           =   18975
         _ExtentX        =   33470
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxBSQTY 
         Height          =   6375
         Left            =   120
         TabIndex        =   10
         Tag             =   "XY"
         Top             =   600
         Width           =   18735
         _ExtentX        =   33046
         _ExtentY        =   11245
         _Version        =   393216
         Cols            =   13
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   14737632
         ScrollBars      =   2
         AllowUserResizing=   1
         FormatString    =   "|^PARTNUMBER |<DESCRIPTION |^PARTSGROUP |^MC MODEL  |^QOH |^AVQ 3MSQ |^STDQ |<LACKING |<EXCESS |^SPECIAL ORDER"
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
      Begin lvButton.lvButtons_H btnFind 
         Height          =   375
         Left            =   15600
         TabIndex        =   11
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
         Image           =   "Branch SQTY.frx":6E4A
         cBack           =   -2147483633
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   16920
         TabIndex        =   4
         Tag             =   "XY"
         Top             =   195
         Width           =   1935
      End
      Begin VB.CommandButton btnAltA 
         Caption         =   "&A"
         Height          =   315
         Left            =   9360
         TabIndex        =   29
         Top             =   2280
         Width           =   495
      End
      Begin VB.CommandButton btnAltD 
         Caption         =   "&D"
         Height          =   375
         Left            =   9120
         TabIndex        =   30
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblPartsGroup 
         Caption         =   "Parts Group"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Tag             =   "XY"
         Top             =   240
         Width           =   11055
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "&New Identifying"
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
      Image           =   "Branch SQTY.frx":73E4
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "&Edit Identifying"
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
      Image           =   "Branch SQTY.frx":797E
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   19320
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   19320
      m_MDI_Height    =   10560
   End
   Begin VB.Menu mnuStatus 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "AltA"
      End
   End
End
Attribute VB_Name = "BranchSQTY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Strings variable
Dim formAction As String
Dim strPartial As String, strTotal As String
Dim MyErr As String, myMsg As String
Dim Qry As String
Dim strTtlExcess As String
Dim strTtlLacking As String
Dim strTtlEmpty As String
Dim strTtlQOH As String
Dim Tmp() As String
Dim Ans As String
Dim SixMSQ As String, Lacking As String, Excess As String, STDQ As String
Dim pPMSO As String
Dim strBoolean As String
Dim branchCode As String
Dim strPG As String

Dim arrMODEL() As String
Dim arrBRAND() As String
Dim arrBRANCH() As String

Public mySrc As String

'recordset variable
Dim mainRs As New ADODB.Recordset
Dim tmpRs  As New ADODB.Recordset

'integer variable
Dim i As Long, j As Long, a As Integer
Dim ActRow As Integer, ActCol As Integer, fRow As Integer, fCol As Integer
Dim LoopSTDQ As Integer
Dim intTotalSTDQ As Integer
Dim LoopLACKING As Integer
Dim LoopExcess As Integer
Dim LoopEmpty As Integer
Dim LoopQOH As Integer
Dim flxLack As Double
Dim matrixID As Integer
Dim pSPAREPARTID As Integer
Dim intPG As Integer
Dim NewPrtsGroup As String
Dim OldPrtsGroup As String

'******************* CODE UPDATED: 04/08/17 *************************
Dim intComputeTtlLacking As Integer, intComputeTtlExcess As Integer
Dim intComputeTtlEmpty As Integer, intComputeTtlQtyOnHand As Integer
Dim intComputeTtlSTDQ As Integer, intComputeTtlPN As Integer, intComputePG As Integer
'*************** END CODE UPDATED: 04/08/17 *************************

Dim intQOH As Double
Dim intExcess As Double
Dim intLacking As Double
Dim intEmpty As Double
Dim intSTDQ As Double
Dim intFlxEmpty As Double
Dim intTotlQOH As Double
Dim pPMSTDQ As Double
Dim pPMLACKING As Double
Dim pPMEXCESS As Double

'Boolean variable
Dim StatSrch As Boolean, AddNew As Boolean, ProperExit As Boolean
Dim PrtBln As Boolean, PrtSrch As Boolean, cboBln As Boolean

'===================BUTTON VISIBILITY===================
Private Sub SetButtonVisibility(DataEntry As Boolean)
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->VIEW SUMM
'3 ->EXIT
'4 ->SAVE
'5 ->CANCEL

    'hide all button
    For i = 0 To 3
        cmdTrans(i).Visible = False
    Next i

    If DataEntry Then
        cmdTrans(4).Visible = True
        cmdTrans(5).Visible = True
    Else
    For i = 0 To 3
        cmdTrans(i).Visible = True
        cmdTrans(4).Visible = False
        cmdTrans(5).Visible = False
    Next i

    End If
End Sub


Public Sub flxChanger()
    With flxBSQTY
        .Cols = 14
        .FormatString = " |^PARTNUMBER |<DESCRIPTION |^PARTSGROUP |^MC MODEL  |^QOH |^AVQ 3MSQ |^NBS |^STDQ |^LACKING |^EXCESS |^ORDER FREQ "
        
        .ColWidth(0) = 400
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.12   'PARTS NUMBER
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.2  'DESCRIPTION
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.1 'PARTSGROUP
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.1 'MC MODEL
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.06 'QUANTITY ON HAND
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.06 '6MONTHS SALE QUANTITY
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.06 'NEW BRANCH STOCKING
        .ColWidth(8) = (.Width - .ColWidth(0)) * 0.06 'STANDARD QUANTITY
        .ColWidth(9) = (.Width - .ColWidth(0)) * 0.06 'LACKING
        .ColWidth(10) = (.Width - .ColWidth(0)) * 0.06 'EXCESS
        .ColWidth(11) = (.Width - .ColWidth(0)) * 0.1 'ORDER FREQ
        .ColWidth(12) = 0 '(.Width - .ColWidth(0)) * 0.03  ' APPMODEL
        .ColWidth(13) = 0 '(.Width - .ColWidth(0)) * 0.03  ' ID
    End With
End Sub


'Private Sub btnAltD_Click()
'    Call mnuDelete_Click
'End Sub
'
'Private Sub btnAltA_Click()
'    Call mnuAdd_Click
'End Sub


Private Sub btnFind_Click()

    If txtFind.Text = "" Then Exit Sub
    With flxBSQTY
        For i = fRow To .Rows - 1
            For j = fCol To 10
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 0: .Row = i
                    .ColSel = 10:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 10, i + 1, i)
                    fCol = IIf(j = 10, 1, j + 1)
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





Private Sub cmbBranch_Change()
On Error GoTo er
    If StatSrch Then
            StatSrch = False
            Exit Sub
    End If

    With cmbBranch
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmbBranch-Change()")
End Sub

Private Sub cmbBranch_Click()
    
    cmbBrand.Clear
    cmbModel.Clear
    flxBSQTY.Rows = 1
    flxBSQTY.Rows = 2
    lblPartsGroup.Caption = "Parts Group"
    
    branchCode = arrBRANCH(cmbBranch.ListIndex, 0)
    
    Call CboBrandDisplay
End Sub

Private Sub cmbBranch_GotFocus()
Dim Tmp
    Const CB_SHOWDROPDOWN = &H14F
    Tmp = SendMessage(cmbBranch.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cmbBranch_KeyDown(KeyCode As Integer, Shift As Integer)
'************** CODE UPDATED: 04/07/17 **************************
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
'********** END CODE UPDATED: 04/07/17 **************************
End Sub

Private Sub cmbBranch_KeyPress(KeyAscii As Integer)
'************** CODE UPDATED: 04/07/17 **************************
Dim cbolistIndex As Integer

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cmbBranch, KeyAscii, ValidChars)
    
    If KeyAscii = 13 Then
        cbolistIndex = modObjTools.getListIndex(cmbBranch, cmbBranch.Text)
        
        If cbolistIndex < 0 Then
            cmbBranch.SetFocus
        Else
            cmbBranch.ListIndex = cbolistIndex
        End If
        
        If formAction = "add" Then
            cmbBrand.Clear
        End If
    
        Call CboBrandDisplay
        cmbBrand.SetFocus
        
    End If
'********** END CODE UPDATED: 04/07/17 **************************
End Sub

Private Sub cmbBrand_Change()
On Error GoTo er
    If StatSrch Then
            StatSrch = False
            Exit Sub
    End If

    With cmbBrand
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmbBrand-Change()")
End Sub



Private Sub cmbBrand_Click()

    Call cmbBrand_Change
    
    If formAction = "add" Or formAction = "edit" Then
        cmbModel.Clear
    End If

    Call CboModelDisplay
    

End Sub


Private Sub cmbBrand_GotFocus()
Dim Tmp
    Const CB_SHOWDROPDOWN = &H14F
    Tmp = SendMessage(cmbBrand.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cmbBrand_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cmbBrand_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cmbBrand, KeyAscii, ValidChars)
    
    If KeyAscii = 13 Then
    
        If formAction = "add" Then
            cmbModel.Clear
        End If
    
        Call CboModelDisplay
        cmbModel.SetFocus
        
    End If

End Sub



Private Sub cmbModel_Change()
On Error GoTo er

    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If
    
    With cmbModel
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            
            j = Len(strTotal) - Len(strPartial)
            
            matrixID = arrMODEL(i, 0)
           
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmbModel()")
End Sub

Private Sub cmbModel_Click()
    
    DoEvents
    matrixID = arrMODEL(cmbModel.ListIndex, 0)
    txtPMID.Text = matrixID
'********** CODE UPDATED: 04/11/17 ****************
    pBar.Visible = True
'****** END CODE UPDATED: 04/11/17 ****************
    Call ShowRecord(matrixID, branchCode)

    flxBSQTY.SetFocus
    flxBSQTY.Col = 7

End Sub



Private Sub cmbModel_GotFocus()
Dim Tmp
    Const CB_SHOWDROPDOWN = &H14F
    Tmp = SendMessage(cmbModel.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cmbModel_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cmbModel_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cmbModel, KeyAscii, ValidChars)
    
    If KeyAscii = 13 Then

        flxBSQTY.SetFocus
        flxBSQTY.Col = 7
        
    End If
End Sub




Private Sub cmdTrans_Click(Index As Integer)
On Error GoTo er

    If Index = 0 Then 'view summary
        frmSummaryRepl.mySrc = "Branch Quantity"
        frmSummaryRepl.Show
        Unload Me
    ElseIf Index = 1 Then 'edit
        If txtPMID = "" Then
            MsgBox "No available PARTS MATRIX to edit", vbOKOnly + vbExclamation, "Edit"
        Else
            formAction = "edit"
            AddNew = False
            Call SetButtonVisibility(True)
            Call LockData(False)
            flxBSQTY.SetFocus
            flxBSQTY.Col = 8
        End If
    ElseIf Index = 2 Then 'exit
        formAction = "EXIT"
        ProperExit = True
        Call clearform
        Unload Me
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdTrans-Click()")
End Sub


Private Sub flxBSQTY_DblClick()
    If formAction = "add" Or formAction = "edit" Then
        If flxBSQTY.TextMatrix(flxBSQTY.Row, 10) = "SO" Then
            flxBSQTY.TextMatrix(ActRow, 10) = ""
        End If
    End If
End Sub


Private Sub flxBSQTY_Scroll()
txtSTDQ.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo er

    If ProperExit Then
        If mainRs.State = adStateOpen Then mainRs.Close
        Call cmdTrans_Click(3)
        ProperExit = False
    Else
        Cancel = True
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->QueryUnload()")
End Sub

Private Sub Form_Resize()
    Call RXResizer1_ResizeComplete
End Sub







Private Sub RXResizer1_ResizeComplete()
    
    Call flxChanger
'************* CODE UPDATED: 3/14/17 *****************
    '    'Frame width
    '    Frame1.Width = Me.Width - 465
    '    Frame2.Width = Me.Width - 465
    '    Frame3.Width = Me.Width - 465
    '
    '    'Frame1 label Left
    '    lblBrand.Left = 120
    '    lblModel.Left = 120
    '    lblBranch.Left = 120
    '    lblTotalPN.Left = Frame1.Width - 5055
    '    lblTotalSTDQ.Left = Frame1.Width - 5535
    '    lblTotalPG.Left = Frame1.Width - 6135
    '    lblExcess.Left = Frame1.Width - 2415
    '    lblLacking.Left = Frame1.Width - 2415
    '    lblEmpty.Left = Frame1.Width - 2415
    '
    '    'Frame1 Label Width
    '    lblTotalPN.Width = 855
    '    lblTotalSTDQ.Width = 1335
    '    lblTotalPG.Width = 1935
    '    lblBrand.Width = 1695
    '    lblModel.Width = 1695
    '    lblBranch.Width = 1695
    '    lblExcess.Width = 735
    '    lblLacking.Width = 735
    '    lblEmpty.Width = 735
    '
    '    'Frame1 Combobox Left
    '    cmbBrand.Left = 1920
    '    cmbModel.Left = 1920
    '    cmbBranch.Left = 1920
    '
    '    'Frame1 Combobox Width
    '    cmbBrand.Width = 1935
    '    cmbModel.Width = 1935
    '    cmbBranch.Width = 5175
    '
    '    'Frame1 Text Left
    '    txtTTPN.Left = Frame1.Width - 4095
    '    txtTTSTDQ.Left = Frame1.Width - 4095
    '    txtPG.Left = Frame1.Width - 4095
    '    txtExcess.Left = Frame1.Width - 1575
    '    txtLacking.Left = Frame1.Width - 1575
    '    txtEmpty.Left = Frame1.Width - 1575
    '
    '    'Frame1 Text Width
    '    txtTTPN.Width = 1455
    '    txtTTSTDQ.Width = 1455
    '    txtPG.Width = 1455
    '    txtExcess.Width = 1455
    '    txtLacking.Width = 1455
    '    txtEmpty.Width = 1455
    '
    '
    '    'Frame2 Label Left
    '    lblUP.Left = 240
    '
    '    'Frame2 Label Width
    '    lblUP.Width = 1095
    '
    '    'Frame2 Text Left
    '    txtUpdated.Left = 1440
    '
    '    'Frame2 Text Width
    '    txtUpdated.Width = 4095
    '
    '    'Frame3 Text Left
    '    txtFind.Left = Frame3.Width - 2055
    '
    '    'Frame3 Text Width
    '    txtFind.Width = 1935
    '
    '    'Frame3 Button Left
    '    btnFind.Left = Frame3.Width - 3375
    '
    '    'Frame3 Button Width
    '    btnFind.Width = 1215
    '
    '    'Frame3 FlexGrid Width
    '    flxBSQTY.Width = Frame3.Width - 240
    '
    '    'Frame3 FlexGrid Height
    '    flxBSQTY.Height = Frame3.Height - 720
'********* END CODE UPDATED: 3/14/17 *****************
End Sub



Private Sub Form_Load()

  '  Call modCon.Connection 'database connection
    Call flxChanger
    
    Call LockData(False)
    Call CboDisplay
    txtExcess.Text = "0.0%"
    txtLacking.Text = "0.0%"
    txtEmpty.Text = "0.0%"
    txtUpdated.Text = 2
    '===================BUTTON POSITION===================
    
    If cmdTrans(0).Visible = False Then
            'Align Save Button
            cmdTrans(4).Left = cmdTrans(0).Left
            cmdTrans(4).Top = cmdTrans(0).Top
            'Align Cancel Button
            cmdTrans(5).Left = cmdTrans(1).Left
            cmdTrans(5).Top = cmdTrans(1).Top
    
    End If
    
End Sub

Sub LockData(DataEntry As Boolean)
    DataEntry = Not DataEntry

    txtTTPN.Locked = DataEntry
    txtTTSTDQ.Locked = DataEntry
    txtPG.Locked = DataEntry
    txtLacking.Locked = DataEntry
    txtExcess.Locked = DataEntry
    txtEmpty.Locked = DataEntry
    txtUpdated.Locked = DataEntry
   
    If txtTTPN.Locked Then
        txtTTPN.BackColor = txtHoldGrayBackColor.BackColor
        txtTTSTDQ.BackColor = txtHoldGrayBackColor.BackColor
        txtPG.BackColor = txtHoldGrayBackColor.BackColor
        txtLacking.BackColor = txtHoldGrayBackColor.BackColor
        txtExcess.BackColor = txtHoldGrayBackColor.BackColor
        txtUpdated.BackColor = txtHoldGrayBackColor.BackColor
        txtEmpty.BackColor = txtHoldGrayBackColor.BackColor

    Else
        txtTTPN.BackColor = txtHoldWhiteBackColor.BackColor
        txtTTSTDQ.BackColor = txtHoldWhiteBackColor.BackColor
        txtPG.BackColor = txtHoldWhiteBackColor.BackColor
        txtLacking.Locked = txtHoldWhiteBackColor.BackColor
        txtExcess.Locked = txtHoldWhiteBackColor.BackColor
        txtUpdated.Locked = txtHoldWhiteBackColor.BackColor
        txtEmpty.Locked = txtHoldWhiteBackColor.BackColor
        
    End If
End Sub

Public Sub CboDisplay()
On Error GoTo er
    'Get Data from database
    DoEvents
    Call modPMBranchSQTY.sp_get_Branch(tmpRs)
    If tmpRs.State = 1 Then
        Call modObjTools.addToListObjectWithArrayGivenField(cmbBranch, tmpRs, arrBRANCH, "BRANCHCODE", "BRANCHNAME")
    End If
    If tmpRs.State <> 0 Then tmpRs.Close
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub

Public Sub CboBrandDisplay()
On Error GoTo er
        'Get Data from database
        DoEvents
        Call modPMBranchSQTY.get_Brand(branchCode, tmpRs)
        If tmpRs.State = 1 Then
            Call modObjTools.addToListObjectWithArrayGivenField(cmbBrand, tmpRs, arrBRAND, "nulls", "PMBRAND")
        End If
        If tmpRs.State <> 0 Then tmpRs.Close
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub

Public Sub CboModelDisplay()
On Error GoTo er
    DoEvents
    Call modPMBranchSQTY.get_Model(cmbBrand.Text, branchCode, tmpRs)
    If tmpRs.State = 1 Then
        Call modObjTools.addToListObjectWithArrayGivenField(cmbModel, tmpRs, arrMODEL, "PMATRIXID", "PMMODEL")
    End If
    If tmpRs.State <> 0 Then tmpRs.Close
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboModelDisplay()")
End Sub


Sub clearform()

    flxBSQTY.Rows = 1
    flxBSQTY.Rows = 2

    cmbBrand.Text = ""
    cmbModel.Text = ""
    txtTTPN.Text = ""
    txtTTSTDQ.Text = ""
    txtPG.Text = ""
    txtExcess.Text = "0.0%"
    txtLacking.Text = "0.0%"
    txtEmpty.Text = "0.0%"
    txtPMID.Text = ""
    txtUpdated.Text = 2
    txtTTSTDQ.Text = ""
    

    
    
End Sub




Private Sub flxBSQTY_Click()
On Error GoTo er

    ActRow = flxBSQTY.Row
    ActCol = flxBSQTY.Col
    
    If formAction = "add" Or formAction = "edit" Then 'clickable all
    
    If ActCol = 0 Or ActRow = 0 Then Exit Sub
        Select Case ActCol
        
        Case 7
            If flxBSQTY.TextMatrix(flxBSQTY.Row, 1) <> "" Then
                With txtSTDQ
                    If flxBSQTY.TextMatrix(ActRow, 7) <> "" Then .Tag = flxBSQTY.TextMatrix(ActRow, 7)
                    .Top = Frame3.Top + flxBSQTY.CellTop + flxBSQTY.Top
                    .Left = Frame3.Left + flxBSQTY.CellLeft + flxBSQTY.Left
                    .Height = flxBSQTY.CellHeight
                    .Width = flxBSQTY.CellWidth
                    .Text = flxBSQTY.Text
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .ZOrder
                    .SetFocus
                End With
            Else
                Exit Sub
            End If
            
        Case 10
            If flxBSQTY.TextMatrix(flxBSQTY.Row, 1) <> "" Then
                flxBSQTY.TextMatrix(ActRow, 10) = "SO"
            End If
            
        End Select
    End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->flxBSQTY-click()")
End Sub

Private Sub flxBSQTY_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxBSQTY_Click
End Sub

Public Sub resetVar()
    strPG = ""
    intPG = 0
    flxPG.Rows = 1
    flxPG.Rows = 2
    NewPrtsGroup = ""
    OldPrtsGroup = ""
End Sub
Public Sub ShowRecord(MyID As Integer, MyBranchCode As String)
Dim intComputeLacking As Integer, intComputeExcess As Integer
Dim intSTDQty As Integer
On Error GoTo er
   
    txtPMID.Text = MyID 'used for Edit verification
    matrixID = MyID 'saving data
    Screen.MousePointer = vbHourglass
    
    pBar.Visible = True
    pBar.Value = 0
    
    DoEvents
    Call modPMBranchSQTY.sp_brnchQty_getDetails(MyBranchCode, MyID, mainRs)
    If mainRs.State <> 1 Then Exit Sub
    
    If mainRs.RecordCount <= 0 Then
        ' & vbCrLf & " If you want to Edit Please go to" & vbCrLf & " View Summary -> Show Form -> Edit"
        MsgBox "There are no Available Parts Group has assigned" & vbCrLf & " If you want to Edit Please go to" & vbCrLf & " View Summary -> Show Form -> Edit", vbInformation + vbOKOnly, "Show Details"
        cmbModel.Text = ""
        pBar.Visible = False
        pBar.Value = 0
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    With mainRs
        If .RecordCount > 0 Then
            .MoveFirst
            flxBSQTY.Rows = .RecordCount + 1
            Call resetVar
            pBar.Max = .RecordCount
            
            For i = 1 To flxBSQTY.Rows - 1
            intTotalSTDQ = intTotalSTDQ + IIf(IsNull(!PMSTDQ), 0, !PMSTDQ)

            intComputeLacking = 0
            intComputeExcess = 0
            intSTDQty = 0

            'get LACKING and EXCESS
            If !QtyOnHand <= IIf(IsNull(!PMSTDQ), 0, !PMSTDQ) Then
                intComputeLacking = IIf(IsNull(!PMSTDQ), 0, !PMSTDQ) - !QtyOnHand
            ElseIf !QtyOnHand >= IIf(IsNull(!PMSTDQ), 0, !PMSTDQ) Then
                intComputeExcess = !QtyOnHand - IIf(IsNull(!PMSTDQ), 0, !PMSTDQ)
            End If
            
            NewPrtsGroup = !partsGroup
            
            If NewPrtsGroup <> OldPrtsGroup Then
                intPG = intPG + 1
                flxPG.Rows = intPG + 1
                flxPG.TextMatrix(intPG, 1) = NewPrtsGroup
            End If
            
             flxBSQTY.TextMatrix(i, 1) = !PartNumber
             flxBSQTY.TextMatrix(i, 2) = !Description
             flxBSQTY.TextMatrix(i, 3) = !partsGroup
             flxBSQTY.TextMatrix(i, 4) = !PMMODEL
             flxBSQTY.TextMatrix(i, 5) = !QtyOnHand
             flxBSQTY.TextMatrix(i, 6) = !PM3MSQty
             flxBSQTY.TextMatrix(i, 7) = IIf(IsNull(!PMNBS), 0, !PMNBS)
             flxBSQTY.TextMatrix(i, 8) = IIf(IsNull(!PMSTDQ), 0, !PMSTDQ)
             flxBSQTY.TextMatrix(i, 9) = Format$(intComputeLacking, 0)
             flxBSQTY.TextMatrix(i, 10) = Format$(intComputeExcess, 0)
             flxBSQTY.TextMatrix(i, 11) = Format$(!PMORDERFREQ, "")
             flxBSQTY.TextMatrix(i, 12) = !SPAREPARTID
             flxBSQTY.TextMatrix(i, 13) = !partsGroup
            .MoveNext
            pBar.Value = pBar.Value + 1
            OldPrtsGroup = NewPrtsGroup
            Next i

            For i = 1 To flxPG.Rows - 1
                If i > 0 Then
                    strPG = strPG & flxPG.TextMatrix(i, 1) & "/"
                    intComputePG = intComputePG + 1
                End If
            Next
            strPG = Left(strPG, Len(strPG) - 1) ' To remove the last slash
            lblPartsGroup.Caption = strPG & " " & "ONLY"
        End If
    End With
        
    If mainRs.State <> 0 Then
        Screen.MousePointer = vbDefault
        mainRs.Close
    End If
    
    ' Showrecord of BRAND,MODEL TOTAL PN and etc.
    DoEvents
    Call modPMBranchSQTY.sp_brnchQty_getDetails_total(MyBranchCode, matrixID, tmpRs)

    If tmpRs.State <> 1 Then Exit Sub

    If tmpRs.RecordCount <= 0 Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

        With tmpRs
            If .RecordCount > 0 Then
                .MoveFirst
        '****************** CODE UPDATED: 04/08/17 ************************
                Call ComputeTotal
                'Percentage Calculation
                Call calculatePercentageShowRecord
        
                cmbBranch.Text = IIf(IsNull(!BNAME), "", !BNAME)
                txtTTPN.Text = Format$(intComputeTtlPN, 0)
                txtTTSTDQ.Text = Format$(intComputeTtlSTDQ, 0)
                txtPG.Text = Format$(intComputePG, 0)
                txtExcess.Text = Format$(strTtlExcess, "0.0%")
                txtLacking.Text = Format$(strTtlLacking, "0.0%")
                txtEmpty.Text = Format$(strTtlEmpty, "0.0%")
                txtUpdated.Text = IIf(IsNull(!UPDATEDBY), UserName, !UPDATEDBY)
                cmbBrand.Text = !PMBRAND
                cmbModel.Text = !PMMODEL
            End If
        '************** END CODE UPDATED: 04/08/17 ************************
        End With
     If tmpRs.State <> 0 Then
        Screen.MousePointer = vbDefault
        tmpRs.Close
    End If
    
    pBar.Visible = False
    pBar.Value = 0
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->branchQtyShowRecord()")
        
End Sub

'****************** CODE UPDATED: 04/08/17 ************************
Public Sub ComputeTotal()
Dim cTLacking As Integer, cTExcess As Integer, cTSTDQ As Integer, cTEmpty As Integer, cTQtyOnHand As Integer
Dim cTPN As Integer

    cTPN = 0
    intComputeTtlPN = 0
    cTQtyOnHand = 0
    intComputeTtlQtyOnHand = 0
    cTLacking = 0
    intComputeTtlLacking = 0
    cTExcess = 0
    intComputeTtlExcess = 0
    cTSTDQ = 0
    intComputeTtlSTDQ = 0
    cTEmpty = 0
    intComputeTtlEmpty = 0
    
    For i = 1 To flxBSQTY.Rows - 1
        If Val(flxBSQTY.TextMatrix(i, 5)) = 0 Then
            cTEmpty = cTEmpty + 1
        End If
        If flxBSQTY.TextMatrix(i, 1) <> "" Then
            cTPN = cTPN + 1
        End If
        cTQtyOnHand = cTQtyOnHand + Val(Replace(flxBSQTY.TextMatrix(i, 5), ",", ""))
        cTSTDQ = cTSTDQ + Val(Replace(flxBSQTY.TextMatrix(i, 7), ",", ""))
        cTLacking = cTLacking + Val(Replace(flxBSQTY.TextMatrix(i, 8), ",", ""))
        cTExcess = cTExcess + Val(Replace(flxBSQTY.TextMatrix(i, 9), ",", ""))
    Next i
    intComputeTtlQtyOnHand = cTQtyOnHand
    intComputeTtlSTDQ = cTSTDQ
    intComputeTtlLacking = cTLacking
    intComputeTtlExcess = cTExcess
    intComputeTtlEmpty = cTEmpty
    intComputeTtlPN = cTPN
End Sub
'************** END CODE UPDATED: 04/08/17 ************************

'****************** CODE UPDATED: 04/08/17 ************************
Public Sub calculatePercentageShowRecord()
    strTtlQOH = intComputeTtlQtyOnHand
    'Excess Percentage
    intExcess = Format$(intComputeTtlExcess, 0)
    If intExcess <> 0 Then
        If intExcess <= strTtlQOH Then
            strTtlExcess = intExcess / strTtlQOH * 100 & "%"
        ElseIf intExcess >= strTtlQOH Then
            strTtlExcess = strTtlQOH / intExcess * 100 & "%"
        End If
     Else
        strTtlExcess = "0.0%"
    End If
    
    'Lacking Percentage
    intLacking = Format$(intComputeTtlLacking, 0)
    If intLacking <> 0 Then
        If Val(intLacking) <= intComputeTtlSTDQ Then
            strTtlLacking = intLacking / Format$(intComputeTtlSTDQ, 0) * 100 & "%"
        ElseIf Val(intLacking) >= intComputeTtlSTDQ Then
            strTtlLacking = Format$(intComputeTtlSTDQ, 0) / intLacking * 100 & "%"
        End If
    Else
        strTtlLacking = "0.0%"
    End If
    
    'Empty Percentage
    intEmpty = Format$(intComputeTtlEmpty, 0)
    If intEmpty <> 0 Then
        If intEmpty <= intComputeTtlPN Then
            strTtlEmpty = intEmpty / intComputeTtlPN * 100 & "%"
        ElseIf intEmpty >= intComputeTtlPN Then
            strTtlEmpty = intComputeTtlPN / intEmpty * 100 & "%"
        End If
    Else
        strTtlEmpty = "0.0%"
    End If
    
End Sub
'************** END CODE UPDATED: 04/08/17 ************************

Private Sub txtFind_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFind, KeyAscii, ValidChars)
    
    If KeyAscii = 13 Then
        Call btnFind_Click
    End If
    
End Sub


'Private Sub flxBSQTY_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'
'    If Button = vbRightButton Then
'        PopupMenu mnuStatus, vbPopupMenuRightButton
'    End If
'
'End Sub


'Private Sub mnuAdd_Click()
'
'    If flxBSQTY.TextMatrix(actrow, 1) <> flxBSQTY.TextMatrix(flxBSQTY.Rows - 1, 1) Then
'        flxBSQTY.SetFocus
'        Exit Sub
'    End If
'
'    If flxBSQTY.TextMatrix(flxBSQTY.Rows - 1, 5) = "" Or flxBSQTY.TextMatrix(flxBSQTY.Rows - 1, 9) = "" Then
'        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
'        flxBSQTY.SetFocus
'        'PrtAdd = False
'        Exit Sub
'    End If
'
'    flxBSQTY.AddItem ""
'    flxBSQTY.SetFocus
'    flxBSQTY.Row = flxBSQTY.Rows - 1
'    flxBSQTY.Col = 1
'    txtTTPN.Text = flxBSQTY.Rows - 1
'
'End Sub


'Private Sub mnuDelete_Click()
'
'    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
'    If Ans = vbYes Then
'
'        If flxBSQTY.Rows = 2 Then
'            flxBSQTY.Rows = 1
'            flxBSQTY.Rows = 2
'        Else
'            flxBSQTY.RemoveItem actrow
'            For i = 1 To flxBSQTY.Rows - 1
'                  If flxBSQTY.TextMatrix(i, 1) <> "" Then
'                  txtTTPN.Text = i
'                 End If
'             Next i
'        End If
'    End If
'
'End Sub


Private Sub txtSTDQ_KeyDown(KeyCode As Integer, Shift As Integer)
      
    If KeyCode = vbKeyEscape Then
            Tmp = Split(txtSTDQ.Tag, "|")
        
            txtSTDQ.Visible = False
            flxBSQTY.SetFocus
            flxBSQTY.Col = 7
    End If
End Sub

Public Sub DataPercentage()
On Error GoTo er

    intSTDQ = 0
    intLacking = 0
    intExcess = 0
    intEmpty = 0
    intQOH = 0
    strTtlEmpty = 0
    strTtlExcess = 0
    strTtlLacking = 0
    'Getting the total of Standard Quantity
    For LoopSTDQ = 1 To flxBSQTY.Rows - 1
        intSTDQ = intSTDQ + Val(flxBSQTY.TextMatrix(LoopSTDQ, 7))
    Next LoopSTDQ
    txtTTSTDQ.Text = intSTDQ
    
    
    'Getting the Total of Quantity of Hand
    For LoopQOH = 1 To flxBSQTY.Rows - 1
        intQOH = intQOH + Val(flxBSQTY.TextMatrix(LoopQOH, 5))
    Next LoopQOH
    intTotlQOH = intQOH
    
    
    'Getting the Total Percentage of Lacking
    For LoopLACKING = 1 To flxBSQTY.Rows - 1
        intLacking = intLacking + Val(flxBSQTY.TextMatrix(LoopLACKING, 8))
    Next LoopLACKING
    
    strTtlLacking = intLacking / intSTDQ * 100 & "%"
    txtLacking.Text = Format$(strTtlLacking, "0.0%")
    
    
    'Getting the Total Percentage of Empty
    For LoopEmpty = 1 To flxBSQTY.Rows - 1
              
        If flxBSQTY.TextMatrix(LoopEmpty, 5) <> "" Then
            If flxBSQTY.TextMatrix(LoopEmpty, 5) = 0 Then
            intEmpty = intEmpty + 1
            
            Else
            
            GoTo SKIP
            End If
            
        End If
SKIP:
intFlxEmpty = LoopEmpty
    Next LoopEmpty
    
    strTtlEmpty = intEmpty / intFlxEmpty * 100 & "%"
    txtEmpty.Text = Format$(strTtlEmpty, "0.0%")
    
    'Getting the Total Percentage of Excess
    For LoopExcess = 1 To flxBSQTY.Rows - 1
        intExcess = intExcess + Val(flxBSQTY.TextMatrix(LoopExcess, 9))
    Next LoopExcess
     If intExcess <> intTotlQOH Then
    strTtlExcess = intExcess / intTotlQOH * 100 & "%"
    txtExcess.Text = Format$(strTtlExcess, "0.0%")
      Else
       Exit Sub
     End If
     
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->DataPercentage()")
End Sub

Private Sub txtSTDQ_KeyPress(KeyAscii As Integer)
On Error GoTo er

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtSTDQ, KeyAscii, NumbersOnly)

    If KeyAscii = 13 Then

        If flxBSQTY.TextMatrix(ActRow, 1) <> "" Then
            flxBSQTY.TextMatrix(ActRow, 7) = txtSTDQ
            
       
            'QOH > STDQ = EXCESS
            If Val(flxBSQTY.TextMatrix(ActRow, 5)) >= Val(flxBSQTY.TextMatrix(ActRow, 7)) Then
                    
                    flxLack = Val(flxBSQTY.TextMatrix(ActRow, 5)) - Val(flxBSQTY.TextMatrix(ActRow, 7))
                  
                   
                   flxBSQTY.TextMatrix(ActRow, 9) = flxLack
                   flxBSQTY.TextMatrix(ActRow, 8) = 0
                   
            Else
            
            'QOH < STDQ = LACKING
                flxLack = Val(flxBSQTY.TextMatrix(ActRow, 7)) - Val(flxBSQTY.TextMatrix(ActRow, 5))
                flxBSQTY.TextMatrix(ActRow, 9) = 0
                flxBSQTY.TextMatrix(ActRow, 8) = flxLack
                
            End If
    
            If txtSTDQ.Text <> 0 Then
    
                Call DataPercentage
            
                For i = ActRow To ActRow + 5
                    If flxBSQTY.TextMatrix(i, 7) = "" Or Val(flxBSQTY.TextMatrix(i, 7)) = 0 Then
                        txtSTDQ.Visible = False
                        flxBSQTY.SetFocus
                        flxBSQTY.Row = i
                        flxBSQTY.Col = 7
                        Exit For
                    Else
                        txtSTDQ.Visible = False
                        flxBSQTY.SetFocus
                        flxBSQTY.Row = ActRow
                        flxBSQTY.Col = 7
                    End If
                Next i
                
            Else
                flxBSQTY.SetFocus
                flxBSQTY.Col = 7
                txtSTDQ.Visible = False
            
        End If
       
    End If
      
End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtSTDQ-KeyPress()")
End Sub

'Sub LockData(DataEntry As Boolean)
'    DataEntry = Not DataEntry
'
'    cmbBranch.Locked = DataEntry
'    cmbBrand.Locked = DataEntry
'    cmbModel.Locked = DataEntry
'    txtTTPN.Locked = DataEntry
'    txtTTSTDQ.Locked = DataEntry
'    txtPG.Locked = DataEntry
'    txtExcess.Locked = DataEntry
'    txtLacking.Locked = DataEntry
'    txtEmpty.Locked = DataEntry
'    txtUpdated.Locked = DataEntry
'
'    If cmbBrand.Locked Then
'        cmbBranch.BackColor = txtHoldGrayBackColor.BackColor
'        cmbBrand.BackColor = txtHoldGrayBackColor.BackColor
'        cmbModel.BackColor = txtHoldGrayBackColor.BackColor
'        txtTTPN.BackColor = txtHoldGrayBackColor.BackColor
'        txtTTSTDQ.BackColor = txtHoldGrayBackColor.BackColor
'        txtPG.BackColor = txtHoldGrayBackColor.BackColor
'        txtExcess.BackColor = txtHoldGrayBackColor.BackColor
'        txtLacking.BackColor = txtHoldGrayBackColor.BackColor
'        txtEmpty.BackColor = txtHoldGrayBackColor.BackColor
'        txtUpdated.BackColor = txtHoldGrayBackColor.BackColor
'
'    Else
'        cmbBranch.BackColor = txtHoldWhiteBackColor.BackColor
'        cmbBrand.BackColor = txtHoldWhiteBackColor.BackColor
'        cmbModel.BackColor = txtHoldWhiteBackColor.BackColor
''        txtTTPN.BackColor = txtHoldWhiteBackColor.BackColor
''        txtTTSTDQ.BackColor = txtHoldWhiteBackColor.BackColor
''        txtPG.BackColor = txtHoldWhiteBackColor.BackColor
''        txtExcess.BackColor = txtHoldWhiteBackColor.BackColor
''        txtLacking.BackColor = txtHoldWhiteBackColor.BackColor
''        txtEmpty.BackColor = txtHoldWhiteBackColor.BackColor
''        txtUpdated.BackColor = txtHoldWhiteBackColor.BackColor
'
'    End If
'End Sub

Private Sub txtSTDQ_LostFocus()
On Error GoTo er

    If txtSTDQ.Tag = "skiplostfocus" Then
        txtSTDQ.Tag = ""
        Exit Sub
    End If
    
    If flxBSQTY.TextMatrix(ActRow, 7) = "" Or txtSTDQ.Text = "" Then
        MsgBox "Wrong Entry of Standard Quantity!" & vbCrLf & "Please Enter a Correct Standard Quantity.", vbExclamation + vbOKOnly, "Warning"
        txtSTDQ.Visible = False

        flxBSQTY.TextMatrix(ActRow, 7) = ""
        flxBSQTY.SetFocus
        flxBSQTY.Col = 7
        
        Exit Sub
    End If
    
    flxBSQTY.TextMatrix(ActRow, 7) = txtSTDQ.Text
    txtSTDQ.Visible = False
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtSTDQ-LostFocus()")
End Sub

