VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form BranchIdentifying 
   Caption         =   "Branch Identifying"
   ClientHeight    =   10290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19185
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10290
   ScaleWidth      =   19185
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtpmID 
      Height          =   285
      Left            =   19320
      TabIndex        =   35
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtALQ 
      Height          =   285
      Left            =   19320
      TabIndex        =   16
      Tag             =   "XYF"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   19320
      TabIndex        =   30
      Text            =   "HoldBackColor"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   19320
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Height          =   6855
      Left            =   120
      TabIndex        =   28
      Tag             =   "XYF"
      Top             =   2400
      Width           =   18975
      Begin MSFlexGridLib.MSFlexGrid flxSPARTS 
         Height          =   6495
         Left            =   120
         TabIndex        =   14
         Tag             =   "XYF"
         Top             =   240
         Width           =   18735
         _ExtentX        =   33046
         _ExtentY        =   11456
         _Version        =   393216
         Rows            =   3
         Cols            =   13
         FixedRows       =   2
         BackColorBkg    =   14737632
         AllowUserResizing=   1
         FormatString    =   "|^MODEL BRAND (PM) |^PARTNUMBER |^DESCRIPTION|^PARTSGROUP|^QH |^STDQ |^SRQ |^QH |^RO |^AVQ |^ALQ |<LOCATION"
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
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   26
      Tag             =   "XYF"
      Top             =   9240
      Width           =   18975
      Begin VB.TextBox txtProcsBy 
         Height          =   285
         Left            =   14760
         TabIndex        =   15
         Tag             =   "XYF"
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label lblProcBy 
         Alignment       =   1  'Right Justify
         Caption         =   "PROCESSED BY"
         Height          =   255
         Left            =   12720
         TabIndex        =   27
         Tag             =   "XYF"
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   17
      Tag             =   "XYF"
      Top             =   480
      Width           =   18975
      Begin VB.ComboBox cboModel 
         Height          =   315
         Left            =   1200
         TabIndex        =   32
         Top             =   1320
         Width           =   1575
      End
      Begin VB.ComboBox cboBrand 
         Height          =   315
         Left            =   1200
         TabIndex        =   31
         Top             =   840
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtDateFinal 
         Height          =   255
         Left            =   17280
         TabIndex        =   13
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDatePrepared 
         Height          =   255
         Left            =   17280
         TabIndex        =   12
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtRefNo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   17280
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtTotalALQ 
         Height          =   285
         Left            =   6240
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtTotalAVQ 
         Height          =   285
         Left            =   6240
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtTotalSRQ 
         Height          =   285
         Left            =   3840
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtTotalPN 
         Height          =   285
         Left            =   3840
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cboBranch 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   360
         Width           =   6255
      End
      Begin VB.Label lblMCMODEL 
         Alignment       =   1  'Right Justify
         Caption         =   "MC MODEL"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblMCBRAND 
         Alignment       =   1  'Right Justify
         Caption         =   "MC BRAND"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblDateFinal 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE FINAL"
         Height          =   255
         Left            =   15840
         TabIndex        =   25
         Tag             =   "XYF"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblDatePrepared 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE PREPARED"
         Height          =   255
         Left            =   15720
         TabIndex        =   24
         Tag             =   "XYF"
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblRefNo 
         Alignment       =   1  'Right Justify
         Caption         =   "REF. NO"
         Height          =   255
         Left            =   15720
         TabIndex        =   23
         Tag             =   "XYF"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblTOTALALQ 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL ALQ"
         Height          =   255
         Left            =   5160
         TabIndex        =   22
         Tag             =   "XYF"
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblTOTALAVQ 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL AVQ"
         Height          =   255
         Left            =   5160
         TabIndex        =   21
         Tag             =   "XYF"
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblTOTALSRQ 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL SRQ"
         Height          =   255
         Left            =   2760
         TabIndex        =   20
         Tag             =   "XYF"
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblTOTALPN 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL PN"
         Height          =   255
         Left            =   2760
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblBranch 
         Alignment       =   1  'Right Justify
         Caption         =   "BRANCH"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Tag             =   "XYF"
         Top             =   360
         Width           =   975
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Tag             =   "XYF"
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "&Add "
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
      Image           =   "Identifying.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Tag             =   "XYF"
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
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
      Image           =   "Identifying.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3960
      TabIndex        =   2
      Tag             =   "XYF"
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
      Image           =   "Identifying.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5880
      TabIndex        =   3
      Tag             =   "XYF"
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "E&xit Parts Matrix"
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
      Image           =   "Identifying.frx":10CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   19320
      TabIndex        =   4
      Top             =   1080
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
      Image           =   "Identifying.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   19320
      TabIndex        =   5
      Top             =   1560
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
      Image           =   "Identifying.frx":1C02
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   19320
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   19305
      m_MDI_Height    =   10800
   End
   Begin VB.Menu mnuStatus 
      Caption         =   "Popupmenu"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "AddRow"
      End
   End
End
Attribute VB_Name = "BranchIdentifying"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strPartial As String, strTotal As String
Dim formAction As String
Dim Ans As String
Dim branchCode As String
Dim branchName As String


Dim Tmp() As String
Dim MyErr As String, myMsg As String
Dim Qry As String, Qry2 As String
Dim intID As Integer, strBoolean As String
Dim AVQ As Integer, intQH As Integer, intRO As Integer, intSRQ As Integer, intTOTALSRQ As Integer
Dim ALQ As Integer, intALQ As Integer
Public mySrc As String

Dim arrMODEL() As String
Dim arrBRAND() As String
Dim arrBRANCH() As String

Dim intTotalALQ As Long
Dim intTotalAVQ As Long
Dim i As Long, j As Long
Dim ActRow As Integer, ActCol As Integer
Dim matrixID As Integer
Dim intREFID As Integer
Dim SEQNO As Integer
Dim intTtlPN As Integer
Dim tmpRs  As New ADODB.Recordset
Dim mainRs As New ADODB.Recordset


Dim StatSrch As Boolean, AddNew As Boolean, ProperExit As Boolean

Public Sub flxWidth()

    With flxSPARTS
        '.Clear
        .Cols = 15
        '.FormatString = "|^MODEL BRAND (PM) |^PARTNUMBER |^DESCRIPTION |^QOH |^STDQ |^SRQ |^QOH |^RO |^AVQ |^ALQ |<LOCATION |^ID"
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.12  'MODEL BRAND (PM)
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.12 'PARTNUMBER
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.17 'DESCRIPTION
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.09 'PARTSGROUP
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.06 'QOH
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.06 'STDQ
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.06 'SRQ
        .ColWidth(8) = (.Width - .ColWidth(0)) * 0.06 'QOH
        .ColWidth(9) = (.Width - .ColWidth(0)) * 0.06 'RO
        .ColWidth(10) = (.Width - .ColWidth(0)) * 0.06 'AVQ
        .ColWidth(11) = (.Width - .ColWidth(0)) * 0.06  ' ALQ
        .ColWidth(12) = (.Width - .ColWidth(0)) * 0.06  ' LOCATION
        .ColWidth(13) = 0 ' (.Width - .ColWidth(0)) * 0.01 ' PMATRIXID
        .ColWidth(14) = 0 ' (.Width - .ColWidth(0)) * 0.1 ' SPAREPARTID
    End With
End Sub

Public Sub flxColChanger()

With flxSPARTS
    If mySrc = "From Summary" Then
        'MERGE COL
    .TextMatrix(0, 0) = " "
    .TextMatrix(1, 0) = " "
    'MERGE ROW
    .TextMatrix(0, 1) = "SPAREPARTS"
    .TextMatrix(0, 2) = "SPAREPARTS"
    .TextMatrix(0, 3) = "SPAREPARTS"
    .TextMatrix(0, 4) = "SPAREPARTS"
    .TextMatrix(0, 5) = cboBranch.Text
    .TextMatrix(0, 6) = cboBranch.Text
    .TextMatrix(0, 7) = cboBranch.Text
    .TextMatrix(0, 8) = "WAREHOUSE"
    .TextMatrix(0, 9) = "WAREHOUSE"
    .TextMatrix(0, 10) = "WAREHOUSE"
    .TextMatrix(0, 11) = "WAREHOUSE"
    .TextMatrix(0, 12) = "WAREHOUSE"
    
    'MERGE COL =FALSE
    .TextMatrix(1, 1) = "MODEL BRAND (PM)"
    .TextMatrix(1, 2) = "PARTNUMBER"
    .TextMatrix(1, 3) = "DESCRIPTION"
    .TextMatrix(1, 4) = "PARTSGROUP"
    .TextMatrix(1, 5) = "QH"
    .TextMatrix(1, 6) = "STDQ"
    .TextMatrix(1, 7) = "SRQ"
    .TextMatrix(1, 8) = "QH"
    .TextMatrix(1, 9) = "RO"
    .TextMatrix(1, 10) = "AVQ"
    .TextMatrix(1, 11) = "ALQ"
    .TextMatrix(1, 12) = "LOCATION"

    'SET MERGING
    .MergeCells = 1
    .MergeRow(0) = True
    .MergeCol(0) = True
    .MergeCol(1) = True
    For i = 1 To 3
        .MergeCol(i) = False
    Next i
    'SET THE COLUMN COLOR GREEN
    .Row = 0
    .Col = 0
    .Col = 1
   ' .CellBackColor = vbGreen

    Else
    'MERGE COL
    .TextMatrix(0, 0) = " "
    .TextMatrix(1, 0) = " "
    'MERGE ROW
    .TextMatrix(0, 1) = "SPAREPARTS"
    .TextMatrix(0, 2) = "SPAREPARTS"
    .TextMatrix(0, 3) = "SPAREPARTS"
    .TextMatrix(0, 4) = "SPAREPARTS"
    .TextMatrix(0, 5) = "BRANCH NAME"
    .TextMatrix(0, 6) = "BRANCH NAME"
    .TextMatrix(0, 7) = "BRANCH NAME"
    .TextMatrix(0, 8) = "WAREHOUSE"
    .TextMatrix(0, 9) = "WAREHOUSE"
    .TextMatrix(0, 10) = "WAREHOUSE"
    .TextMatrix(0, 11) = "WAREHOUSE"
    .TextMatrix(0, 12) = "WAREHOUSE"
    
    'MERGE COL =FALSE
    .TextMatrix(1, 1) = "MODEL BRAND (PM)"
    .TextMatrix(1, 2) = "PARTNUMBER"
    .TextMatrix(1, 3) = "DESCRIPTION"
    .TextMatrix(1, 4) = "PARTSGROUP"
    .TextMatrix(1, 5) = "QH"
    .TextMatrix(1, 6) = "STDQ"
    .TextMatrix(1, 7) = "SRQ"
    .TextMatrix(1, 8) = "QH"
    .TextMatrix(1, 9) = "RO"
    .TextMatrix(1, 10) = "AVQ"
    .TextMatrix(1, 11) = "ALQ"
    .TextMatrix(1, 12) = "LOCATION"

    'SET MERGING
    .MergeCells = 1
    .MergeRow(0) = True
    .MergeCol(0) = True
    .MergeCol(1) = True
    For i = 1 To 3
        .MergeCol(i) = False
    Next i
    'SET THE COLUMN COLOR GREEN
    .Row = 0
    .Col = 0
    .Col = 1
   ' .CellBackColor = vbGreen
    End If
End With
    
    Call flxWidth
    
End Sub

Private Sub cboBranch_Click()

    Call cboBranch_Change

    DoEvents
    
    branchCode = arrBRANCH(cboBranch.ListIndex, 0)
    
    cboBrand.Clear
    cboModel.Clear
    flxSPARTS.Rows = 2
    flxSPARTS.Rows = 3
    
    Call CboBrandDisplay
  
    
    With flxSPARTS
        .TextMatrix(0, 5) = cboBranch.Text
        .TextMatrix(0, 6) = cboBranch.Text
        .TextMatrix(0, 7) = cboBranch.Text
        'SET MERGING
        .MergeCells = 1
        .MergeRow(0) = True
        .MergeCol(0) = True
        .MergeCol(1) = True
        For i = 1 To 3
            .MergeCol(i) = False
        Next i
    End With
    
End Sub

Private Sub cboBranch_GotFocus()
Dim Tmp
    Const CB_SHOWDROPDOWN = &H14F
    Tmp = SendMessage(cboBranch.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboBranch_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboBranch, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Call cboBranch_Click
        cboBrand.SetFocus
    End If
End Sub

Private Sub cboBrand_Change()
On Error GoTo er
    If StatSrch Then
            StatSrch = False
            Exit Sub
    End If

    With cboBrand
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

Private Sub cboBrand_Click()

    Call cboBrand_Change
    
    If formAction = "add" Or formAction = "edit" Then
        cboModel.Clear
    End If

    Call CboModelDisplay
End Sub



Private Sub cboBrand_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboBrand, KeyAscii, ValidChars)
    
    If KeyAscii = 13 Then
    
        If formAction = "add" Then
            cboModel.Clear
        End If
    
        Call CboModelDisplay
        cboModel.SetFocus
        
    End If
End Sub

Private Sub cboModel_Change()
On Error GoTo er

    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If
    
    With cboModel
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboModel()")
End Sub

Private Sub cboModel_Click()
    
    Call cboModel_Change
    
    DoEvents
    matrixID = arrMODEL(cboModel.ListIndex, 0)
    txtpmID.Text = matrixID
    Call ShowRecord(branchCode, matrixID, "FALSE")

    txtDatePrepared.SetFocus

End Sub



Private Sub cboModel_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboModel, KeyAscii, ValidChars)
    
    If KeyAscii = 13 Then

        txtRefNo.SetFocus
        
    End If
End Sub


Private Sub cmdTrans_Click(Index As Integer)
On Error GoTo er

    If Index = 0 Then 'add
    
        formAction = "add"
        AddNew = True
      
        Call clearform
        Call CboDisplay
        Call LockData(True)
        Call SetButtonVisibility(True)
        txtProcsBy.Text = EmployeesName
        
        
        
    ElseIf Index = 1 Then 'edit
        If txtpmID = "" Then
            MsgBox "No available PARTS MATRIX to edit", vbOKOnly + vbExclamation, "Edit"

        Else
        
            formAction = "edit"
            AddNew = False
            Call SetButtonVisibility(True)
            Call LockData(True)
            cboBranch.BackColor = txtHoldGrayBackColor.BackColor
            cboBrand.BackColor = txtHoldGrayBackColor.BackColor
            cboModel.BackColor = txtHoldGrayBackColor.BackColor
            txtProcsBy.Text = EmployeesName
            flxSPARTS.SetFocus
            flxSPARTS.Row = 2
            flxSPARTS.Col = 11
            
        End If
        
    ElseIf Index = 2 Then 'view summary
        
        frmSummaryRepl.mySrc = "Branch Identifying"
        frmSummaryRepl.Show
        Unload Me
        
    ElseIf Index = 3 Then 'exit
        
        formAction = "EXIT"
        ProperExit = True
        Call clearform
        Unload Me
        
    ElseIf Index = 4 Then 'save
    
        myMsg = ValidationPartsMatrix
        
        If myMsg <> "" Then
            MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
            Exit Sub
        
        End If
        
        Ans = MsgBox("Do you want to save?", vbQuestion + vbYesNo, "Save")
        
        If Ans = vbNo Then Exit Sub
            
            
        If formAction = "add" Then
                Call CreateTempTable
                DoEvents
                Call modPMBranchIdentifying.sp_brnchIdentify_create(branchCode, matrixID, Format(txtDatePrepared.Text, "YYYY-mm-dd"), Format(txtDateFinal.Text, "YYYY-mm-dd"), UserID, UserID, 0, mainRs)
                txtRefNo.Text = mainRs!RefNo
        Else
            Call CreateTempTable
            DoEvents
            Call modPMBranchIdentifying.sp_brnchIdentify_update(intREFID, txtRefNo.Text, Format(txtDatePrepared.Text, "YYYY-mm-dd"), Format(txtDateFinal.Text, "YYYY-mm-dd"), UserID, mainRs)
            Call ShowRecord(branchCode, intID, strBoolean) 'to refresh
        End If
            Call SetButtonVisibility(False)
            Call LockData(False)
            formAction = ""
            
        
        
    ElseIf Index = 5 Then 'cancel
        
        Ans = MsgBox("Do you want to cancel?", vbQuestion + vbYesNo)
        
        If Ans = vbNo Then Exit Sub
        
        Call SetButtonVisibility(False)
        Call LockData(False)
'        Call clearform
        formAction = ""
        If AddNew Then
            branchCode = ""
        End If
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdTrans-Click()")
End Sub

Sub CreateTempTable()
On Error GoTo er
        LYRcon.Execute "Drop Temporary Table IF  Exists sp_partsmatrix.TMP_BRNCHIDENTIFY_DTLS;"

         DoEvents
       LYRcon.Execute " Create Temporary Table IF Not Exists sp_partsmatrix.TMP_BRNCHIDENTIFY_DTLS( " & _
        "    PMALQ Int(11)            " & _
        "   ,PMATRIXID Int(11)        " & _
        "   ,SPAREPARTID Int(11)      " & _
        "   ,REMARKS Varchar(11)      " & _
        "   ,RO INT(11)      " & _
        "   ,SEQNO INT(11)      " & _
        ");"
        
        Qry = "Insert into sp_partsmatrix.TMP_BRNCHIDENTIFY_DTLS(PMALQ,PMATRIXID,SPAREPARTID,REMARKS,RO,SEQNO) VALUES "
        SEQNO = 0
        For i = 2 To flxSPARTS.Rows - 1
            If i > 2 Then Qry = Qry & ","
            SEQNO = SEQNO + 1
           Qry = Qry & "(" & Val(flxSPARTS.TextMatrix(i, 11)) & "," & Val(flxSPARTS.TextMatrix(i, 13)) & "," & Val(flxSPARTS.TextMatrix(i, 14)) & ", '" & "RESERVED" & "'," & Val(flxSPARTS.TextMatrix(i, 9)) & "," & SEQNO & ") "
    
        Next i
        
       
       Call LYRcon.Execute(Qry)
       
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CreateTempTable()")
End Sub
Private Sub flxSPARTS_Click()
On Error GoTo er

    ActRow = flxSPARTS.Row
    ActCol = flxSPARTS.Col
    
    If formAction = "add" Or formAction = "edit" Then
    
    If ActCol = 0 Or ActRow = 0 Then Exit Sub
        Select Case ActCol
        
        Case 11
            If flxSPARTS.TextMatrix(flxSPARTS.Row, 1) <> "" Then
                With txtALQ
                    If flxSPARTS.TextMatrix(ActRow, 6) <> "" Then .Tag = flxSPARTS.TextMatrix(ActRow, 6)
                    .Top = Frame3.Top + flxSPARTS.CellTop + flxSPARTS.Top
                    .Left = Frame3.Left + flxSPARTS.CellLeft + flxSPARTS.Left
                    .Height = flxSPARTS.CellHeight
                    .Width = flxSPARTS.CellWidth
                    .Text = flxSPARTS.Text
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .ZOrder
                    .SetFocus
                End With
            Else
                Exit Sub
            End If
            
        End Select
    End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->flxSPARTS-click()")
End Sub

Private Sub flxSPARTS_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then Call flxSPARTS_Click

End Sub

Private Sub flxSPARTS_Scroll()
txtALQ.Visible = False
End Sub

Private Sub Form_Load()
    Call flxColChanger
   ' Call modCon.Connection
    Call LockData(False)
    
    txtTotalALQ.Text = 0
    txtRefNo.Text = "<Auto>"
    txtDatePrepared.Text = Format(Date, "mm/dd/yyyy")
    txtDateFinal.Text = Format(Date, "mm/dd/yyyy")
    'txtProcsBy.Text = Warehouse.sBar.Panels(3).Text
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


Private Sub cboBranch_Change()
On Error GoTo er
    If StatSrch Then
            StatSrch = False
            Exit Sub
    End If

    With cboBranch
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboBranch-Change()")
End Sub

Public Sub CboDisplay()
On Error GoTo er
    If formAction = "add" Then
        'Get Data from database
        DoEvents
        Call modPMBranchIdentifying.sp_get_Branch(tmpRs)
        If tmpRs.State = 1 Then
           Call modObjTools.addToListObjectWithArrayGivenField(cboBranch, tmpRs, arrBRANCH, "BRANCHCODE", "BRANCHNAME")
        End If
        If tmpRs.State <> 0 Then tmpRs.Close
    End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub

Public Sub CboBrandDisplay()
On Error GoTo er
    If formAction = "add" Then
        'Get Data from database
        DoEvents
        Call modPMBranchIdentifying.get_Brand(branchCode, tmpRs)
        If tmpRs.State = 1 Then
            Call modObjTools.addToListObjectWithArrayGivenField(cboBrand, tmpRs, arrBRAND, "nulls", "PMBRAND")
        End If
        If tmpRs.State <> 0 Then tmpRs.Close
    End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboBrandDisplay()")
End Sub

Public Sub CboModelDisplay()
On Error GoTo er
    If formAction = "add" Then
        DoEvents
        Call modPMBranchIdentifying.get_Model(cboBrand.Text, branchCode, tmpRs)
        If tmpRs.State = 1 Then
            Call modObjTools.addToListObjectWithArrayGivenField(cboModel, tmpRs, arrMODEL, "PMATRIXID", "PMMODEL")
        End If
        If tmpRs.State <> 0 Then tmpRs.Close
    End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboModelDisplay()")
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
    Call flxColChanger

    'Frame width
    Frame1.Width = Me.Width - 450
    Frame2.Width = Me.Width - 450
    Frame3.Width = Me.Width - 450
  
    'Frame1 label Left
    lblBranch.Left = 120
    lblMCBRAND.Left = 120
    lblMCMODEL.Left = 120
    lblTOTALPN.Left = 2760
    lblTOTALSRQ.Left = 2760
    lblTOTALAVQ.Left = 5160
    lblTOTALALQ.Left = 5160
    lblRefNo.Left = Frame1.Width - 3255
    lblDatePrepared.Left = Frame1.Width - 3255
    lblDateFinal.Left = Frame1.Width - 3255
    
    
    'Frame1 Label Width
    lblBranch.Width = 975
    lblTOTALPN.Width = 975
    lblTOTALSRQ.Width = 975
    lblTOTALAVQ.Width = 975
    lblTOTALALQ.Width = 975
    lblMCBRAND.Width = 975
    lblMCMODEL.Width = 975
    lblRefNo.Width = 1455
    lblDatePrepared.Width = 1455
    lblDateFinal.Width = 1455

    
    'Frame1 Combobox Left
    cboBranch.Left = 1200
    cboBrand.Left = 1200
    cboModel.Left = 1200
    
    'Frame1 Combobox Width
    cboBranch.Width = 6255
    cboBrand.Width = 1575
    cboModel.Width = 1575
    
    'Frame1 Text Left
    txtTOTALPN.Left = 3840
    txtTotalSRQ.Left = 3840
    
    txtTotalAVQ.Left = 6240
    txtTotalALQ.Left = 6240
    txtRefNo.Left = Frame1.Width - 1695
    txtDatePrepared.Left = Frame1.Width - 1695
    txtDateFinal.Left = Frame1.Width - 1695
    
    'Frame1 Text Width
    txtTOTALPN.Width = 1215
    txtTotalSRQ.Width = 1215
    txtTotalAVQ.Width = 1215
    txtTotalALQ.Width = 1215
    txtRefNo.Width = 1335
    txtDatePrepared.Width = 1335
    txtDateFinal.Width = 1335
    
    'Frame2 Label Left
    lblProcBy.Left = Frame3.Width - 6255
    
    'Frame2 Label Width
    lblProcBy.Width = 1935
    
    'Frame2 Text Left
    txtProcsBy.Left = Frame2.Width - 4215
    
    'Frame2 Text Width
    txtProcsBy.Width = 3975
    
    'Frame3 FlexGrid Width
    flxSPARTS.Width = Frame3.Width - 240
    
    'Frame3 FlexGrid Height
    flxSPARTS.Height = Frame3.Height - 360
End Sub


Sub LockData(DataEntry As Boolean)
    DataEntry = Not DataEntry

    cboBranch.Locked = DataEntry
    cboBrand.Locked = DataEntry
    cboModel.Locked = DataEntry
    txtTOTALPN.Locked = DataEntry
    txtTotalSRQ.Locked = DataEntry
    txtTotalAVQ.Locked = DataEntry
    txtTotalALQ.Locked = DataEntry
    txtRefNo.Locked = DataEntry
    'txtDatePrepared.Locked = DataEntry
    'txtDateFinal.Locked = DataEntry
    txtProcsBy.Locked = DataEntry
   
    If cboBranch.Locked Then
        cboBranch.BackColor = txtHoldGrayBackColor.BackColor
        cboBrand.BackColor = txtHoldGrayBackColor.BackColor
        cboModel.BackColor = txtHoldGrayBackColor.BackColor
        txtTOTALPN.BackColor = txtHoldGrayBackColor.BackColor
        txtTotalSRQ.BackColor = txtHoldGrayBackColor.BackColor
        txtTotalAVQ.BackColor = txtHoldGrayBackColor.BackColor
        txtTotalALQ.BackColor = txtHoldGrayBackColor.BackColor
        txtRefNo.BackColor = txtHoldGrayBackColor.BackColor
        txtDatePrepared.BackColor = txtHoldGrayBackColor.BackColor
        txtDateFinal.BackColor = txtHoldGrayBackColor.BackColor
        txtProcsBy.BackColor = txtHoldGrayBackColor.BackColor

    Else
        cboBranch.BackColor = txtHoldWhiteBackColor.BackColor
        cboBrand.BackColor = txtHoldWhiteBackColor.BackColor
        cboModel.BackColor = txtHoldWhiteBackColor.BackColor
        txtTOTALPN.Locked = True
        txtTotalSRQ.Locked = True
        txtTotalAVQ.Locked = True
        txtTotalALQ.Locked = True
        txtRefNo.Locked = True
        txtDatePrepared.BackColor = txtHoldWhiteBackColor.BackColor
        txtDateFinal.BackColor = txtHoldWhiteBackColor.BackColor
        txtProcsBy.BackColor = txtHoldWhiteBackColor.BackColor
        
    End If
End Sub


Private Sub SetButtonVisibility(DataEntry As Boolean)
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->VIEW SUMM
'3 ->EXIT
'4 ->SAVE
'5 ->CANCEL


    'hide all button
    For i = 0 To 4
        cmdTrans(i).Visible = False
    Next i

    If DataEntry Then
        cmdTrans(4).Visible = True
        cmdTrans(5).Visible = True
    Else
    For i = 0 To 4
        cmdTrans(i).Visible = True
        cmdTrans(4).Visible = False
        cmdTrans(5).Visible = False
    Next i
    
    
    End If
End Sub

Sub clearform()

    flxSPARTS.Rows = 2
    flxSPARTS.Rows = 3
    txtpmID.Text = ""
    cboBranch.Text = ""
    cboBrand.Text = ""
    cboModel.Text = ""
    txtTOTALPN.Text = ""
    txtTotalSRQ.Text = ""
    txtTotalAVQ.Text = ""
    txtTotalALQ.Text = 0
    txtRefNo.Text = "<Auto>"
    txtDatePrepared.Text = Format(Date, "mm/dd/yyyy")
    txtDateFinal.Text = Format(Date, "mm/dd/yyyy")
    'txtProcsBy.Text = ""
    branchCode = ""
    
End Sub
Public Sub getREFID(myREFID As Integer)
   intREFID = myREFID
End Sub


Public Sub ShowRecord(MyBranchCode As String, MyID As Integer, myBoolean As String)
On Error GoTo er

    Screen.MousePointer = vbHourglass
    branchCode = MyBranchCode 'use to verify to edit
    txtpmID = MyID
    intID = MyID
    strBoolean = myBoolean
    intTotalAVQ = 0
    intTOTALSRQ = 0
    intTtlPN = 0
    intALQ = 0
 'get Branch Name to Display in ComboBox
 DoEvents
 Call modPMBranchIdentifying.sp_get_BranchName(MyBranchCode, mainRs)
    cboBranch.Text = mainRs!branchName

 'get Details to Display in FlexGrid by PMATIRIXID
 DoEvents
 Call modPMBranchIdentifying.sp_brnchIdentify_getDetails(MyBranchCode, MyID, myBoolean, intREFID, tmpRs)
    
    If tmpRs.State <> 1 Then Exit Sub
    
    If tmpRs.RecordCount <= 0 Then
        MsgBox "This Parts Matrix has already Allocated or No SRQ", vbInformation + vbOKOnly, "Show Details"
        cboModel.Text = ""
        flxSPARTS.Rows = 2
        flxSPARTS.Rows = 3
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    With tmpRs
        

        If .RecordCount > 0 Then
         
            .MoveFirst
            intTtlPN = .RecordCount
            flxSPARTS.Rows = .RecordCount + 2
            
            For i = 2 To flxSPARTS.Rows - 1
            intSRQ = !PMLACKING
            ALQ = !ALLOCATEDQTY
            AVQ = !QH - !RO
            If AVQ < 0 Then
                AVQ = 0
            End If
    
            intALQ = intALQ + ALQ
            intTotalAVQ = intTotalAVQ + AVQ
            intTOTALSRQ = intTOTALSRQ + intSRQ
                flxSPARTS.TextMatrix(i, 1) = !MODELBRAND
                flxSPARTS.TextMatrix(i, 2) = !PartNumber
                flxSPARTS.TextMatrix(i, 3) = !Description
                flxSPARTS.TextMatrix(i, 4) = !partsGroup
                flxSPARTS.TextMatrix(i, 5) = !QtyOnHand 'from mrc2retail.spareparts
                flxSPARTS.TextMatrix(i, 6) = !PMSTDQ
                flxSPARTS.TextMatrix(i, 7) = !PMLACKING
                flxSPARTS.TextMatrix(i, 8) = !QH 'from mrcwarehouse.spareparts
                flxSPARTS.TextMatrix(i, 9) = !RO
                flxSPARTS.TextMatrix(i, 10) = AVQ  '!QH - !RO
                flxSPARTS.TextMatrix(i, 11) = Format$(!ALLOCATEDQTY, "") '!PMALQ
                flxSPARTS.TextMatrix(i, 12) = Format$(!Location, "") 'from mrcwarehouse.spareparts
                flxSPARTS.TextMatrix(i, 13) = !PMATRIXID
                flxSPARTS.TextMatrix(i, 14) = !SPAREPARTID
                .MoveNext
            Next i
           
        End If
    End With
    
    If mainRs.State <> 0 Then
        Screen.MousePointer = vbDefault
        mainRs.Close
    End If
    'get the Total by PMATRIXID
    DoEvents
    Call modPMBranchIdentifying.sp_brnchIdentify_getDetails_total(MyBranchCode, MyID, myBoolean, mainRs)

        With mainRs
            cboBrand.Text = !PMBRAND
            cboModel.Text = !PMMODEL
            txtTOTALPN.Text = Format$(intTtlPN, 0)
            txtTotalSRQ.Text = Format$(intTOTALSRQ, 0)
            txtTotalAVQ.Text = Format$(intTotalAVQ, 0)
            If mySrc = "From Summary" Then
                txtTotalALQ.Text = Format$(intALQ, 0)
                txtRefNo.Text = Format$(!RefNo, "")
                txtDatePrepared.Text = Format$(!DATEPREPARED, "mm/dd/yyyy")
                txtDateFinal.Text = Format$(!DATEFINAL, "mm/dd/yyyy")
                txtProcsBy.Text = Format$(!PROCESSEDBY, "")
            End If
        End With

    mySrc = ""
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->branchIdentifyingShowRecord()")
End Sub


Private Sub txtALQ_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
            Tmp = Split(txtALQ.Tag, "|")
            
            For i = LBound(Tmp) To UBound(Tmp)
                flxSPARTS.TextMatrix(ActRow, i + 1) = Tmp(i)
            Next i
            
            txtALQ.Visible = False
            flxSPARTS.SetFocus
            flxSPARTS.Col = 11
    End If
End Sub
Public Sub ComputeTotalAllocation()
     
    intTotalALQ = 0
    With flxSPARTS
        For i = 2 To .Rows - 1
               intTotalALQ = intTotalALQ + Val(Replace(.TextMatrix(i, 11), ",", ""))
        Next i
    End With
    txtTotalALQ.Text = Format(intTotalALQ, "0")
    
End Sub


Private Sub txtALQ_KeyPress(KeyAscii As Integer)
On Error GoTo er

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtALQ, KeyAscii, NumbersOnly)

    
    If KeyAscii = 13 Then
        
        flxSPARTS.TextMatrix(ActRow, 11) = txtALQ.Text
        
        If Val(flxSPARTS.TextMatrix(ActRow, 11)) > Val(flxSPARTS.TextMatrix(ActRow, 10)) Then
            MsgBox "Allocated Quantity is Greater Than to Available Quantity", vbInformation + vbOKOnly, "Validation"
            flxSPARTS.TextMatrix(ActRow, 11) = 0
            txtALQ.Text = ""
            txtALQ.Visible = False
            flxSPARTS.SetFocus
            flxSPARTS.Row = ActRow
            flxSPARTS.Col = 11
            Exit Sub
        End If
        
        
        If txtALQ.Text <> "" Then
                
            For i = ActRow To flxSPARTS.Rows - 1
                If flxSPARTS.TextMatrix(i, 11) = "" Or Val(flxSPARTS.TextMatrix(i, 11)) = 0 Then
                    txtALQ.Visible = False
                    flxSPARTS.SetFocus
                    flxSPARTS.Row = i
                    flxSPARTS.Col = 11
                Exit For
                Else
                    txtALQ.Visible = False
                    flxSPARTS.SetFocus
                    flxSPARTS.Row = ActRow
                    flxSPARTS.Col = 11
                End If
                
            Next i
        
        Else
            flxSPARTS.SetFocus
            flxSPARTS.Col = 11
            txtALQ.Visible = False
        End If
        Call ComputeTotalAllocation
        
    End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtALQ-KeyPress()")
End Sub

Private Sub txtALQ_LostFocus()
On Error GoTo er

    If txtALQ.Tag = "skiplostfocus" Then
        txtALQ.Tag = ""
        Exit Sub
    End If
    
    If flxSPARTS.TextMatrix(ActRow, 11) = "" Or txtALQ.Text = "" Then
        MsgBox "Wrong Entry of Allocate Quantity!" & vbCrLf & "Please Enter a Correct Allocate Quantity.", vbExclamation + vbOKOnly, "Warning"
        txtALQ.Visible = False
        flxSPARTS.TextMatrix(ActRow, 11) = ""
        flxSPARTS.SetFocus
        flxSPARTS.Col = 11
        
        Exit Sub
    End If
    
    flxSPARTS.TextMatrix(ActRow, 11) = txtALQ.Text
    txtALQ.Visible = False
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtALQ-LostFocus()")
End Sub



Private Sub txtDateFinal_GotFocus()
    txtDateFinal.SelStart = 0
    txtDateFinal.SelLength = Len(txtDateFinal.Text)
End Sub

Private Sub txtDateFinal_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtDateFinal, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then
        flxSPARTS.SetFocus
        flxSPARTS.Row = 2
        flxSPARTS.Col = 11
    End If
End Sub

Private Sub txtDatePrepared_GotFocus()
    txtDatePrepared.SelStart = 0
    txtDatePrepared.SelLength = Len(txtDatePrepared.Text)
End Sub

Private Sub txtDatePrepared_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtDatePrepared, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then txtDateFinal.SetFocus
End Sub






Private Sub txtPROCSBY_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtProcsBy, KeyAscii, ValidChars)
End Sub

Private Sub txtrefno_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtRefNo, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtDatePrepared.SetFocus
End Sub


Function ValidationPartsMatrix() As String

    MyErr = ""
    If flxSPARTS.TextMatrix(flxSPARTS.Rows - 1, 1) = "" Then
        If flxSPARTS.Rows = 2 Then
            flxSPARTS.Rows = 1
            flxSPARTS.Rows = 2
        Else
            flxSPARTS.RemoveItem flxSPARTS.Rows - 1
        End If
    End If
        
'    For i = 1 To flxSPARTS.Rows - 1
'        If flxSPARTS.TextMatrix(i, 10) = "" Then MyErr = MyErr & vbCrLf & "There is no Allocated Quantity @ Row " & i & "!"
'    Next i


    ValidationPartsMatrix = MyErr

End Function

