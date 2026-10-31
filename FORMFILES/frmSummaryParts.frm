VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSummaryParts 
   Caption         =   "Form Summary"
   ClientHeight    =   8625
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13905
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   13905
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   120
      TabIndex        =   9
      Tag             =   "XYF"
      Top             =   8280
      Visible         =   0   'False
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Tag             =   "XY"
      Top             =   480
      Width           =   13695
      Begin VB.ComboBox cboGROUP 
         Height          =   315
         Left            =   5760
         TabIndex        =   17
         Top             =   600
         Width           =   3495
      End
      Begin VB.ComboBox cboVERSION 
         Height          =   315
         Left            =   5760
         TabIndex        =   16
         Top             =   240
         Width           =   3495
      End
      Begin VB.ComboBox cboMODEL 
         Height          =   315
         Left            =   1200
         TabIndex        =   13
         Top             =   600
         Width           =   3615
      End
      Begin VB.ComboBox cboBRAND 
         Height          =   315
         Left            =   1200
         TabIndex        =   12
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   11640
         TabIndex        =   4
         Tag             =   "XYF"
         Top             =   195
         Width           =   1935
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   10320
         TabIndex        =   5
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
         Image           =   "frmSummaryParts.frx":0000
         cBack           =   -2147483633
      End
      Begin VB.Label Label4 
         Caption         =   "GROUP:"
         Height          =   255
         Left            =   4920
         TabIndex        =   15
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "VERSION:"
         Height          =   255
         Left            =   4920
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "MODEL:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "BRAND:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
   End
   Begin lvButton.lvButtons_H cmdShow 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "Show &Form"
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
      Image           =   "frmSummaryParts.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Print Summary"
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
      Image           =   "frmSummaryParts.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
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
      Image           =   "frmSummaryParts.frx":10CE
      cBack           =   -2147483633
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSum 
      Height          =   6975
      Left            =   120
      TabIndex        =   8
      Tag             =   "XYF"
      Top             =   1680
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   12303
      _Version        =   393216
      RowHeightMin    =   285
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   7800
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14145
      m_MDI_Height    =   9195
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Form SUMMARY"
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
      Height          =   375
      Index           =   2
      Left            =   9120
      TabIndex        =   6
      Top             =   120
      Width           =   4500
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Form SUMMARY"
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
      Height          =   375
      Index           =   3
      Left            =   9135
      TabIndex        =   7
      Top             =   150
      Width           =   4500
   End
   Begin VB.Menu mnuSum 
      Caption         =   "Summary"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "FrmSummaryParts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSummary As New ADODB.Recordset
Dim rsMain As New ADODB.Recordset
Dim rsParts As New ADODB.Recordset
Dim ActRow As Integer, fRow As Integer, fCol As Integer, TCol As Integer, dCodePrnt As Integer
Dim catID As Integer
Dim Qry As String
Dim arrGROUP() As String
Dim arrBRAND() As String
Dim arrBLOCK() As String
Dim arrMODEL() As String
Dim arrVERSION() As String
Public mySrc As String
Dim i As Long, j As Long
Dim strPartial As String, strTotal As String, brandName As String
Dim StatSrch As Boolean, ProperExit As Boolean, PrtSrch As Boolean, PrtBln As Boolean

Public Sub CboDisplay()
On Error GoTo er
    'If formAction = "add" Then
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
    'End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub

Public Sub CboDisplayModel()
On Error GoTo er
    'If formAction = "add" Then
        'Get Data from database
        DoEvents
        Qry = "Call partscatalog.sp_mc_getModel('" & cboBRAND.Text & "')"
        
        With rsMain
            .Open Qry, conPartsCat, adOpenStatic, adLockReadOnly
            If .State = 1 Then
               Call modObjTools.addToListObjectWithArrayGivenField(cboMODEL, rsMain, arrMODEL, "cat_ID", "cat_Model")
            End If
            If .State <> 0 Then rsMain.Close
        End With
    'End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplayModel()")
End Sub

Public Sub CboDisplayVersion()
On Error GoTo er
    'If formAction = "add" Then
        'Get Data from database
        DoEvents
        Qry = "Call partscatalog.sp_mc_getVersion('" & cboMODEL.Text & "')"
        
        With rsMain
            .Open Qry, conPartsCat, adOpenStatic, adLockReadOnly
            If .State = 1 Then
               Call modObjTools.addToListObjectWithArrayGivenField(cboVERSION, rsMain, arrVERSION, "cat_ID", "cat_Version")
            End If
            If .State <> 0 Then rsMain.Close
        End With
    'End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplayVersion()")
End Sub

Public Sub CboDisplayGroup()
On Error GoTo er
    'If formAction = "add" Then
        'Get Data from database
        DoEvents
        Qry = "Call partscatalog.sp_pc_groupList_summ()"
        
        With rsMain
            .Open Qry, conPartsCat, adOpenStatic, adLockReadOnly
            If .State = 1 Then
               Call modObjTools.addToListObjectWithArrayGivenField(cboGROUP, rsMain, arrGROUP, "ID", "GROUPNAME")
            End If
            If .State <> 0 Then rsMain.Close
        End With
    'End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplayGroup()")
End Sub

Public Sub flxChnger()
    lblHeading(2).Left = Frame1.Left + (Frame1.Width - lblHeading(2).Width) - 30
    lblHeading(3).Left = Frame1.Left + (Frame1.Width - lblHeading(3).Width)
    With flxSum
        '.Clear
        .Cols = 11
        .FormatString = "|<DATE |<BRAND |<MODEL |<VERSION |<GROUP |<BLOCK |<TOTAL PN |<LASTUPDATE"
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.1  'DATE
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.1 'BRAND
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.13 'MODEL
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.13 'VERSION
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.1 'GROUP
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.13 'BLOCK
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.1 'TOTAL PN
        .ColWidth(8) = (.Width - .ColWidth(0)) * 0.15 'LASTUPDATE
        .ColWidth(9) = 0 '(.Width - .ColWidth(0)) * 0.1 'CATALOG ID
        .ColWidth(10) = 0 '(.Width - .ColWidth(0)) * 0.1 'CATEGORY ID
    End With
    
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
    Call CboDisplayModel
    Call CboDisplayGroup
    cboMODEL.SetFocus
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
If cboBRAND.Text <> "" And cboMODEL <> "" And cboVERSION <> "" Then
    Call RefreshRecord
    'cboBlock.SetFocus
Else
    If cboBRAND.Text = "" Then
        MsgBox "BRAND IS REQUIRED!", vbExclamation, "Validation"
        cboBRAND.SetFocus
        Exit Sub
    ElseIf cboMODEL.Text = "" Then
        MsgBox "MODEL IS REQUIRED!", vbExclamation, "Validation"
        cboMODEL.SetFocus
        Exit Sub
    ElseIf cboVERSION.Text = "" Then
        MsgBox "VERSION IS REQUIRED!", vbExclamation, "Validation"
        cboVERSION.SetFocus
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

Private Sub cboMODEL_Change()
On Error GoTo er
    If StatSrch Then
            StatSrch = False
            Exit Sub
    End If
    With cboMODEL
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboMODEL-Change()")
End Sub

Private Sub cboMODEL_Click()
If cboBRAND.Text <> "" Then
    Call CboDisplayVersion
    cboVERSION.SetFocus
Else
   MsgBox "BRAND IS REQUIRED!", vbExclamation, "Validation"
   cboBRAND.SetFocus
   Exit Sub
End If
End Sub

Private Sub cboMODEL_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(txtBrand, KeyAscii, ValidChars)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = 13 Then Call cboMODEL_Click
End Sub

Private Sub cboVERSION_Change()
On Error GoTo er
    If StatSrch Then
            StatSrch = False
            Exit Sub
    End If
    catID = 0
    With cboVERSION
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboVERSION-Change()")
End Sub

Private Sub cboVERSION_Click()
If cboMODEL.Text <> "" Then
    cboGROUP.SetFocus
    catID = arrVERSION(cboVERSION.ListIndex, 0)
Else
    MsgBox "MODEL IS REQUIRED!", vbExclamation, "Validation"
    cboMODEL.SetFocus
    Exit Sub
End If
End Sub

Private Sub cboVERSION_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(txtBrand, KeyAscii, ValidChars)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = 13 Then Call cboVERSION_Click
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxSum
        For i = fRow To .Rows - 1
            For j = fCol To TCol - 1
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .row = i
                    .ColSel = TCol - 1: .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = TCol - 1, i + 1, i)
                    fCol = IIf(j = TCol - 1, 1, j + 1)
                    ActRow = i
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

'Private Sub cmdNext_Click()
'    cmdNext.Tag = "2"
'    cmdNext.Visible = False
'    cmdPrev.Left = cmdNext.Left
'    cmdPrev.Top = cmdNext.Top
'    cmdPrev.Visible = True
'    dTo.Visible = True
'    cmdRefresh.Visible = True
'    dTo.Value = dFrom.Value
'End Sub

'Private Sub cmdPrev_Click()
'    cmdNext.Tag = "1"
'    cmdNext.Visible = True
'    cmdPrev.Visible = False
'    dTo.Visible = False
'    cmdRefresh.Visible = False
'    Call RefreshRecord(1)
'End Sub

Private Sub cmdPrint_Click()
    If mySrc = "Item Entry" Then
        If rsSummary.State = adStateOpen Then rsSummary.Close
        Qry = "call partscatalog.sp_partscatalog_summary('" & Val(catID) & "','" & Trim(cboGROUP.Text) & "')"
        
        frmPrint.QryPrn = Qry
            
        'frmPrint.ParamValue1 = Format(dFrom.Value, "YYYY-mm-dd") & "TO" & Format(dTo.Value, "YYYY-mm-dd")
        frmPrint.MySource = "Parts Catalog Summary"
        frmPrint.Show
    Else
        MsgBox "Not yet available", vbInformation + vbOKOnly, "Print"
    End If
End Sub

Private Sub cmdShow_Click()
With flxSum
        If .Rows = 2 And .TextMatrix(1, 1) = "" Then
            MsgBox "There is no record to view!", vbInformation + vbOKOnly, "Confirm"
            Exit Sub
        End If
        If mySrc = "Item Entry" Then
            Call frmDataEntry.ShowRecord(Val(.TextMatrix(ActRow, 9)), Val(.TextMatrix(ActRow, 10)))
            Unload Me
        End If
    End With

End Sub

'Private Sub dFrom_CloseUp()
'    FrDate = dFrom.Value
'    If cmdNext.Tag = "1" Then
'        Call RefreshRecord(1)
'    End If
'End Sub

'Private Sub dTo_CloseUp()
'    ToDate = dTo.Value
'    If dTo.Value < dFrom.Value Then
'        MsgBox "Please choose a date not less than " & dFrom.Value, vbExclamation + vbOKOnly, "Validation Error"
'        dTo.Value = dFrom.Value
'    End If
'End Sub

Private Sub flxSum_EnterCell()
ActRow = flxSum.row
End Sub

Private Sub Form_Load()
    Dim Qry As String, ShowDateRange As Boolean, i As Integer
    'Call ConnectPartsCatalog
    'dFrom.Value = Date
    'cmdNext.Tag = "1"
    'identify total columns of trancsaction and Access ID to identify user accessibility
    If mySrc = "Item Entry" Then
        TCol = 11
        Call CboDisplay
        'Call cmdNext_Click
        'dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        'dTo.Value = Date
        Call RefreshRecord
    End If
    lblHeading(2).Caption = mySrc & " Summary"
    lblHeading(3).Caption = mySrc & " Summary"
    Me.Caption = mySrc & " Summary"
    ActRow = 1
End Sub

Public Sub RefreshRecord()
    If mySrc = "Item Entry" Then
        Qry = "call partscatalog.sp_partscatalog_summary('" & Val(catID) & "','" & Trim(cboGROUP.Text) & "')"
    End If
    'dCodePrnt = dCode
    If rsSummary.State = adStateOpen Then rsSummary.Close
    rsSummary.Open Qry, conPartsCat, adOpenKeyset, adLockOptimistic
    If rsSummary.RecordCount = 0 Then
        flxSum.Clear
        flxSum.Rows = 2
        flxSum.Cols = TCol
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxSum.DataSource = rsSummary
    ActRow = 1
    flxSum.Visible = True
    Call RXResizer1_ResizeComplete
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If rsSummary.State = adStateOpen Then rsSummary.Close
End Sub

Private Sub Form_Resize()
 Call RXResizer1_ResizeComplete
End Sub

Private Sub RXResizer1_ResizeComplete()
Call flxChnger
End Sub

Private Sub txtFind_Change()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

