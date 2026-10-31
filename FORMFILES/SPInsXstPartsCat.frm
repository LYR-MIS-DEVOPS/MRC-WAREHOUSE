VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form SPInsXstPartsCat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Existing Spare Parts"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13230
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   13230
   StartUpPosition =   2  'CenterScreen
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
      Left            =   10560
      TabIndex        =   5
      Top             =   240
      Width           =   2535
   End
   Begin VB.VScrollBar UpDRef 
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   240
      Width           =   255
   End
   Begin MSComctlLib.ProgressBar Pbar 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   6240
      Visible         =   0   'False
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
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
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
   Begin MSFlexGridLib.MSFlexGrid flxSP 
      Height          =   5505
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   12960
      _ExtentX        =   22860
      _ExtentY        =   9710
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   16777215
      FocusRect       =   2
      ScrollBars      =   2
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "|<PARTNUMBER |<DESCRIPTION |<APPLIC MODEL |<CLASSIFICATION"
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
   Begin lvButton.lvButtons_H cmdFind 
      Height          =   495
      Left            =   7200
      TabIndex        =   8
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
      Image           =   "SPInsXstPartsCat.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdInsert 
      Height          =   375
      Left            =   9000
      TabIndex        =   9
      Top             =   6240
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
      Image           =   "SPInsXstPartsCat.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   375
      Left            =   10380
      TabIndex        =   10
      Top             =   6240
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
      Image           =   "SPInsXstPartsCat.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   11760
      TabIndex        =   11
      Top             =   6240
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
      Image           =   "SPInsXstPartsCat.frx":10CE
      cBack           =   -2147483633
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
      Left            =   6720
      TabIndex        =   7
      Top             =   6240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblClass 
      Caption         =   "CLASSIFICATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9120
      TabIndex        =   6
      Top             =   315
      Width           =   1455
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   8640
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   13320
      m_MDI_Height    =   7170
   End
   Begin VB.Label lblRefSearch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FIND PARTS DESCRIPTION"
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
      Left            =   120
      TabIndex        =   1
      Top             =   315
      Width           =   2100
   End
   Begin VB.Menu mnuSP 
      Caption         =   "SP"
      Visible         =   0   'False
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
End
Attribute VB_Name = "SPInsXstPartsCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsMain As New ADODB.Recordset
Dim ShowOpt As Boolean, SuppSrch As Boolean
Dim ActRow As Integer, ActCol As Integer
Dim arrClass() As String
Dim i As Long, j As Long
Public MySource As String

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

Public Sub CboDisplay()
On Error GoTo er
    'Get Data from database
    DoEvents
    Qry = "Call partscatalog.sp_getClassification()"
    
    With rsMain
        .Open Qry, conPartsCat, adOpenStatic, adLockReadOnly
        If .State = 1 Then
           Call modObjTools.addToListObjectWithArrayGivenField(cboClass, rsMain, arrClass, "", "REMARKS")
        End If
        If .State <> 0 Then rsMain.Close
    End With

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub

Private Sub cboClass_Change()
On Error GoTo er
    If StatSrch Then
            StatSrch = False
            Exit Sub
    End If
    With cboClass
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboClass-Change()")
End Sub


Private Sub cboClass_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboClass_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboClass, KeyAscii, ValidChars)
    If KeyAscii = 13 Then flxSP.SetFocus
End Sub

Private Sub cboFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboFind, KeyAscii, ValidChars)
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim rsSP As New ADODB.Recordset
    Dim Qry As String
    Dim intFindVal As Integer
    
    intFindVal = 0
    intFindVal = Val(UpDRef.Value)
    Qry = "call partscatalog.sp_catalog_search_item('" & cboFind.Text & "','" & cboClass.Text & "'," & intFindVal & ")"

    If rsSP.State = adStateOpen Then rsSP.Close
    rsSP.Open Qry, conPartsCat, adOpenKeyset, adLockOptimistic
    With rsSP
        If .RecordCount <= 0 Then
            MsgBox "No Record Found!", vbInformation + vbOKOnly, "Warning"
            .Close
            Exit Sub
        End If
        
        If .RecordCount > 0 Then
            .MoveFirst
            flxSP.Rows = .RecordCount + 1
            For i = 1 To flxSP.Rows - 1
                flxSP.TextMatrix(i, 1) = !PartNumber
                flxSP.TextMatrix(i, 2) = !Description
                flxSP.TextMatrix(i, 3) = !ApplicModel
                flxSP.TextMatrix(i, 4) = !Remarks 'CLASSIFICATION
                flxSP.TextMatrix(i, 5) = Format(!STOREPRICE, "##,##0.00")
                flxSP.TextMatrix(i, 6) = !SPAREPARTID
                .MoveNext
            Next i
        End If
    End With
    If rsSP.State <> 0 Then rsSP.Close
 End Sub

Private Sub cmdInsert_Click()
    Dim i As Integer
    
    
    If MySource = "Item Entry" Then
        With frmDataEntry.flxParts
            If .TextMatrix(.Rows - 1, 1) = "" Then 'Remove last row if empty
                If .Rows = 2 Then
                    .Rows = 1
                Else
                    .RemoveItem .Rows - 1
                End If
            End If
            For i = 0 To flxSP.Rows - 1
                If flxSP.TextMatrix(i, 0) <> "" Then
                    If Not SPinFlx(flxSP.TextMatrix(i, 1)) Then
                        .AddItem .Rows
                        .TextMatrix(.Rows - 1, 1) = flxSP.TextMatrix(i, 1)  'PARTNUMBER
                        .TextMatrix(.Rows - 1, 2) = flxSP.TextMatrix(i, 2)  'DESCRIPTION
                        .TextMatrix(.Rows - 1, 4) = Format(flxSP.TextMatrix(i, 5), "##,##0.00") 'STOREPRICE
                        .TextMatrix(.Rows - 1, 6) = flxSP.TextMatrix(i, 6)   'SPID
                    End If
                End If
            Next i
            If .Rows = 1 Then .AddItem "1"
        End With
    End If
End Sub

Function SPinFlx(myPN As String) As Boolean
    Dim i As Integer
    
    If MySource = "Item Entry" Then
        With frmDataEntry.flxParts
            For i = 0 To .Rows - 1
                'If myPN = .TextMatrix(i, 1) Then
                
                If myPN = .TextMatrix(i, 1) Then
                    SPinFlx = True
                    Exit Function
                End If
            Next i
        End With
        SPinFlx = False
        Exit Function
    End If
End Function

Private Sub cmdOK_Click()
    Call cmdInsert_Click
    Unload Me
End Sub

Private Sub flxSP_DblClick()
    If flxSP.TextMatrix(flxSP.row, 0) = "" Or Val(flxSP.TextMatrix(flxSP.row, 1)) > 0 Then
        Call mnuMark_Click
    Else
        Call mnuUnmark_Click
    End If
End Sub

Private Sub flxSP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If flxSP.TextMatrix(flxSP.row, 1) = "" And Val(flxSP.TextMatrix(flxSP.row, 1)) > 0 Then
            Call mnuMark_Click
        Else
            Call mnuUnmark_Click
        End If
    End If
End Sub

Private Sub flxSP_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If flxSP.TextMatrix(flxSP.row, 1) = "" Then Exit Sub
    If Button = vbRightButton Then
        Call PopupMenu(mnuSP)
    End If
End Sub

Private Sub Form_Load()
    Call CboDisplay
    UpDRef.Value = 1
    ActRow = 1
    Call flxChnger
    If frmDataEntry.LockFlx = True Then
        cmdInsert.Enabled = False
        cmdOK.Enabled = False
    End If
End Sub

Public Sub flxChnger()

    With flxSP
        '.Clear
        .Cols = 7
        .ColWidth(0) = 400
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.16  'PARTS NUMBER
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.45 'DESCRIPTION
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.19 'APPLIC MODEL
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.17 'CLASSIFICATION
        .ColWidth(5) = 0 '(.Width - .ColWidth(0)) * 0.07 'STOREPRICE
        .ColWidth(6) = 0 '(.Width - .ColWidth(0)) * 0.07 'SPAREPARTID
    End With
    
End Sub

Private Sub UpDRef_Change()
    UpDRef.Max = 3
    If UpDRef.Value = 0 Then
        lblRefSearch.Caption = "FIND PARTNUMBER"
    ElseIf UpDRef.Value = 1 Then
        lblRefSearch.Caption = "FIND PARTS DESCRIPTION"
    ElseIf UpDRef.Value = 2 Then
        lblRefSearch.Caption = "FIND PARTS APPLCMODEL"
    ElseIf UpDRef.Value = 3 Then
        UpDRef.Value = 0
    End If
   
End Sub

Private Sub ClearFRow()
Dim cRw As Integer
    For cRw = 0 To 18
        flxSP.TextMatrix(0, cRw) = ""
    Next cRw
End Sub
