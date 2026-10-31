VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form BranchAssign 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Branch Assignment"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   14385
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   14400
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   14400
      TabIndex        =   8
      Text            =   "HoldBackColor"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   8895
      Left            =   120
      TabIndex        =   4
      Tag             =   "XY"
      Top             =   480
      Width           =   14175
      Begin MSComctlLib.ProgressBar pBar 
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   8640
         Visible         =   0   'False
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxPMatrix 
         Height          =   8055
         Left            =   120
         TabIndex        =   7
         Tag             =   "XY"
         Top             =   720
         Width           =   13935
         _ExtentX        =   24580
         _ExtentY        =   14208
         _Version        =   393216
         Cols            =   11
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
      Begin VB.ComboBox cmboBranch 
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Tag             =   "XY"
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label lblBranch 
         Alignment       =   1  'Right Justify
         Caption         =   "SELECT BRANCH:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Tag             =   "XY"
         Top             =   240
         Width           =   1455
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Tag             =   "XY"
      Top             =   0
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   661
      Caption         =   "&Edit Branch Assignment"
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
      Image           =   "PMBranchAssign.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Tag             =   "XY"
      Top             =   0
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   661
      Caption         =   "E&xit Branch Assignment"
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
      Image           =   "PMBranchAssign.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
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
      Image           =   "PMBranchAssign.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   2520
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
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
      Image           =   "PMBranchAssign.frx":10CE
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   9240
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14475
      m_MDI_Height    =   9870
   End
End
Attribute VB_Name = "BranchAssign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ans As String
Dim Qry As String
Dim arrBRANCH() As String
Dim formAction As String
Dim BranchCode As String
Dim partsGroup As String
Dim i As Long, j As Long
Dim Tmp() As String
Dim tmpRs  As New ADODB.Recordset
Dim mainRs As New ADODB.Recordset
Dim strPartial As String, strTotal As String
Dim ActRow As Integer, ActCol As Integer
Dim StatSrch As Boolean
Dim ProperExit As Boolean

Sub CreateTempTable()
On Error GoTo er
        LYRcon.Execute "Drop Temporary Table IF  Exists sp_partsmatrix.TMP_BRANCH_ASSIGN;"

         DoEvents
       LYRcon.Execute " Create Temporary Table IF Not Exists sp_partsmatrix.TMP_BRANCH_ASSIGN( " & _
        "    BRANCHCODE Varchar(10)            " & _
        "   ,PMATRIXID Int(11)        " & _
        "   ,PARTSGROUP Varchar(20)      " & _
        ");"
       
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CreateTempTable()")
End Sub

Public Sub DeterminePartsgroup()

For i = 1 To flxPMatrix.Rows - 1
     If flxPMatrix.TextMatrix(i, 3) = "X" Then
            partsGroup = "COMPLETE"
            Qry = "Insert into sp_partsmatrix.TMP_BRANCH_ASSIGN(BRANCHCODE,PMATRIXID,PARTSGROUP) VALUES "
            Qry = Qry & "('" & BranchCode & "'," & Val(flxPMatrix.TextMatrix(i, 10)) & ",'" & partsGroup & "') "
            Call LYRcon.Execute(Qry)
     End If
     If flxPMatrix.TextMatrix(i, 5) = "X" Then
            partsGroup = "MAINTENANCE"
            Qry = "Insert into sp_partsmatrix.TMP_BRANCH_ASSIGN(BRANCHCODE,PMATRIXID,PARTSGROUP) VALUES "
            Qry = Qry & "('" & BranchCode & "'," & Val(flxPMatrix.TextMatrix(i, 10)) & ",'" & partsGroup & "') "
            Call LYRcon.Execute(Qry)
     End If
     If flxPMatrix.TextMatrix(i, 7) = "X" Then
            partsGroup = "PLASTIC"
            Qry = "Insert into sp_partsmatrix.TMP_BRANCH_ASSIGN(BRANCHCODE,PMATRIXID,PARTSGROUP) VALUES "
            Qry = Qry & "('" & BranchCode & "'," & Val(flxPMatrix.TextMatrix(i, 10)) & ",'" & partsGroup & "') "
            Call LYRcon.Execute(Qry)
     End If
     If flxPMatrix.TextMatrix(i, 9) = "X" Then
            partsGroup = "UNIVERSAL"
            Qry = "Insert into sp_partsmatrix.TMP_BRANCH_ASSIGN(BRANCHCODE,PMATRIXID,PARTSGROUP) VALUES "
            Qry = Qry & "('" & BranchCode & "'," & Val(flxPMatrix.TextMatrix(i, 10)) & ",'" & partsGroup & "') "
            Call LYRcon.Execute(Qry)
     End If

Next i
    
End Sub


Private Sub cmboBranch_Change()
On Error GoTo er
    If StatSrch Then
            StatSrch = False
            Exit Sub
    End If

    With cmboBranch
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmboBRANCH_Change()")
End Sub

Private Sub cmboBRANCH_Click()
  
    BranchCode = arrBRANCH(cmboBranch.ListIndex, 2)
    flxPMatrix.Rows = 1
    flxPMatrix.Rows = 2
 '************* CODE UPDATED: 04/11/17 ***************
    pBar.Visible = True
 '********* END CODE UPDATED: 04/11/17 ***************
    Call RefreshRecord
    Call ShowRecord(BranchCode)
End Sub

Private Sub cmboBranch_KeyDown(KeyCode As Integer, Shift As Integer)
'************** CODE UPDATED: 04/07/17 **************************
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
'********** END CODE UPDATED: 04/07/17 **************************
End Sub

Private Sub cmboBranch_KeyPress(KeyAscii As Integer)
'************** CODE UPDATED: 04/07/17 **************************
Dim cbolistIndex As Integer

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cmboBranch, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        cbolistIndex = modObjTools.getListIndex(cmboBranch, cmboBranch.Text)
        
        If cbolistIndex < 0 Then
            cmboBranch.SetFocus
        Else
            cmboBranch.ListIndex = cbolistIndex
        End If
    
        Call cmboBRANCH_Click
    End If
'********** END CODE UPDATED: 04/07/17 **************************
End Sub

Private Sub cmdTrans_Click(Index As Integer)
On Error GoTo er

    If Index = 0 Then 'edit
    
        formAction = "edit"
        Call clearform
        Call cmboDisplay
       ' Call RefreshRecord
        Call LockData(True)
        Call SetButtonVisibility(True)
        
        
    ElseIf Index = 1 Then 'exit
        formAction = "EXIT"
        ProperExit = True
        'Call clearform
        Unload Me
        
    ElseIf Index = 2 Then 'save
        ans = MsgBox("Do you want to save?", vbQuestion + vbYesNo, "Save")
        
        If ans = vbNo Then Exit Sub
        
        If formAction = "edit" Then
            Call CreateTempTable
            Call DeterminePartsgroup
            DoEvents
            Call modPMBranchAssign.sp_brnchAssign_create(BranchCode, mainRs)
        End If
        '****************** CODE UPDATED: 04/18/17 *************************
            MsgBox "Assigning Parts Matrix was successful!", vbInformation + vbOKOnly, "Save"
        '************** END CODE UPDATED: 04/18/17 *************************
            Call SetButtonVisibility(False)
            Call LockData(False)
            formAction = ""
    ElseIf Index = 3 Then 'cancel
        
        ans = MsgBox("Do you want to cancel?", vbQuestion + vbYesNo)
        
        If ans = vbNo Then Exit Sub
        
        Call SetButtonVisibility(False)
        Call LockData(False)
'        Call clearform
        formAction = ""
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdTrans-Click()")
End Sub

Private Sub flxPMatrix_Click()
On Error GoTo er

    ActRow = flxPMatrix.row
    ActCol = flxPMatrix.Col
    
    If formAction = "edit" Then
    
    If ActCol = 0 Or ActRow = 0 Then Exit Sub
    
    If cmboBranch <> "" Then
        Select Case ActCol
    
        Case 3
            flxPMatrix.TextMatrix(flxPMatrix.row, 3) = "X"
            flxPMatrix.CellFontBold = True
        Case 5
            flxPMatrix.TextMatrix(flxPMatrix.row, 5) = "X"
            flxPMatrix.CellFontBold = True
        Case 7
            flxPMatrix.TextMatrix(flxPMatrix.row, 7) = "X"
            flxPMatrix.CellFontBold = True

        Case 9
            flxPMatrix.TextMatrix(flxPMatrix.row, 9) = "X"
            flxPMatrix.CellFontBold = True
    
        End Select
    Else
        MsgBox "Please Select a Branch", vbInformation + vbOKOnly, "Branch Assignment"
    End If
    
    End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->flxPMatrix-click()")
End Sub

Private Sub flxPMatrix_DblClick()
    If formAction = "edit" Then
        If flxPMatrix.Col = 3 Then 'complete
            flxPMatrix.TextMatrix(flxPMatrix.row, 3) = ""
            flxPMatrix.CellFontBold = False
        ElseIf flxPMatrix.Col = 5 Then 'maintenance
            flxPMatrix.TextMatrix(flxPMatrix.row, 5) = ""
            flxPMatrix.CellFontBold = False
        ElseIf flxPMatrix.Col = 7 Then 'plastic
            flxPMatrix.TextMatrix(flxPMatrix.row, 7) = ""
            flxPMatrix.CellFontBold = False
        ElseIf flxPMatrix.Col = 9 Then 'universal
            flxPMatrix.TextMatrix(flxPMatrix.row, 9) = ""
            flxPMatrix.CellFontBold = False
        End If
    End If
End Sub

Private Sub Form_Load()
    Call LockData(False)
    Call flxColChanger
    '===================BUTTON POSITION===================
    
    If cmdTrans(0).Visible = False Then
            'Align Save Button
            cmdTrans(2).Left = cmdTrans(0).Left
            cmdTrans(2).Top = cmdTrans(0).Top
            'Align Cancel Button
            cmdTrans(3).Left = cmdTrans(1).Left
            cmdTrans(3).Top = cmdTrans(1).Top
     End If
End Sub

Public Sub flxColChanger()

With flxPMatrix
    'MERGE COL
    .TextMatrix(0, 0) = " "
    .TextMatrix(1, 0) = " "
    'MERGE ROW
    .TextMatrix(0, 2) = "COMPLETE"
    .TextMatrix(0, 3) = "COMPLETE"
    .TextMatrix(0, 4) = "MAINTENANCE"
    .TextMatrix(0, 5) = "MAINTENANCE"
    .TextMatrix(0, 6) = "PLASTIC"
    .TextMatrix(0, 7) = "PLASTIC"
    .TextMatrix(0, 8) = "UNIVERSAL"
    .TextMatrix(0, 9) = "UNIVERSAL"
    
    'MERGE COL =FALSE
    .TextMatrix(0, 1) = "PARTS MATRIX"

    'SET MERGING
    .MergeCells = 1
    .MergeRow(0) = True
    .MergeCol(0) = True
    .MergeCol(0) = True
    For i = 1 To 2
        .MergeCol(i) = False
    Next i
    'SET THE COLUMN COLOR GREEN
    .row = 0
    .Col = 0
    .Col = 1
End With
    
    Call flxWidth
    
End Sub
Public Sub flxWidth()

    With flxPMatrix
        .Cols = 11
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignCenterCenter
        .ColAlignment(7) = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignCenterCenter
        .ColAlignment(9) = flexAlignCenterCenter
       ' .FormatString = "|<PARTS MATRIX |^COMPLETE |^MAINTENANCE |^PLASTIC |^UNIVERSAL"
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.26  'PARTS MATRIX
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.09 'COMPLETE
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.09 'COMPLETE
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.09 'MAINTENANCE
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.09 'MAINTENANCE
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.09 'PLASTIC
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.09 'PLASTIC
        .ColWidth(8) = (.Width - .ColWidth(0)) * 0.09 'UNIVERSAL
        .ColWidth(9) = (.Width - .ColWidth(0)) * 0.09 'UNIVERSAL
        .ColWidth(10) = 0 '(.Width - .ColWidth(0)) * 0.06 'ID
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo er

    If ProperExit Then
        If mainRs.State = adStateOpen Then mainRs.Close
        Call cmdTrans_Click(1)
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
'************* CODE UPDATED: 3/14/17 *****************
    '
    'Frame1.Width = Me.Width - 285
    '
    'lblBranch.Left = 120
    'lblBranch.Width = 1455
    '
    'cmboBranch.Left = 1680
    'cmboBranch.Width = 4095
    '
    'flxPMatrix.Left = 120
    'flxPMatrix.Width = Frame1.Width - 240
'********* END CODE UPDATED: 3/14/17 *****************

End Sub

Public Sub cmboDisplay()
On Error GoTo er

   ' If formAction = "add" Then
        'Get Data from database
        DoEvents
        Call modPMBranchAssign.sp_get_Branch(tmpRs)
        If tmpRs.State = 1 Then
           
            Call modObjTools.fillArray(modGlobalVar.AllArrBranch, tmpRs)
          
        End If
        If tmpRs.State <> 0 Then tmpRs.Close
        
        'Load Data to combo box
        Call SetCboListData
        
  '  End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub

Sub SetCboListData()
On Error GoTo er

    If cmboBranch.Text = "" Then
        DoEvents
        Call modObjTools.addToListObjectUsingArrayWithCols(cmboBranch, arrBRANCH, modGlobalVar.AllArrBranch, 2, 1)
       End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->SetCboListData()")
End Sub

Sub LockData(DataEntry As Boolean)
    DataEntry = Not DataEntry

    cmboBranch.Locked = DataEntry
   
    If cmboBranch.Locked Then
        cmboBranch.BackColor = txtHoldGrayBackColor.BackColor
    Else
        cmboBranch.BackColor = txtHoldWhiteBackColor.BackColor

        
    End If
End Sub

Private Sub SetButtonVisibility(DataEntry As Boolean)
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->SAVE
'3 ->CANCEL


    'hide all button
    For i = 0 To 2
        cmdTrans(i).Visible = False
    Next i

    If DataEntry Then
        cmdTrans(2).Visible = True
        cmdTrans(3).Visible = True
    Else
    For i = 0 To 2
        cmdTrans(i).Visible = True
        cmdTrans(2).Visible = False
        cmdTrans(3).Visible = False
    Next i
    
    
    End If
End Sub

Sub clearform()

    flxPMatrix.Rows = 1
    flxPMatrix.Rows = 2
    cmboBranch.Text = ""
    
End Sub
Public Sub RefreshRecord()

DoEvents
Call modPMBranchAssign.sp_brnchAssign_getDetails(mainRs)
'************* CODE UPDATED: 3/14/17 *****************
    pBar.Visible = True
    pBar.Value = 0
'********* END CODE UPDATED: 3/14/17 *****************
    If mainRs.State <> 1 Then Exit Sub
    
    If mainRs.RecordCount <= 0 Then
'************* CODE UPDATED: 3/14/17 *****************
        pBar.Visible = False
        pBar.Value = 0
'********* END CODE UPDATED: 3/14/17 *****************
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
     With mainRs
        If .RecordCount > 0 Then
            .MoveFirst
            flxPMatrix.Rows = .RecordCount + 1
        '************* CODE UPDATED: 3/14/17 *****************
            pBar.Max = .RecordCount
        '********* END CODE UPDATED: 3/14/17 *****************
            For i = 1 To flxPMatrix.Rows - 1
                flxPMatrix.TextMatrix(i, 1) = !PARTSMATRIX
                flxPMatrix.TextMatrix(i, 2) = !COMPLETE
                flxPMatrix.TextMatrix(i, 4) = !MAINTENANCE
                flxPMatrix.TextMatrix(i, 6) = !PLASTIC
                flxPMatrix.TextMatrix(i, 8) = !UNIVERSAL
                flxPMatrix.TextMatrix(i, 10) = !PMATRIXID
                .MoveNext
        '************* CODE UPDATED: 3/14/17 *****************
                pBar.Value = pBar.Value + 1
        '********* END CODE UPDATED: 3/14/17 *****************
            Next i
        End If
    End With
    
     If mainRs.State <> 0 Then
        Screen.MousePointer = vbDefault
        mainRs.Close
    End If

End Sub

Public Sub ShowRecord(myBranchCode As String)

DoEvents
Call modPMBranchAssign.sp_brnchAssign_getDetails_wBCODE(myBranchCode, mainRs)

If mainRs.State <> 1 Then Exit Sub
    
    If mainRs.RecordCount <= 0 Then
    '************* CODE UPDATED: 3/14/17 *****************
        pBar.Visible = False
        pBar.Value = 0
    '********* END CODE UPDATED: 3/14/17 *****************
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
     With mainRs
        If .RecordCount > 0 Then
            .MoveFirst
          Do While Not .EOF
            If mainRs!partsGroup = "COMPLETE" Then
                For i = 1 To flxPMatrix.Rows - 1
                    If flxPMatrix.TextMatrix(i, 10) = !PMATRIXID Then
                        flxPMatrix.TextMatrix(i, 3) = "X"
                        flxPMatrix.row = i
                        flxPMatrix.Col = 3
                        flxPMatrix.CellFontBold = True
                    End If
                Next i
            ElseIf mainRs!partsGroup = "MAINTENANCE" Then
                For i = 1 To flxPMatrix.Rows - 1
                    If flxPMatrix.TextMatrix(i, 10) = !PMATRIXID Then
                        flxPMatrix.TextMatrix(i, 5) = "X"
                        flxPMatrix.row = i
                        flxPMatrix.Col = 5
                        flxPMatrix.CellFontBold = True
                    End If
                Next i
            ElseIf mainRs!partsGroup = "PLASTIC" Then
                For i = 1 To flxPMatrix.Rows - 1
                    If flxPMatrix.TextMatrix(i, 10) = !PMATRIXID Then
                        flxPMatrix.TextMatrix(i, 7) = "X"
                        flxPMatrix.row = i
                        flxPMatrix.Col = 7
                        flxPMatrix.CellFontBold = True
                    End If
                Next i
            ElseIf mainRs!partsGroup = "UNIVERSAL" Then
                For i = 1 To flxPMatrix.Rows - 1
                    If flxPMatrix.TextMatrix(i, 10) = !PMATRIXID Then
                        flxPMatrix.TextMatrix(i, 9) = "X"
                        flxPMatrix.row = i
                        flxPMatrix.Col = 9
                        flxPMatrix.CellFontBold = True
                    End If
                Next i
            End If
            .MoveNext
            Loop
        End If
    End With
    
     If mainRs.State <> 0 Then
        Screen.MousePointer = vbDefault
        mainRs.Close
    End If
    
'************* CODE UPDATED: 3/14/17 *****************
        pBar.Visible = False
        pBar.Value = 0
'********* END CODE UPDATED: 3/14/17 *****************

End Sub
