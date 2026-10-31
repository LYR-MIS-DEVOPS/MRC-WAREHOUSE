VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form BrnBBOparts 
   Caption         =   "Branch Back Order Part Number"
   ClientHeight    =   7500
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   15210
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOrder 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   2400
      TabIndex        =   10
      Top             =   7440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   240
      TabIndex        =   9
      Tag             =   "xy"
      Top             =   6840
      Visible         =   0   'False
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox txtFind 
      Height          =   315
      Left            =   12240
      TabIndex        =   4
      Tag             =   "XYF"
      Top             =   375
      Width           =   2535
   End
   Begin MSComCtl2.UpDown updFilter 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Tag             =   "XYF"
      Top             =   360
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      Value           =   1
      Max             =   7
      Min             =   1
      Enabled         =   -1  'True
   End
   Begin VB.ComboBox cboFilter 
      Height          =   315
      Left            =   2520
      TabIndex        =   2
      Tag             =   "xyf"
      Top             =   360
      Width           =   3615
   End
   Begin MSFlexGridLib.MSFlexGrid flxBBO 
      Height          =   6015
      Left            =   240
      TabIndex        =   0
      Tag             =   "xyf"
      Top             =   840
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   10610
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   -2147483633
      FormatString    =   "     |<PART NUMBER |<DESCRIPTION       |^PO REF.|^PO DATE|<BRANCH|^BO QTY.  |^ORDER QTY.|<SUPPLIER  "
   End
   Begin lvButton.lvButtons_H cmdFind 
      Height          =   375
      Left            =   10920
      TabIndex        =   5
      Tag             =   "XY"
      Top             =   330
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Find Below"
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
      Image           =   "BrnBBOparts.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdInsert 
      Height          =   375
      Left            =   10680
      TabIndex        =   6
      Tag             =   "xyf"
      Top             =   6960
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
      Image           =   "BrnBBOparts.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   375
      Left            =   12060
      TabIndex        =   7
      Tag             =   "xyf"
      Top             =   6960
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
      Image           =   "BrnBBOparts.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   13440
      TabIndex        =   8
      Tag             =   "xyf"
      Top             =   6960
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
      Image           =   "BrnBBOparts.frx":10CE
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
      Left            =   0
      TabIndex        =   11
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblFilter 
      Caption         =   "SHOW ALL BBO PARTS:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Tag             =   "xyf"
      Top             =   420
      Width           =   2415
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   7920
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   15450
      m_MDI_Height    =   8070
   End
End
Attribute VB_Name = "BrnBBOparts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ActRow As Integer, ActCol As Integer, fRow As Integer, fCol As Integer


Private Sub DisplayBBO()
    Dim rsBBO As New ADODB.Recordset
    Dim Qry As String, myW1 As String
    Dim MySP() As String
    Dim i As Integer
    
    If updFilter.Value = 1 Then 'All Back Order Parts
        Qry = "SELECT spareparts.PARTNUMBER, spareparts.DESCRIPTION, backorder.REFERENCE, backorder.TRANSDATE, backorder.BRANCHNAME, backorder.BOBALANCE, supplier.SUPPLIERNAME, spareparts.SPAREPARTID, spareparts.COST, backorder.BACKORDERID " & _
              "From backorder INNER JOIN spareparts ON (backorder.SPAREPARTID = spareparts.SPAREPARTID) INNER JOIN supplier ON (spareparts.SUPPLIERCODE = supplier.SUPPLIERCODE) " & _
              "Order By backorder.TRANSDATE, supplier.SUPPLIERNAME, spareparts.PARTNUMBER;"
    
    ElseIf updFilter.Value = 2 Then 'part number
        MySP = Split(cboFilter.Text, " ")
        myW1 = ""
        For i = LBound(MySP) To UBound(MySP)
            If myW1 = "" Then 'First Criteria
                myW1 = "SPAREPARTS.PARTNUMBER Like '%" & MySP(i) & "%'"
            Else
                myW1 = myW1 & " AND SPAREPARTS.PARTNUMBER Like '%" & MySP(i) & "%'"
            End If
        Next i
        Qry = "SELECT spareparts.PARTNUMBER, spareparts.DESCRIPTION, backorder.REFERENCE, backorder.TRANSDATE, backorder.BRANCHNAME, backorder.BOBALANCE, supplier.SUPPLIERNAME, spareparts.SPAREPARTID, spareparts.COST, backorder.BACKORDERID " & _
              "From backorder INNER JOIN spareparts ON (backorder.SPAREPARTID = spareparts.SPAREPARTID) INNER JOIN supplier ON (spareparts.SUPPLIERCODE = supplier.SUPPLIERCODE) " & _
              " Where " & myW1 & " Order By backorder.TRANSDATE, supplier.SUPPLIERNAME, spareparts.PARTNUMBER;"
    ElseIf updFilter.Value = 3 Then 'Descrition
        MySP = Split(cboFilter.Text, " ")
        myW1 = ""
        For i = LBound(MySP) To UBound(MySP)
            If myW1 = "" Then 'First Criteria
                myW1 = "SPAREPARTS.DESCRIPTION Like '%" & MySP(i) & "%'"
            Else
                myW1 = myW1 & " AND SPAREPARTS.DESCRIPTION Like '%" & MySP(i) & "%'"
            End If
        Next i
        Qry = "SELECT spareparts.PARTNUMBER, spareparts.DESCRIPTION, backorder.REFERENCE, backorder.TRANSDATE, backorder.BRANCHNAME, backorder.BOBALANCE, supplier.SUPPLIERNAME, spareparts.SPAREPARTID, spareparts.COST, backorder.BACKORDERID " & _
              "From backorder INNER JOIN spareparts ON (backorder.SPAREPARTID = spareparts.SPAREPARTID) INNER JOIN supplier ON (spareparts.SUPPLIERCODE = supplier.SUPPLIERCODE) " & _
              " Where " & myW1 & " Order By backorder.TRANSDATE, supplier.SUPPLIERNAME, spareparts.PARTNUMBER;"
    
    ElseIf updFilter.Value = 4 Then 'Classification
        Qry = "SELECT spareparts.PARTNUMBER, spareparts.DESCRIPTION, backorder.REFERENCE, backorder.TRANSDATE, backorder.BRANCHNAME, backorder.BOBALANCE, supplier.SUPPLIERNAME, spareparts.SPAREPARTID, spareparts.COST, backorder.BACKORDERID " & _
              "From backorder INNER JOIN spareparts ON (backorder.SPAREPARTID = spareparts.SPAREPARTID) INNER JOIN supplier ON (spareparts.SUPPLIERCODE = supplier.SUPPLIERCODE) " & _
              "Where spareparts.REMARKS = '" & cboFilter.Text & "' Order By backorder.TRANSDATE, supplier.SUPPLIERNAME, spareparts.PARTNUMBER;"
    ElseIf updFilter.Value = 5 Then 'applicable model
        Qry = "SELECT spareparts.PARTNUMBER, spareparts.DESCRIPTION, backorder.REFERENCE, backorder.TRANSDATE, backorder.BRANCHNAME, backorder.BOBALANCE, supplier.SUPPLIERNAME, spareparts.SPAREPARTID, spareparts.COST, backorder.BACKORDERID " & _
              "From backorder INNER JOIN spareparts ON (backorder.SPAREPARTID = spareparts.SPAREPARTID) INNER JOIN supplier ON (spareparts.SUPPLIERCODE = supplier.SUPPLIERCODE) " & _
              "Where spareparts.APPLICMODEL = '" & cboFilter.Text & "' Order By backorder.TRANSDATE, supplier.SUPPLIERNAME, spareparts.PARTNUMBER;"
    ElseIf updFilter.Value = 6 Then 'SUPPLIER
        Qry = "SELECT spareparts.PARTNUMBER, spareparts.DESCRIPTION, backorder.REFERENCE, backorder.TRANSDATE, backorder.BRANCHNAME, backorder.BOBALANCE, supplier.SUPPLIERNAME, spareparts.SPAREPARTID, spareparts.COST, backorder.BACKORDERID " & _
              "From backorder INNER JOIN spareparts ON (backorder.SPAREPARTID = spareparts.SPAREPARTID) INNER JOIN supplier ON (spareparts.SUPPLIERCODE = supplier.SUPPLIERCODE) " & _
              "Where supplier.SUPPLIERNAME = '" & cboFilter.Text & "' Order By backorder.TRANSDATE, supplier.SUPPLIERNAME, spareparts.PARTNUMBER;"
    ElseIf updFilter.Value = 7 Then 'BRANCH
        Qry = "SELECT spareparts.PARTNUMBER, spareparts.DESCRIPTION, backorder.REFERENCE, backorder.TRANSDATE, backorder.BRANCHNAME, backorder.BOBALANCE, supplier.SUPPLIERNAME, spareparts.SPAREPARTID, spareparts.COST, backorder.BACKORDERID " & _
              "From backorder INNER JOIN spareparts ON (backorder.SPAREPARTID = spareparts.SPAREPARTID) INNER JOIN supplier ON (spareparts.SUPPLIERCODE = supplier.SUPPLIERCODE) " & _
              "Where backorder.BRANCHNAME = '" & cboFilter.Text & "' Order By backorder.TRANSDATE, supplier.SUPPLIERNAME, spareparts.PARTNUMBER;"

    End If
    
    With rsBBO
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            pBar.Max = .RecordCount + 1
            pBar.Value = 1
            pBar.Visible = True
            flxBBO.Rows = 1
            Do While Not .EOF
                flxBBO.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & !REFERENCE & vbTab & Format(!TRANSDATE, "mm/dd/yyyy") & vbTab & IIf(IsNull(!branchName), "", !branchName) & vbTab & !BOBALANCE & vbTab & vbTab & !SUPPLIERNAME & vbTab & !SPAREPARTID & vbTab & !Cost & vbTab & !BACKORDERID
                .MoveNext
                pBar.Value = pBar.Value + 1
            Loop
        Else
            flxBBO.Rows = 1
            flxBBO.Rows = 2
        End If
        .Close
        pBar.Visible = False
    End With
    
End Sub




Private Sub cboFilter_Click()
    Call DisplayBBO
End Sub

Private Sub cboFilter_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 And cboFilter.Text <> "" Then
        Call DisplayBBO
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxBBO
        For i = fRow To .Rows - 1
            For j = fCol To 8
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .row = i
                    .ColSel = 8: .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 8, i + 1, i)
                    fCol = IIf(j = 8, 1, j + 1)
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

Private Sub cmdInsert_Click()
    Dim i As Integer, j As Integer
    
    With BrnBackOrderPO
        If .flxParts.TextMatrix(.flxParts.Rows - 1, 1) = "" Then  'Remove last row if empty
            If .flxParts.Rows = 2 Then
                .flxParts.Rows = 1
            Else
                .flxParts.RemoveItem .flxParts.Rows - 1
            End If
        End If
        If .flxPO.TextMatrix(.flxPO.Rows - 1, 1) = "" Then  'Remove last row if empty
            If .flxPO.Rows = 2 Then
                .flxPO.Rows = 1
            Else
                .flxPO.RemoveItem .flxPO.Rows - 1
            End If
        End If

        For i = 1 To flxBBO.Rows - 1
            If Val(flxBBO.TextMatrix(i, 7)) > 0 And flxBBO.TextMatrix(i, 0) <> "" Then
                If Not SPinFlx(flxBBO.TextMatrix(i, 1)) Then
                    .flxParts.AddItem .flxParts.Rows & vbTab & flxBBO.TextMatrix(i, 1) & vbTab & flxBBO.TextMatrix(i, 2) & vbTab & flxBBO.TextMatrix(i, 7) & vbTab & Format(flxBBO.TextMatrix(i, 10), "#,##0.00") & vbTab & Format(Val(flxBBO.TextMatrix(i, 7)) * Val(flxBBO.TextMatrix(i, 10)), "#,##0.00") & vbTab & flxBBO.TextMatrix(i, 9)
                    .flxPO.AddItem vbTab & flxBBO.TextMatrix(i, 9) & vbTab & flxBBO.TextMatrix(i, 3) & vbTab & flxBBO.TextMatrix(i, 4) & vbTab & flxBBO.TextMatrix(i, 5) & vbTab & flxBBO.TextMatrix(i, 6) & vbTab & flxBBO.TextMatrix(i, 7) & vbTab & flxBBO.TextMatrix(i, 11)
                Else
                    For j = 1 To .flxParts.Rows - 1
                        If flxBBO.TextMatrix(i, 1) = .flxParts.TextMatrix(j, 1) Then 'Partnumber already exist
                            .flxParts.TextMatrix(j, 3) = Val(.flxParts.TextMatrix(j, 3)) + Val(flxBBO.TextMatrix(i, 7))
                            .flxParts.TextMatrix(j, 5) = Format(Val(.flxParts.TextMatrix(j, 3)) * Val(Replace(.flxParts.TextMatrix(j, 4), ",", "")), "#,##0.00")
                            .flxPO.AddItem vbTab & flxBBO.TextMatrix(i, 9) & vbTab & flxBBO.TextMatrix(i, 3) & vbTab & flxBBO.TextMatrix(i, 4) & vbTab & flxBBO.TextMatrix(i, 5) & vbTab & flxBBO.TextMatrix(i, 6) & vbTab & flxBBO.TextMatrix(i, 7) & vbTab & flxBBO.TextMatrix(i, 11)
                            Exit For
                        End If
                    Next j
                End If
                flxBBO.TextMatrix(i, 0) = ""
            End If
        Next i
        If .flxParts.Rows = 1 Then .flxParts.AddItem "1"
    
    End With
    Call BrnBackOrderPO.ComputeTotal
    
End Sub

Function SPinFlx(myPN As String) As Boolean
    Dim i As Integer
    
    With BrnBackOrderPO.flxParts
        For i = 1 To .Rows - 1
            If myPN = .TextMatrix(i, 1) Then
                SPinFlx = True
                Exit Function
            End If
        Next i
    End With
    SPinFlx = False
End Function


Private Sub cmdOK_Click()
    Call cmdInsert_Click
    Unload Me
End Sub

Private Sub flxBBO_Click()
    ActRow = flxBBO.row
    ActCol = flxBBO.Col
    If ActCol <> 7 Then Exit Sub
    With txtOrder
        .Top = flxBBO.CellTop + flxBBO.Top
        .Left = flxBBO.CellLeft + flxBBO.Left
        .Width = flxBBO.CellWidth
        .Text = Trim(flxBBO.Text)
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .ZOrder
        .SetFocus
    End With
    RwSel = True
    flxBBO.RowSel = flxBBO.row
    flxBBO.Col = 0
    flxBBO.ColSel = 10
    RwSel = False
    flxBBO.FillStyle = flexFillRepeat

End Sub

Private Sub flxBBO_EnterCell()
    ActRow = flxBBO.row
    ActCol = flxBBO.Col
End Sub

Private Sub flxBBO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And ActCol = 7 Then Call flxBBO_Click
End Sub

Private Sub flxBBO_Scroll()
    txtOrder.Visible = False
End Sub

Private Sub Form_Load()
    updFilter.Value = 1
End Sub

Private Sub RXResizer1_ResizeComplete()
    With flxBBO
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.15
        .ColWidth(2) = (.Width - 300) * 0.2
        .ColWidth(3) = (.Width - 300) * 0.1
        .ColWidth(4) = (.Width - 300) * 0.08
        .ColWidth(5) = (.Width - 300) * 0.12
        .ColWidth(6) = (.Width - 300) * 0.08
        .ColWidth(7) = (.Width - 300) * 0.1
        .ColWidth(8) = (.Width - 300) * 0.15
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        .ColWidth(11) = 0
    End With

End Sub

Private Sub txtOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        flxBBO.TextMatrix(ActRow, 7) = ""
        flxBBO.TextMatrix(ActRow, 0) = ""
        txtOrder.Visible = False
        flxBBO.Col = 7
    End If
End Sub

Private Sub txtOrder_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        flxBBO.SetFocus
        flxBBO.row = ActRow
        flxBBO.Col = 7
   End If

End Sub

Private Sub txtOrder_LostFocus()
    
    With flxBBO
        If Val(txtOrder.Text) > Val(.TextMatrix(ActRow, 6)) Then
            MsgBox ("Order Quantity must not be Greater than Back-Order Quantity"), vbExclamation, "Order"
            .TextMatrix(ActRow, 7) = ""
            .TextMatrix(ActRow, 0) = ""
            If txtOrder.Visible = True Then txtOrder.SetFocus: SendKeys "{Home}+{End}"
            Exit Sub
        End If  'Else
        .TextMatrix(ActRow, 7) = txtOrder.Text
        txtOrder.Visible = False
        If Val(.TextMatrix(ActRow, 7)) > 0 Then
            .row = ActRow
            .Col = 0
            .CellFontName = "WINGDINGS"
            .CellFontSize = 12
            .Text = lblChk.Caption
            .CellAlignment = 4
            .Col = 7
        Else
            .TextMatrix(ActRow, 7) = ""
            .TextMatrix(ActRow, 0) = ""
            .Col = 7
        End If
    End With

End Sub

Private Sub updFilter_Change()
    Dim rsFilter As New ADODB.Recordset
    Dim Qry As String
    
    If updFilter.Value = 1 Then 'All BBO Parts
        cboFilter.Clear
        cboFilter.AddItem "All Back Order Parts"
        lblFilter.Caption = "SHOW ALL BBO PARTS"
    ElseIf updFilter.Value = 2 Then 'Part Number
        cboFilter.Clear
        lblFilter.Caption = "FILTER BY PARTNUMBER"
    ElseIf updFilter.Value = 3 Then 'Description
        cboFilter.Clear
        lblFilter.Caption = "FILTER BY DESCRIPTION"
    ElseIf updFilter.Value = 4 Then 'Classification
        cboFilter.Clear
        lblFilter.Caption = "SELECT CLASSIFICATION"
    ElseIf updFilter.Value = 5 Then 'Applicable Model
        cboFilter.Clear
        lblFilter.Caption = "SELECT APPLICABLE MODEL"
    ElseIf updFilter.Value = 6 Then 'Supplier
        cboFilter.Clear
        lblFilter.Caption = "SELECT SUPPLIER"
    ElseIf updFilter.Value = 7 Then 'Branch Code
        cboFilter.Clear
        lblFilter.Caption = "SELECT BRANCH"
    End If
    
    If updFilter.Value > 3 Then
        If updFilter.Value = 4 Then     'classification
            Qry = "SELECT DISTINCT spareparts.REMARKS From backorder INNER JOIN spareparts ON (backorder.SPAREPARTID = spareparts.SPAREPARTID) Order By spareparts.REMARKS"
        ElseIf updFilter.Value = 5 Then 'Applicable Model
            Qry = "SELECT DISTINCT spareparts.APPLICMODEL From backorder INNER JOIN spareparts ON (backorder.SPAREPARTID = spareparts.SPAREPARTID) Order By spareparts.APPLICMODEL"
        ElseIf updFilter.Value = 6 Then 'supplier
            Qry = "SELECT DISTINCT supplier.SUPPLIERNAME From backorder INNER JOIN spareparts ON (backorder.SPAREPARTID = spareparts.SPAREPARTID) INNER JOIN supplier ON (spareparts.SUPPLIERCODE = supplier.SUPPLIERCODE) Order By supplier.SUPPLIERNAME"
        ElseIf updFilter.Value = 7 Then 'BRANCH
            Qry = "SELECT DISTINCT backorder.BRANCHNAME From backorder Order By backorder.BRANCHNAME;"
        End If
        With rsFilter
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                pBar.Value = 1
                pBar.Max = .RecordCount + 1
                pBar.Visible = True
                cboFilter.Clear
                Do While Not .EOF
                    cboFilter.AddItem IIf(IsNull(.Fields(0).Value), "", .Fields(0).Value)
                    .MoveNext
                    pBar.Value = pBar.Value + 1
                Loop
                pBar.Visible = False
            Else
                cboFilter.Clear
            End If
            .Close
        End With
    End If
End Sub
