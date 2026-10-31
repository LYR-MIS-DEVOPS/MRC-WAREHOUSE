VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVerifyCst 
   Caption         =   "Spare Parts Price Factor"
   ClientHeight    =   8040
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   10440
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   6360
      Visible         =   0   'False
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9000
      Top             =   240
   End
   Begin lvButton.lvButtons_H cmdverifycost 
      Height          =   372
      Left            =   8640
      TabIndex        =   7
      Tag             =   "xyf"
      Top             =   1200
      Width           =   1212
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Verify Cost"
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
      Image           =   "frmVerifyCst.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdaddcost 
      Height          =   372
      Left            =   8640
      TabIndex        =   6
      Tag             =   "xyf"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1212
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Add Cost"
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
      Image           =   "frmVerifyCst.frx":0C52
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtpackingnum 
      Height          =   288
      Left            =   2280
      TabIndex        =   5
      Top             =   900
      Width           =   3252
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   372
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   1572
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Quit Verification"
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
      Image           =   "frmVerifyCst.frx":0D08
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxverify 
      Height          =   4935
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   8705
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "^ PARTNUMBER       |<DESCRIPTION      |>  QUANTITY  |>    UNIT COST  |>    AMOUNT"
   End
   Begin VB.Image verified 
      Height          =   975
      Left            =   8880
      Picture         =   "frmVerifyCst.frx":12A2
      Tag             =   "xyf"
      Top             =   720
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   6840
      X2              =   9600
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Label lbl12vat 
      Alignment       =   1  'Right Justify
      Caption         =   "12% VAT:"
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
      Left            =   6600
      TabIndex        =   12
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   6720
      Width           =   1335
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   8280
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   10560
      m_MDI_Height    =   8550
   End
   Begin VB.Label Label1 
      Caption         =   "Search Packing Number:"
      Height          =   372
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   2052
   End
   Begin VB.Label lbltotalamount 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Amount:"
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
      Left            =   6600
      TabIndex        =   11
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label withvat 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   8040
      TabIndex        =   9
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label lblgross 
      Alignment       =   1  'Right Justify
      Caption         =   "Gross:"
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
      Left            =   6600
      TabIndex        =   13
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label totalamnt 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   8040
      TabIndex        =   8
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label gross 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   8040
      TabIndex        =   10
      Top             =   7560
      Width           =   1575
   End
End
Attribute VB_Name = "frmVerifyCst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdaddcost_Click()
addcostparam = txtpackingnum.Text
Call addzerocost
Call cmdExit_Click
End Sub

Private Sub cmdExit_Click()
SPpricing.Enabled = True
Unload Me
End Sub
Private Sub loadzerocost()
 Dim rszerocost As New ADODB.Recordset
 Dim Qry As String, i As Integer
 Dim zerocost As Integer, totalamount As Double
 Dim transcost As Integer
 i = 0
 zerocost = 0
 transcost = 0
 pBar.Value = 0
 totalamount = 0
  With rszerocost
  
  flxverify.Rows = 1
  flxverify.FormatString = "^ PARTNUMBER       |<DESCRIPTION      |>  QUANTITY  |>    UNIT COST  |>    AMOUNT  |>    spid  |>    spitcost"
        Qry = "SELECT verifycost.PARTNUMBER, verifycost.DESCRIPTION, verifycost.RECEIVE, verifycost.COST, Format(verifycost.RECEIVE * verifycost.cost,2) AS AMOUNT, verifycost.SPAREPARTID, verifycost.SPITRCOST From verifycost Where verifycost.REFERENCE = '" & txtpackingnum.Text & "' and verifycost.TRANSTYPE ='receive' ORDER BY verifycost.PARTNUMBER, verifycost.COST"
      Call RXResizer1_ResizeComplete
      
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        
        If .RecordCount > 0 Then
        
            .MoveFirst
            Do While Not .EOF
                 flxverify.AddItem !PartNumber & vbTab & !Description & vbTab & !RECEIVE & vbTab & !Cost & vbTab & !Amount & vbTab & !SPAREPARTID & vbTab & !SPITRCOST
                 flxverify.TopRow = flxverify.Rows - 1
                 pBar.Visible = True
                 Label2.Caption = !PartNumber
                 
                 .MoveNext
                pBar.Max = .RecordCount
                pBar.Value = pBar.Value + 1
                Percent = Round((100 * pBar.Value) / .RecordCount)
                Me.Cls
                Label2.Caption = ""
            Loop
        
        Else
            flxverify.AddItem "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & ""
            MsgBox "Please check packing number and try again!"
            Exit Sub
        End If
        
        If .RecordCount > 0 Then
        
            With flxverify
            For j = 1 To .Rows - 1
                If .TextMatrix(j, 3) = 0 Then 'zero cost
                .Col = 0: .Row = j: .ColSel = 4: .RowSel = j: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 128, 128)
                .Col = 2
                zerocost = zerocost + 1
                ElseIf .TextMatrix(j, 6) > 0 Then 'with cost
                .Col = 0: .Row = j: .ColSel = 4: .RowSel = j: .FillStyle = flexFillRepeat: .CellBackColor = RGB(128, 255, 128)
                .Col = 2
                transcost = transcost + 1
                Else 'yellow color mar roxas (di ko alam to!)
                .Col = 0: .Row = j: .ColSel = 4: .RowSel = j: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 255, 165)
                .Col = 2
                End If
            Next j
            End With
        End If
 End With
 
 flxverify.TopRow = 1
 Label2.Caption = Percent & " %"
 
    If zerocost > 0 Then
        MsgBox "Packing Number have " & zerocost & " Partnumber/s with unassigned cost"
    ElseIf zerocost = 0 And transcost = flxverify.Rows - 1 Then
        MsgBox "Packing number already verified!"
    End If
 
    If zerocost = 0 And transcost <> flxverify.Rows - 1 Then
        cmdverifycost.Visible = True
        cmdaddcost.Visible = False
        verified.Visible = False
    ElseIf zerocost = 0 And transcost = flxverify.Rows - 1 Then
        cmdverifycost.Visible = False
        cmdaddcost.Visible = False
        verified.Visible = True
    Else
        cmdverifycost.Visible = False
        cmdaddcost.Visible = True
        verified.Visible = False
    End If
    
    Call RXResizer1_ResizeComplete
    Label2.Caption = ""
    pBar.Visible = False

    Call computepacking
    'For i = 1 To flxverify.Rows - 1
    '
    'totalamount = totalamount + Val(Replace(flxverify.TextMatrix(i, 4), ",", ""))
    '
    'Next i
    'totalamnt.Caption = Format(totalamount, "###,##0.00")
    'withvat.Caption = Format(totalamount * 0.12, "###,##0.00")
    'gross.Caption = Format(totalamount + (totalamount * 0.12), "###,##0.00")
    '
    'totalamnt.Visible = True
    'withvat.Visible = True
    'gross.Visible = True
    'lbltotalamount.Visible = True
    'lbl12vat.Visible = True
    'lblgross.Visible = True
    'Line1.Visible = True
    'MsgBox zerocost
    'MsgBox transcost
    'MsgBox flxverify.Rows - 1
 End Sub

Private Sub Command1_Click()
Call loadzerocost
End Sub

Private Sub cmdverifycost_Click()
Call verifycost
End Sub

Private Sub Form_Load()
    Timer1.Enabled = True
End Sub


Private Sub RXResizer1_ResizeComplete()
With flxverify
        .ColWidth(0) = .Width * 0.2
        .ColWidth(1) = .Width * 0.2
        .ColWidth(2) = .Width * 0.16
        .ColWidth(3) = .Width * 0.2
        .ColWidth(4) = .Width * 0.2
        .ColWidth(5) = .Width * 0
        .ColWidth(6) = .Width * 0
    End With
End Sub

Private Sub Timer1_Timer()
    Dim rszerocost As New ADODB.Recordset
    Dim Qry As String
    Dim zerocost As Integer
    Dim transcost As Integer
    Dim initialcost As Integer
 
    zerocost = 0
    transcost = 0
    initialcost = 0
 
    With rszerocost
        If addcostparam <> "" Then
            Qry = "SELECT verifycost.PARTNUMBER, verifycost.DESCRIPTION, verifycost.RECEIVE, " & _
                    "verifycost.COST, Format(verifycost.RECEIVE * verifycost.cost,2) AS AMOUNT, " & _
                    "verifycost.SPAREPARTID, verifycost.SPITRCOST From verifycost " & _
                    "Where verifycost.REFERENCE = '" & addcostparam & "' and verifycost.TRANSTYPE ='receive' " & _
                    "ORDER BY verifycost.COST"
        
            txtpackingnum.Text = addcostparam
            'cmdverifycost.Visible = True
        Else
            Qry = "SELECT verifycost.PARTNUMBER, verifycost.DESCRIPTION, verifycost.RECEIVE, verifycost.COST," & _
                    "Format(verifycost.RECEIVE * verifycost.cost, 2) AS AMOUNT, verifycost.SPAREPARTID, verifycost.SPITRCOST " & _
                    "From verifycost Where verifycost.COST = '0' AND verifycost.DESCRIPTION NOT LIKE '%(INACTIVE)%' AND " & _
                    "verifycost.TRANSTYPE = 'receive' Order By verifycost.PARTNUMBER "
        End If
        Call RXResizer1_ResizeComplete
          
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
           If .RecordCount > 0 Then
    
                    .MoveFirst
                    flxverify.Rows = 1
                    Do While Not .EOF
                       flxverify.FormatString = "^ PARTNUMBER       |<DESCRIPTION      |>  QUANTITY  |>    UNIT COST  |>    AMOUNT  |>    spid  |>    spitcost"
                       flxverify.AddItem !PartNumber & vbTab & !Description & vbTab & !RECEIVE & vbTab & !Cost & vbTab & !Amount & vbTab & !SPAREPARTID & vbTab & !SPITRCOST
                       .MoveNext
                    Loop
            If .RecordCount > 0 Then
                With flxverify
                    For j = 1 To .Rows - 1
                       If .TextMatrix(j, 3) = 0 And addcostparam = "" Then 'zero cost
                        .Col = 0: .Row = j: .ColSel = 4: .RowSel = j: .FillStyle = flexFillRepeat: .CellBackColor = RGB(198, 223, 255)
                        .Col = 2
                        zerocost = zerocost + 1
                       ElseIf .TextMatrix(j, 3) = 0 And addcostparam <> "" Then
                        .Col = 0: .Row = j: .ColSel = 4: .RowSel = j: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 128, 128)
                        .Col = 2
                        zerocost = zerocost + 1
                        ElseIf .TextMatrix(j, 6) > 0 Then 'with cost
                        .Col = 0: .Row = j: .ColSel = 4: .RowSel = j: .FillStyle = flexFillRepeat: .CellBackColor = RGB(128, 255, 128)
                        .Col = 2
                        transcost = transcost + 1
                        Else
                        .Col = 0: .Row = j: .ColSel = 4: .RowSel = j: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 255, 165)
                        .Col = 2
                        initialcost = initialcost + 1
                        End If
                    Next j
                End With
            End If
        End If
        Timer1.Enabled = False
    
    End With
        Call computepacking
        Call RXResizer1_ResizeComplete
      
    If addcostparam = "" Then
        cmdverifycost.Visible = False
        cmdaddcost.Visible = False
        verified.Visible = False
    ElseIf zerocost = 0 And initialcost <> flxverify.Rows - 1 Then
        cmdverifycost.Visible = True
        cmdaddcost.Visible = False
        verified.Visible = False
        
    ElseIf zerocost = 0 And initialcost = flxverify.Rows - 1 Then
        cmdverifycost.Visible = True
        cmdaddcost.Visible = False
        verified.Visible = False
        
    ElseIf zerocost > 0 Then
     
        verified.Visible = False
        cmdverifycost.Visible = False
        cmdaddcost.Visible = True
    Else
        cmdverifycost.Visible = False
        cmdaddcost.Visible = True
        verified.Visible = False
    End If
End Sub

Private Sub txtpackingnum_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "'" Then KeyAscii = 0
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
    
        verified.Visible = False
        lbltotalamount.Visible = False
        lbl12vat.Visible = False
        lblgross.Visible = False
        Line1.Visible = False
        totalamnt.Visible = False
        withvat.Visible = False
        gross.Visible = False
        Call loadzerocost
    End If
End Sub

Private Sub addzerocost()
Dim rsParts As New Recordset
Qry = "SELECT spareparts.PARTNUMBER, spareparts.DESCRIPTION, Format(SPAREPARTS.COST, 2) AS CWOV, Format(SPAREPARTS.STOREPRICE, 2) AS SRP, spareparts.SPAREPARTID, spareparts.PRICEHIST," & _
  "spareparts.PRICEREF, spareparts.SPFID, Format(SPAREPARTS.DLRPRICE, 2) AS DP, DATE_FORMAT(spareparts.LASTUPDATE, '%m/%d/%Y') AS DATE_ENTRY, DATE_FORMAT(spareparts.EFFECTIVEDATE, '%m/%d/%Y') AS EFFECTIVE_DATE," & _
  "DATE_FORMAT(spareparts.EFFECTIVEDATEDP, '%m/%d/%Y') AS EFFECTIVE_DATE_DP, FORMAT(SPAREPARTS.RETAILPR, 2) AS FIELD_1,  FORMAT(spareparts.dealerPR, 2) As FIELD_2, spareparts.QTYONHAND  " & _
  "From spareparts INNER JOIN spitr ON (spareparts.SPAREPARTID = spitr.SPAREPARTID) Where spitr.REFERENCE = '" & txtpackingnum.Text & "' AND  spareparts.cost = 0"
  
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set SPpricing.flxParts.DataSource = rsParts
    SPpricing.flxParts.Refresh
 
 Call SPpricing.flxParts_Click

End Sub

Private Sub verifycost()
With flxverify
 For z = 1 To .Rows - 1
  LYRcon.Execute "UPDATE SPITR SET COST = '" & .TextMatrix(z, 3) & "' WHERE spitr.SPAREPARTID = '" & .TextMatrix(z, 5) & "' "
  .Col = 0: .Row = z: .ColSel = 4: .RowSel = z: .FillStyle = flexFillRepeat: .CellBackColor = RGB(128, 255, 128)
  .Col = 2
  Next z
  
  Ans = MsgBox("Successful Transaction Cost verification!" & vbCrLf & vbCrLf & _
          "Do you want to verify another Packing Number?", vbQuestion + vbYesNo, "Verified")
    verified.Visible = True
    totalamnt.Visible = True
    withvat.Visible = True
    gross.Visible = True
    lbltotalamount.Visible = True
    lbl12vat.Visible = True
    lblgross.Visible = True
    Line1.Visible = True
 If Ans = vbYes Then
        txtpackingnum.Text = ""
        txtpackingnum.SetFocus
        cmdverifycost.Visible = False
        Label2.Caption = ""
        pBar.Visible = False
        addcostparam = ""
        
  Else
  
 Call cmdExit_Click
 Call SPpricing.cmdRefresh_Click
 'Call SPpricing.cmdRefresh_Click
 End If
End With

End Sub


Private Sub computepacking()
Dim totalamount As Double
For i = 1 To flxverify.Rows - 1

totalamount = totalamount + Val(Replace(flxverify.TextMatrix(i, 4), ",", ""))

Next i

totalamnt.Caption = Format(totalamount, "###,##0.00")
withvat.Caption = Format(totalamount * 0.12, "###,##0.00")
gross.Caption = Format(totalamount + (totalamount * 0.12), "###,##0.00")

totalamnt.Visible = True
withvat.Visible = True
gross.Visible = True
lbltotalamount.Visible = True
lbl12vat.Visible = True
lblgross.Visible = True
Line1.Visible = True

End Sub
