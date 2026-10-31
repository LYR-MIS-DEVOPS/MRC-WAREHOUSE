VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form SPCardexRed 
   Caption         =   "WAREHOUSE CARDEX [RED]"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   11265
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid flxOrd 
      Height          =   1695
      Left            =   11640
      TabIndex        =   18
      Top             =   3840
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedCols       =   0
      FormatString    =   "Date          |< PO Number       |<Qty         "
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxParts 
      Height          =   1815
      Left            =   360
      TabIndex        =   12
      Top             =   7800
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3201
      _Version        =   393216
      BackColor       =   16777215
      FixedRows       =   0
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      GridColorUnpopulated=   16777215
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Tag             =   "XYF"
      Top             =   720
      Width           =   11055
      Begin VB.TextBox txtWSBackOrder 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtBackOrder 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   17
         Tag             =   "XYF"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtLocation 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   15
         Tag             =   "XYF"
         Top             =   1080
         Width           =   4575
      End
      Begin VB.TextBox txtPrice 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txtDesc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   480
         Width           =   7695
      End
      Begin VB.TextBox txtPartNo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label7 
         Caption         =   "WS BACK ORDER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9480
         TabIndex        =   20
         Tag             =   "XYF"
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "WH BACK ORDER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7920
         TabIndex        =   16
         Tag             =   "XYF"
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "LOCATION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   14
         Tag             =   "XYF"
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "PRICE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "DESCRIPTION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "PART NUMBER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   240
         Width           =   2415
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxStock 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Tag             =   "XYF"
      Top             =   2280
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   9975
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   16777215
      ScrollBars      =   2
      FormatString    =   "^    Date    |^   P.O. Number      |^  Packing List     |^WH Ord |^ WH Rec |^WS Ord |^WS Rec |^ Cancelled |^  Balance   "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "E&xit Cardex"
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
      Image           =   "SPCardexRed.frx":0000
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   3240
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11505
      m_MDI_Height    =   8580
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CARDEX [RED]"
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
      Index           =   5
      Left            =   6360
      TabIndex        =   4
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CARDEX [RED]"
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
      Height          =   255
      Index           =   4
      Left            =   6390
      TabIndex        =   3
      Top             =   510
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MOTO-R REPUBLIC"
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
      Index           =   3
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MOTO-R REPUBLIC"
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
      Height          =   255
      Index           =   2
      Left            =   150
      TabIndex        =   1
      Top             =   510
      Width           =   2415
   End
   Begin VB.Menu mnuStockCard 
      Caption         =   "StockCard"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "SPCardexRed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsParts As New ADODB.Recordset
Dim PrtSrch As Boolean, PrtBln As Boolean, FromFlx As Boolean

Private Sub LoadSpareParts()
    Dim rsSet As New ADODB.Recordset
    Dim Qry As String
    Dim pPct As Single
    
    Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.QTYONHAND, SPAREPARTS.LOCATION " & _
          "From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxParts.DataSource = rsParts
    With flxParts
        .Width = txtPartNo.Width
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = txtPartNo.Width
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .ColWidth(6) = 0
    End With
End Sub

Private Sub cmdExit_Click()
    Unload Me
    'MotoRep.tbMain.Visible = True
End Sub

Private Sub flxParts_Click()
    FromFlx = True
End Sub

Private Sub flxParts_EnterCell()
    With flxParts
        FromFlx = True
        txtPartNo.Text = .TextMatrix(.row, 2)
        txtDesc.Text = .TextMatrix(.row, 3)
        txtPrice.Text = Format$(.TextMatrix(.row, 4), "###,##0.00")
        txtLocation.Text = .TextMatrix(.row, 6)
    End With
End Sub

Private Sub flxParts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call DisplayCardex(flxParts.TextMatrix(flxParts.row, 1))
        flxStock.SetFocus
        flxStock.row = flxStock.Rows - 1
        flxStock.col = 0
        flxStock.RowSel = flxStock.Rows - 1
        flxStock.ColSel = 4
    End If
End Sub

Private Sub flxParts_LostFocus()
    flxParts.Visible = False
End Sub

Private Sub Form_Load()
    
    Call LoadSpareParts
    
    'MotoRep.sBar.Panels(6).Text = "Spare Parts CARDEX [RED]"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsParts.State = adStateOpen Then rsParts.Close
End Sub

Private Sub RXResizer1_ResizeComplete()
    With flxStock
        .ColWidth(0) = .Width * 0.14
        .ColWidth(1) = .Width * 0.16
        .ColWidth(2) = .Width * 0.14
        .ColWidth(3) = .Width * 0.09
        .ColWidth(4) = .Width * 0.09
        .ColWidth(5) = .Width * 0.09
        .ColWidth(6) = .Width * 0.09
        .ColWidth(7) = .Width * 0.09
        .ColWidth(8) = .Width * 0.08
    End With
    Label3(2).Left = Frame1.Left + 30
    Label3(3).Left = Frame1.Left
    Label3(2).Top = Frame1.Top - Label3(2).Height + 30
    Label3(3).Top = Frame1.Top - Label3(3).Height
    Label3(4).Left = Frame1.Left + Frame1.Width - Label3(4).Width + 30
    Label3(5).Left = Frame1.Left + Frame1.Width - Label3(5).Width
    Label3(4).Top = Frame1.Top - Label3(4).Height + 30
    Label3(5).Top = Frame1.Top - Label3(5).Height
    
    flxParts.ColWidth(2) = txtPartNo.Width
    'flxParts.RowHeightMin = txtPartNo.Height
    flxParts.Font.Size = txtPartNo.FontSize
End Sub


Private Sub DisplayCardex(MyID As Long)
    Dim rsOrder As New ADODB.Recordset
    Dim rsRec As New ADODB.Recordset
    Dim Qry As String, MyPO As String, myDeptCode As String
    Dim xRow As Long, xBal As Long, i As Long
    
    'Clear FlxStock
    flxStock.Rows = 1
    myDeptCode = ""
    
    Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.`ORDER`, SPITR.REMARKS, SPITR.DEPTCODE From SPITR " & _
          "WHERE SPITR.TRANSTYPE='ORDER' AND SPITR.SPAREPARTID=" & MyID & " ORDER BY SPITR.TRANSDATE;"
 
    flxOrd.Rows = 1
    With rsOrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxOrd.AddItem Format(!TRANSDATE, "mm/dd/yyyy") & vbTab & !REFERENCE & vbTab & !Order & vbTab & IIf(IsNull(!Remarks), "", !Remarks) & vbTab & IIf(IsNull(!DEPTCODE), "WH", !DEPTCODE)
               .MoveNext
            Loop
        End If
        .Close
    End With
    xBal = 0
    With flxOrd
        For i = 1 To .Rows - 1
            'Display PO in Stock Card and obtain Balance
            xBal = xBal + Val(.TextMatrix(i, 2))
            If .TextMatrix(i, 4) = "WH" Then
                flxStock.AddItem .TextMatrix(i, 0) & vbTab & .TextMatrix(i, 1) & vbTab & .TextMatrix(i, 3) & vbTab & .TextMatrix(i, 2) & vbTab & vbTab & vbTab & vbTab & vbTab & xBal    '.TextMatrix(i, 2)
            Else
                flxStock.AddItem .TextMatrix(i, 0) & vbTab & .TextMatrix(i, 1) & vbTab & .TextMatrix(i, 3) & vbTab & vbTab & vbTab & .TextMatrix(i, 2) & vbTab & vbTab & vbTab & xBal    '.TextMatrix(i, 2)
            End If
            'Acquire Packing list for P.O.
            MyPO = .TextMatrix(i, 1)
            'Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.RECEIVE, SPITR.PACKINGLIST From SPITR " & _
                  "WHERE (((SPITR.TRANSTYPE)='RECEIVE') AND ((SPITR.SPAREPARTID)=" & MyID & ") AND ((SPITR.PACKINGLIST)='" & MyPO & "')) ORDER BY SPITR.TRANSDATE;"
            'mySQL
            Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.RECEIVE, CASE WHEN SPITR.TRANSTYPE = 'RECEIVE' THEN SPITR.PACKINGLIST ELSE '' END as PACKINGLIST, SPITR.`OUT`, SPITR.DEPTCODE From SPITR " & _
                  "WHERE SPITR.TRANSTYPE IN ('RECEIVE','BACK ORDER ADJ.') AND SPITR.SPAREPARTID=" & MyID & " AND SPITR.PACKINGLIST='" & MyPO & "' ORDER BY SPITR.TRANSDATE;"
            
            rsRec.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If rsRec.RecordCount > 0 Then
                rsRec.MoveFirst
                Do While Not rsRec.EOF
                    xBal = xBal - (rsRec!RECEIVE + rsRec!OUT)
                    If rsRec!DEPTCODE = "WH" Or IsNull(rsRec!DEPTCODE) Then
                        flxStock.AddItem Format(rsRec!TRANSDATE, "mm/dd/yyyy") & vbTab & MyPO & vbTab & rsRec!REFERENCE & vbTab & vbTab & rsRec!RECEIVE & vbTab & vbTab & vbTab & IIf(rsRec!OUT = 0, "", rsRec!OUT) & vbTab & xBal
                    Else
                        flxStock.AddItem Format(rsRec!TRANSDATE, "mm/dd/yyyy") & vbTab & MyPO & vbTab & rsRec!REFERENCE & vbTab & vbTab & vbTab & vbTab & rsRec!RECEIVE & vbTab & IIf(rsRec!OUT = 0, "", rsRec!OUT) & vbTab & xBal
                    End If
                    rsRec.MoveNext
                Loop
            End If
            rsRec.Close
        Next i
    End With
    Call ComputeTotal
End Sub

Public Sub ComputeTotal()
    Dim myTtlWHORDER As Long, myTtlWHRECEIVE As Long
    Dim myTtlWSORDER As Long, myTtlWSRECEIVE As Long
    Dim myTtlWHCANCELLED As Long, myTtlWSCANCELLED As Long
    Dim myDeptCode As String
    
    txtBackOrder.Text = 0
    txtWSBackOrder.Text = 0
    myTtlWHORDER = 0
    myTtlWHRECEIVE = 0
    myTtlWSORDER = 0
    myTtlWSRECEIVE = 0
    
    For i = 1 To flxStock.Rows - 1
        myTtlWHORDER = myTtlWHORDER + Val(flxStock.TextMatrix(i, 3))
        myTtlWHRECEIVE = myTtlWHRECEIVE + Val(flxStock.TextMatrix(i, 4))
        myTtlWSORDER = myTtlWSORDER + Val(flxStock.TextMatrix(i, 5))
        myTtlWSRECEIVE = myTtlWSRECEIVE + Val(flxStock.TextMatrix(i, 6))
        
        myDeptCode = GetFieldValue("SELECT DEPTCODE FROM SPITR WHERE REFERENCE = '" & flxStock.TextMatrix(i, 1) & "' LIMIT 1")
        If myDeptCode = "WH" And Val(flxStock.TextMatrix(i, 7)) > 0 Then
            myTtlWHCANCELLED = myTtlWHCANCELLED + Val(flxStock.TextMatrix(i, 7))
        ElseIf myDeptCode = "WS" And Val(flxStock.TextMatrix(i, 7)) > 0 Then
            myTtlWSCANCELLED = myTtlWSCANCELLED + Val(flxStock.TextMatrix(i, 7))
        End If
    Next i
    
    txtBackOrder = myTtlWHORDER - myTtlWHRECEIVE - myTtlWHCANCELLED
    txtWSBackOrder = myTtlWSORDER - myTtlWSRECEIVE - myTtlWSCANCELLED
End Sub
Function FinalOrder(MyRef As String) As Boolean
    Dim rsOrd As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT SPORDER.REFERENCE, SPORDER.REMARKS, SPORDER.PRINTCOPY From SPORDER " & _
          "WHERE (((SPORDER.REFERENCE)='" & MyRef & "'));"
    'mySQL
    Qry = "SELECT SPORDER.REFERENCE, SPORDER.REMARKS, SPORDER.PRINTCOPY From SPORDER " & _
          "WHERE SPORDER.REFERENCE='" & MyRef & "';"
    
    With rsOrd
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            If IsNull(!Remarks) Then
                FinalOrder = IIf(!PrintCopy > 0, True, False)
            Else
                FinalOrder = IIf(!Remarks <> "Temp", True, False)
            End If
        Else
            FinalOrder = False
        End If
        .Close
    End With
    
End Function

Private Sub txtPartNo_Change()
    With flxParts
        If PrtSrch Then
            PrtSrch = False
            Exit Sub
        End If
        .Top = Frame1.Top + txtPartNo.Height + txtPartNo.Top
        .Left = Frame1.Left + txtPartNo.Left
        .Width = txtPartNo.Width
        .Visible = True
        If txtPartNo.Text = "" Then .Visible = False: Exit Sub
        rsParts.Find "[PARTNUMBER] LIKE '" + txtPartNo.Text + "%'", , adSearchForward, 1
        If Not rsParts.EOF Then
            If Not PrtBln Then
                txtPartNo.Tag = Len(txtPartNo.Text)
                PrtBln = True
            End If
            If Not FromFlx Then     'top row only if user type @ txtpartno
                .TopRow = rsParts.AbsolutePosition - 1
                FromFlx = False
            End If
            .row = rsParts.AbsolutePosition - 1
            .col = 0: .ColSel = 3
            txtPartNo.Text = rsParts!PartNumber
            'Insert Additional retrieval here
            txtDesc.Text = flxParts.TextMatrix(flxParts.row, 3)
            txtPrice.Text = Format$(flxParts.TextMatrix(flxParts.row, 4), "###,##0.00")
            txtLocation.Text = flxParts.TextMatrix(flxParts.row, 6)
            flxStock.Rows = 1
            'Call DisplayCardex(flxParts.TextMatrix(flxParts.Row, 1))
            
            txtPartNo.SelStart = Val(txtPartNo.Tag)
            txtPartNo.SelLength = Len(txtPartNo.Text)
            PrtBln = False
        Else
            txtDesc.Text = ""
            txtPrice.Text = ""
            flxStock.Rows = 1
            flxStock.Rows = 2
            .Visible = False
        End If
    End With
    
End Sub
Private Sub txtPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdExit.SetFocus
            flxParts.Visible = False
        Case vbKeyDelete
            PrtSrch = Truem
        Case vbKeyBack
            PrtSrch = True
        Case vbKeyDown
            flxParts.Visible = True
            flxParts.SetFocus
        Case Else
            FromFlx = False
    End Select

End Sub

Private Sub txtPartNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtPartNo.Text <> "" Then
            Call DisplayCardex(flxParts.TextMatrix(flxParts.row, 1))
        End If
        flxParts.Visible = False
        flxStock.SetFocus
        flxStock.row = flxStock.Rows - 1
        flxStock.col = 0
        flxStock.RowSel = flxStock.Rows - 1
        flxStock.ColSel = 4
    End If
End Sub

