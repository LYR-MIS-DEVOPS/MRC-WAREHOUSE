VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form findBBO 
   Caption         =   "Find Branch Back Order"
   ClientHeight    =   8220
   ClientLeft      =   120
   ClientTop       =   -2760
   ClientWidth     =   17520
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   17520
   StartUpPosition =   1  'CenterOwner
   Begin lvButton.lvButtons_H cmdLoad 
      Height          =   375
      Left            =   14760
      TabIndex        =   18
      Tag             =   "XYF"
      ToolTipText     =   "Insert Existing Parts [F12]"
      Top             =   1080
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   661
      Caption         =   "LOAD ALL"
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
      Image           =   "findBBO.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   375
      Left            =   16080
      TabIndex        =   17
      Tag             =   "xyf"
      Top             =   1080
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   661
      Caption         =   "PRINT"
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
      Image           =   "findBBO.frx":059A
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtBranchAddress 
      Height          =   285
      Left            =   14880
      TabIndex        =   22
      Top             =   10440
      Width           =   2415
   End
   Begin VB.TextBox txtAllocBranch 
      Height          =   375
      Left            =   14880
      TabIndex        =   21
      Top             =   9120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtRowNum 
      Height          =   285
      Left            =   15000
      TabIndex        =   20
      Top             =   9600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtBackOrder 
      Height          =   375
      Left            =   15120
      TabIndex        =   19
      Top             =   9960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtBranchID 
      Height          =   285
      Left            =   14880
      TabIndex        =   16
      Top             =   8760
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1680
      Top             =   7800
   End
   Begin VB.TextBox txtBBOaloc 
      Height          =   285
      Left            =   14880
      TabIndex        =   15
      Top             =   8400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtAVLQ 
      Height          =   375
      Left            =   9000
      TabIndex        =   14
      Top             =   9000
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtSPID 
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      Top             =   9000
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid flxRecDtls 
      Height          =   3135
      Left            =   240
      TabIndex        =   12
      Top             =   8400
      Visible         =   0   'False
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   5530
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      FormatString    =   $"findBBO.frx":0B34
   End
   Begin MSComctlLib.ProgressBar Pbar 
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Tag             =   "xyf"
      Top             =   7440
      Visible         =   0   'False
      Width           =   17175
      _ExtentX        =   30295
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1080
      Top             =   7800
   End
   Begin MSFlexGridLib.MSFlexGrid flxBBO 
      Height          =   6255
      Left            =   240
      TabIndex        =   1
      Tag             =   "xyf"
      Top             =   1440
      Width           =   17175
      _ExtentX        =   30295
      _ExtentY        =   11033
      _Version        =   393216
      Cols            =   22
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   -2147483633
      ScrollBars      =   2
      AllowUserResizing=   1
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
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Tag             =   "xyf"
      Top             =   240
      Width           =   17175
      Begin VB.ComboBox cboBranch 
         Height          =   315
         Left            =   13440
         TabIndex        =   10
         Tag             =   "xyf"
         Top             =   240
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker dTo 
         Height          =   315
         Left            =   4200
         TabIndex        =   6
         Tag             =   "xyf"
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Format          =   95617025
         CurrentDate     =   42114
      End
      Begin MSComCtl2.DTPicker dFrom 
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Tag             =   "xyf"
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Format          =   95617025
         CurrentDate     =   42114
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   315
         Left            =   6600
         TabIndex        =   8
         Tag             =   "xyf"
         ToolTipText     =   "Execute"
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
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
         Image           =   "findBBO.frx":0C1F
         cBack           =   -2147483633
      End
      Begin VB.Label Label3 
         Caption         =   "BRANCH:"
         Height          =   255
         Left            =   12480
         TabIndex        =   9
         Tag             =   "xyf"
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "TO"
         Height          =   255
         Left            =   3840
         TabIndex        =   5
         Tag             =   "xyf"
         Top             =   300
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "DATE RECEIVED:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Tag             =   "xyf"
         Top             =   300
         Width           =   1455
      End
   End
   Begin lvButton.lvButtons_H cmdOk 
      Height          =   375
      Index           =   4
      Left            =   13800
      TabIndex        =   2
      Tag             =   "xyf"
      Top             =   7800
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "Ok"
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
      Image           =   "findBBO.frx":11B9
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Index           =   5
      Left            =   15480
      TabIndex        =   3
      Tag             =   "xyf"
      Top             =   7800
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "findBBO.frx":1753
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   360
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   17760
      m_MDI_Height    =   8790
   End
   Begin VB.Menu mnuRows 
      Caption         =   "BBO"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "DELETE ROW"
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "Print"
      Visible         =   0   'False
      Begin VB.Menu mnuPicking 
         Caption         =   "PICKING"
      End
      Begin VB.Menu mnuFinal 
         Caption         =   "FINAL"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "findBBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim recRow As Integer, recCol As Integer, bboRow As Integer, fRow As Integer, fCol As Integer
Dim ActRow As Integer, ActCol As Integer, i As Integer, ans As String
Dim SelectBranch As Boolean

Private Sub cboBranch_LostFocus()
    Dim Qry As String
    
        If Val(txtBranchID.Text) = 0 Then Exit Sub
        Qry = "SELECT BRANCH.ADDRESS From BRANCH WHERE BRANCH.BRANCHID=" & Val(txtBranchID.Text) & ";"
        txtBranchAddress.Text = GetFieldValue(Qry)
End Sub

Private Sub cmdLoad_Click()
    cboBranch.Text = ""
    txtBranchID = ""
    Call loadBBO
    Call cmdRefresh_Click
End Sub

Private Sub cmdPrint_Click()
    If cboBranch.Text = "" Then
        MsgBox "Please select branch to process", vbInformation, "NO BRANCH SELECTED"
        Exit Sub
    Else
        Call PopupMenu(mnuPrint)
    End If
End Sub


Private Sub Form_Load()
    
    dFROM.Value = DateSerial(Year(Date), Month(Date), 1)
    dTO.Value = Date
    Timer1.Enabled = True
    Timer2.Enabled = True
    
End Sub


Public Sub recDtls()
    Dim z1 As Integer, z2 As Integer
    
    With flxBBO
        For z1 = 1 To .Rows - 1
            For z2 = 1 To flxRecDtls.Rows - 1
                If .TextMatrix(z1, 15) = flxRecDtls.TextMatrix(z2, 0) And .TextMatrix(z1, 0) = flxRecDtls.TextMatrix(z2, 1) Then
                    .TextMatrix(z1, 3) = flxRecDtls.TextMatrix(z2, 2)   'receiving number
                    .TextMatrix(z1, 4) = flxRecDtls.TextMatrix(z2, 3)   'rec date
                    .TextMatrix(z1, 7) = IIf(flxRecDtls.TextMatrix(z2, 6) = "", 0, flxRecDtls.TextMatrix(z2, 6))   'receive qty
                    .TextMatrix(z1, 8) = IIf(flxRecDtls.TextMatrix(z2, 7) = "", 0, flxRecDtls.TextMatrix(z2, 7))   'rectemp qty
                    .TextMatrix(z1, 10) = IIf(flxRecDtls.TextMatrix(z2, 10) = "", 0, flxRecDtls.TextMatrix(z2, 10)) 'rst
                    .TextMatrix(z1, 11) = IIf(flxRecDtls.TextMatrix(z2, 12) < 0, 0, flxRecDtls.TextMatrix(z2, 12)) 'avlQ
                End If
            Next z2
        Next z1
        For z1 = .Rows - 1 To 1 Step -1
            If .TextMatrix(z1, 3) = "" Then .TextMatrix(z1, 3) = "-----------" 'receiving number
            If .TextMatrix(z1, 4) = "" Then .TextMatrix(z1, 4) = "-----------"      'rec date
            If .TextMatrix(z1, 7) = "" Then .TextMatrix(z1, 7) = 0                'receive qty
            If .TextMatrix(z1, 8) = "" Then .TextMatrix(z1, 8) = 0                 'rectemp qty
            If .TextMatrix(z1, 10) = "" Then .TextMatrix(z1, 10) = 0                'rst
            If .TextMatrix(z1, 11) = "" Then                                        'avlQ
                .TextMatrix(z1, 11) = .TextMatrix(z1, 9) - .TextMatrix(z1, 10)
                    If .TextMatrix(z1, 11) < 0 Then .TextMatrix(z1, 11) = 0
            End If
        Next z1
    End With
        
End Sub


Private Sub cmdOK_Click(Index As Integer)
    If cboBranch.Text = "" Then
        MsgBox "Please select branch to process", vbInformation, "NO BRANCH SELECTED"
        Exit Sub
    End If
    
'    ans = MsgBox("You want to process the allocated Quantity?", vbQuestion + vbYesNo, "Allocation")
'    If ans = vbNo Then Exit Sub
    
    With AllocBBO.flxAllocRec
        '.Rows = 1
        If .TextMatrix(.Rows - 1, 10) = "" And .TextMatrix(.Rows - 1, 1) = "" Then 'Remove last row if empty
            If .Rows = 2 Then
                .Rows = 1
            Else
                .RemoveItem .Rows - 1
            End If
        End If
        For i = 1 To flxBBO.Rows - 1
            If txtBranchID.Text = flxBBO.TextMatrix(i, 17) And flxBBO.TextMatrix(i, 13) <> "" Then
                If Not SPinFlx(flxBBO.TextMatrix(i, 16)) Then
                    '                       'partnumber                 'description                                'onhand                     'alloc.qty                      'PO ref                             'po date                            'rec ref                        'rec date                             'spid                             'boid                           'SP COST
                    .AddItem vbTab & flxBBO.TextMatrix(i, 1) & vbTab & flxBBO.TextMatrix(i, 2) & vbTab & flxBBO.TextMatrix(i, 11) & vbTab & flxBBO.TextMatrix(i, 13) & vbTab & flxBBO.TextMatrix(i, 5) & vbTab & flxBBO.TextMatrix(i, 6) & vbTab & flxBBO.TextMatrix(i, 3) & vbTab & flxBBO.TextMatrix(i, 4) & vbTab & flxBBO.TextMatrix(i, 15) & vbTab & flxBBO.TextMatrix(i, 16) & vbTab & flxBBO.TextMatrix(i, 21)
                End If
            End If
        Next i
        If .Rows = 1 Then .Rows = 2
    End With
        AllocBBO.txtbranch = cboBranch.Text
        AllocBBO.txtBranchID = txtBranchID.Text
        AllocBBO.lblAddress = txtBranchAddress.Text
        
End Sub

Function SPinFlx(boID As Long) As Boolean
    Dim i As Integer
    
        With AllocBBO.flxAllocRec
            For i = 1 To .Rows - 1
                If boID = Val(.TextMatrix(i, 10)) Then
                    SPinFlx = True
                    Exit Function
                End If
            Next i
        End With
        
        SPinFlx = False
        Exit Function
    
End Function


Private Sub cboBranch_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
    
    If PmanSrch Then
        PmanSrch = False
        Exit Sub
    End If
    With cboBranch
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            txtBranchID.Text = .ItemData(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                PmanSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtBranchID.Text = ""
        End If
    End With

    With flxBBO
        For i = 1 To .Rows - 1
            If txtBranchID.Text = .TextMatrix(i, 17) Then
                .Col = 1: .row = i: .ColSel = 14: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                .Col = 2
                'SelectBranch = True
            Else
                .Col = 1: .row = i: .ColSel = 14: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 250)
                .Col = 2
            End If
                'SelectBranch = False
            
        Next i
    End With
    
    If cboBranch.Text <> "" Then
        Call BranchOnly
    End If
End Sub

Private Sub cboBranch_Click()
    
    If txtAllocBranch.Text = "" Or cboBranch.Text = txtAllocBranch.Text Then
        Call loadBBO
        Call recDtls
        Call cboBranch_Change
    Else
        If cboBranch.Text <> txtAllocBranch.Text Then
            MsgBox "You can't allocate another branch!", vbCritical, "BPO-BO"
            cboBranch.Clear
            Call Timer2_Timer
            Exit Sub
        End If
    End If
    
End Sub

Private Sub cboBranch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then PmanSrch = True
End Sub

Private Sub cboBranch_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then flxBBO.SetFocus
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub mnuPicking_Click()
    
    If Val(txtBranchID.Text) = 0 Then
        MsgBox "Please select branch", vbExclamation + vbOKOnly, "Picking Back Order"
        Exit Sub
    End If
    ans = MsgBox("Do you want to Print Allocation of Branch Back Order?", vbQuestion + vbYesNo, "Print")
    If ans = vbNo Then Exit Sub
    
    LYRcon.Execute "DELETE From temptbl"
    With flxBBO
        For i = 1 To flxBBO.Rows - 1
            LYRcon.Execute "INSERT INTO temptbl VALUES('" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 2) & "', " & _
                            "'" & Val(.TextMatrix(i, 8)) & "', '" & Val(.TextMatrix(i, 9)) & "', '" & Val(.TextMatrix(i, 10)) & "', '" & Val(.TextMatrix(i, 11)) & "')"
        Next i
    End With
    
    frmPrint.MySource = "PICKBBO"
    frmPrint.QryPrn = "SELECT temptbl.PARTNUMBER, temptbl.DESCRIPTION, temptbl.TRECQTY, temptbl.QOH, temptbl.RST, temptbl.AVLQ From temptbl"
    frmPrint.ParamValue1 = cboBranch.Text
    Me.Hide
    frmPrint.Show

End Sub


Private Sub Timer1_Timer()
    
    Call loadBBO
    Timer1.Enabled = False
            
End Sub
    
Private Sub flxBBO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxBBO_Click
End Sub
    
Private Sub flxBBO_Click()
    ActRow = flxBBO.row
    ActCol = flxBBO.Col
    
    If ActCol <> 13 Then
        Exit Sub
    ElseIf cboBranch.Text = "" Then
        MsgBox "Please select branch to allocate", vbCritical And vbOKOnly, "Select Branch"
        Exit Sub
    ElseIf flxBBO.TextMatrix(ActRow, 17) <> txtBranchID.Text Then
        MsgBox "You are not allow to allocate", vbCritical And vbOKOnly, "Branch Selected"
        Exit Sub
    End If
    
        With txtBBOaloc
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
        
        flxBBO.Col = 0
        flxBBO.ColSel = 14
        flxBBO.RowSel = ActRow
        flxBBO.FillStyle = flexFillRepeat
    
    
End Sub

Private Sub flxBBO_Scroll()
        txtBBOaloc.Visible = False
End Sub

Private Sub loadBBO()
    
    Dim rsBBO As New ADODB.Recordset
    
        Qry = "SELECT spareparts.PARTNUMBER, spareparts.DESCRIPTION, spareparts.cost, backorder.REFERENCE, backorder.TRANSDATE, " & _
                "spareparts.QTYONHAND, backorder.BOBALANCE, backorder.BRANCHNAME, backorder.SPAREPARTID, backorder.BACKORDERID, backorder.BRANCHID, backorder.SUPPORDERREF, backorder.SUPPOSTATUS, backorder.POQTY " & _
                "From backorder INNER JOIN spareparts ON (backorder.SPAREPARTID = spareparts.SPAREPARTID) " & _
                "Where backorder.BOBALANCE > 0 Order By backorder.TRANSDATE, spareparts.PARTNUMBER"
                
        With rsBBO
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    flxBBO.Rows = 1
                    pBar.Value = 0
                    pBar.Visible = True
                    pBar.Max = .RecordCount
                        Do While Not .EOF
                            reserveQty = ComputeQty(!SPAREPARTID)
                            flxBBO.AddItem flxBBO.Rows & vbTab & !PartNumber & vbTab & !Description & vbTab & vbTab & vbTab & !REFERENCE & vbTab & !TRANSDATE & vbTab & vbTab & vbTab & !QtyOnHand & vbTab & reserveQty & vbTab & vbTab & !BOBALANCE & vbTab & vbTab & !branchName & vbTab & !SPAREPARTID & vbTab & !BACKORDERID & vbTab & !branchID & vbTab & !SUPPORDERREF & vbTab & !SUPPOSTATUS & vbTab & !POQTY & vbTab & Format(!Cost, "#,###,##0.00")
                            .MoveNext
                            pBar.Value = pBar.Value + 1
                        Loop
                End If
            .Close
        End With
       
        pBar.Visible = False
        
End Sub

Private Sub cmdRefresh_Click()
    Dim rsRecDtls As New ADODB.Recordset
    Dim X1 As Integer, X2 As Integer, recRow As Integer, reserveQty As Integer

            
            
            flxRecDtls.Rows = 1
            flxRecDtls.Rows = 2
            pBar.Value = 0
            pBar.Visible = True
            pBar.Max = flxBBO.Rows - 1
            For X1 = 1 To flxBBO.Rows - 1
                    Qry = "SELECT spitr.REFERENCE, spitr.TRANSDATE, spitr.RECEIVE, spitr.TRANSTYPE, spitr.SPAREPARTID " & _
                            "From spitr Where spitr.TRANSDATE >= '" & Format(dFROM.Value, "yyyy-mm-dd") & "' AND spitr.TRANSDATE <= '" & Format(dTO.Value, "yyyy-mm-dd") & "' AND " & _
                            "spitr.SPAREPARTID = " & Val(flxBBO.TextMatrix(X1, 15)) & " AND spitr.PACKINGLIST <> '' and (spitr.TRANSTYPE = 'RECEIVE' OR spitr.TRANSTYPE = 'RECTEMP')"
                
                        'reserveQty = ComputeQty(flxBBO.TextMatrix(X1, 15))
                        'flxBBO.TextMatrix(X1, 10) = reserveQty
                        rsRecDtls.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                            If rsRecDtls.RecordCount > 0 Then
                                
                                For recRow = 1 To rsRecDtls.RecordCount
                                    For X2 = flxRecDtls.Rows - 1 To flxRecDtls.Rows - 1
                                    
                                        If txtSPid.Text = flxBBO.TextMatrix(X1, 15) And txtRowNum.Text = flxBBO.TextMatrix(X1, 0) Then
                                            
                                            flxRecDtls.TextMatrix(X2, 1) = flxBBO.TextMatrix(X1, 0)
                                            flxRecDtls.TextMatrix(X2, 2) = Left(flxRecDtls.TextMatrix(X2, 2), 1) & "*" & rsRecDtls!REFERENCE
                                            flxRecDtls.TextMatrix(X2, 3) = Left(flxRecDtls.TextMatrix(X2, 3), 1) & "*" & rsRecDtls!TRANSDATE
                                            'flxRecDtls.TextMatrix(X2, 3) = rsRecDtls!TRANSTYPE
                                            flxRecDtls.TextMatrix(X2, 5) = flxBBO.TextMatrix(X1, 19)    'PO Status
                                            
                                            If rsRecDtls!TRANSTYPE = "RECEIVE" Then
                                                flxRecDtls.TextMatrix(X2, 6) = Val(Replace(flxBBO.TextMatrix(X2, 5), ",", "")) + IIf(Val(Replace(rsRecDtls!RECEIVE, ",", "")) = 0, 0, Val(Replace(rsRecDtls!RECEIVE, ",", "")))
                                                'flxRecDtls.TextMatrix(X2, 6) = 0
                                            ElseIf rsRecDtls!TRANSTYPE = "RECTEMP" Then
                                                flxRecDtls.TextMatrix(X2, 7) = Val(Replace(flxBBO.TextMatrix(X2, 6), ",", "")) + IIf(Val(Replace(rsRecDtls!RECEIVE, ",", "")) = 0, 0, Val(Replace(rsRecDtls!RECEIVE, ",", "")))
                                                'flxRecDtls.TextMatrix(X2, 5) = 0
                                            Else
                                                flxRecDtls.TextMatrix(X2, 6) = 0
                                                flxRecDtls.TextMatrix(X2, 7) = 0
                                            End If
                                            
                                            flxRecDtls.TextMatrix(X2, 9) = IIf(flxBBO.TextMatrix(X1, 9) = "", 0, flxBBO.TextMatrix(X1, 9))
                                            flxRecDtls.TextMatrix(X2, 10) = flxBBO.TextMatrix(X1, 10)
                                            flxRecDtls.TextMatrix(X2, 11) = IIf(flxBBO.TextMatrix(X1, 12) = "", 0, flxBBO.TextMatrix(X1, 12))
                                            flxRecDtls.TextMatrix(X2, 0) = rsRecDtls!SPAREPARTID
                                            If flxBBO.TextMatrix(X1, 15) = flxRecDtls.TextMatrix(X2, 0) Then    'POQTY
                                                flxRecDtls.TextMatrix(X2, 8) = IIf(flxBBO.TextMatrix(X1, 20) = "", 0, flxBBO.TextMatrix(X1, 20))
                                            End If
                                        Else
                                            flxRecDtls.TextMatrix(X2, 1) = flxBBO.TextMatrix(X1, 0)
                                            flxRecDtls.TextMatrix(X2, 2) = rsRecDtls!REFERENCE
                                            flxRecDtls.TextMatrix(X2, 3) = rsRecDtls!TRANSDATE
                                            flxRecDtls.TextMatrix(X2, 4) = rsRecDtls!TRANSTYPE
                                            flxRecDtls.TextMatrix(X2, 5) = flxBBO.TextMatrix(X1, 19)    'PO Status
                                            
                                            If rsRecDtls!TRANSTYPE = "RECEIVE" Then
                                                flxRecDtls.TextMatrix(X2, 6) = rsRecDtls!RECEIVE
                                                'flxRecDtls.TextMatrix(X2, 6) = 0
                                            ElseIf rsRecDtls!TRANSTYPE = "RECTEMP" Then
                                                flxRecDtls.TextMatrix(X2, 7) = rsRecDtls!RECEIVE
                                                'flxRecDtls.TextMatrix(X2, 5) = 0
                                            Else
                                                flxRecDtls.TextMatrix(X2, 8) = 0
                                                flxRecDtls.TextMatrix(X2, 9) = 0
                                            End If
                                            
                                            flxRecDtls.TextMatrix(X2, 9) = IIf(flxBBO.TextMatrix(X1, 9) = "", 0, flxBBO.TextMatrix(X1, 9))
                                            flxRecDtls.TextMatrix(X2, 10) = flxBBO.TextMatrix(X1, 10)
                                            flxRecDtls.TextMatrix(X2, 11) = IIf(flxBBO.TextMatrix(X1, 12) = "", 0, flxBBO.TextMatrix(X1, 12))
                                            flxRecDtls.TextMatrix(X2, 0) = rsRecDtls!SPAREPARTID
                                            If flxBBO.TextMatrix(X1, 15) = flxRecDtls.TextMatrix(X2, 0) Then    'POQTY
                                                flxRecDtls.TextMatrix(X2, 8) = IIf(flxBBO.TextMatrix(X1, 20) = "", 0, flxBBO.TextMatrix(X1, 20))
                                            End If
                                        End If
                                        rsRecDtls.MoveNext
                                        txtSPid.Text = flxRecDtls.TextMatrix(X2, 0)
                                        txtRowNum.Text = flxRecDtls.TextMatrix(X2, 1)
                                    Next X2
                                    
                                Next recRow
                                flxRecDtls.Rows = flxRecDtls.Rows + 1
                            End If
                        rsRecDtls.Close
            pBar.Value = pBar.Value + 1
            Next X1
            Call sortSPID
            Call ComputeAVLQ
            Call recDtls
            pBar.Visible = False
End Sub

Public Sub sortSPID()
    
    With flxRecDtls
        .Col = 0
        .ColSel = 1
        .Sort = flexSortGenericAscending
    End With
End Sub


Public Sub ComputeAVLQ() ' available quantity to allocate
    Dim i As Integer
    txtSPid.Text = ""
    txtBackOrder.Text = ""
    With flxRecDtls
        For i = 1 To .Rows - 1
            If txtSPid.Text = .TextMatrix(i, 0) Then
                .TextMatrix(i, 12) = (Val(.TextMatrix(i, 7)) + Val(.TextMatrix(i, 9))) - Val(.TextMatrix(i, 10))
                .TextMatrix(i, 12) = Val(.TextMatrix(i, 12)) - Val(txtBackOrder.Text)
            Else
                .TextMatrix(i, 12) = (Val(.TextMatrix(i, 7)) + Val(.TextMatrix(i, 9))) - Val(.TextMatrix(i, 10))
            End If
            
            If txtSPid.Text = .TextMatrix(i, 0) Then
                txtBackOrder.Text = Val(txtBackOrder.Text) + Val(.TextMatrix(i, 11))
            Else
                txtBackOrder.Text = .TextMatrix(i, 11)
            End If
            txtSPid.Text = .TextMatrix(i, 0)
        Next i
    End With
    
End Sub


Public Sub RevAvlZero()
    
    With flxBBO
        For i = .Rows - 1 To 1 Step -1
            If Val(.TextMatrix(i, 11)) = 0 Then
                .RemoveItem i
            End If
        Next i
    End With
    
End Sub

Public Sub BranchOnly()
    
    With flxBBO
        For i = .Rows - 1 To 1 Step -1
            If Val(.TextMatrix(i, 17)) <> Val(txtBranchID.Text) Then
                .RemoveItem i
            End If
        Next i
    End With
    
End Sub

Function ComputeQty(mySPID As Long) As Long
    
    Dim rsALQ As New ADODB.Recordset
    Dim sumALQ As Long, transPending As Long
    
    Qry = "SELECT SUM(spbranchdtl.ALLOCATEDQTY) As SUMofALQ From spbranchdtl " & _
            "Where spbranchdtl.TRANSTYPE = 'ORDER TEMP' AND spbranchdtl.REMARKS = 'PACKING'  AND spbranchdtl.SPAREPARTID = " & mySPID & ";"
    With rsALQ
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            sumALQ = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            sumALQ = 0
        End If
        .Close
    End With
    
    
    Qry = "SELECT SUM(stocktransdtl.Qty) As SumOfPending From stocktransfer " & _
            "INNER JOIN stocktransdtl ON (stocktransfer.TRANSFERID = stocktransdtl.TRANSFERID) " & _
            "Where stocktransfer.`STATUS` = 'PENDING' AND stocktransdtl.SPAREPARTID = " & mySPID & ";"
    
    With rsALQ
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            transPending = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            transPending = 0
        End If
        .Close
    End With
    
    ComputeQty = Val(transPending) + Val(sumALQ)
    
End Function


Private Sub RXResizer1_ResizeComplete()

    With flxBBO
        '                0        1           2          3        4        5          6         7    8     9    10   11    12   13     14        15
        .FormatString = "^ |<PARTNUMBER|<DESCRIPTION|<REC.REF#|^REC.DATE|^P.O. #|^P.O. DATE|^Q.REC|^T.REC|^QOH|^RST|^AVLQ|^BOQ|^ALQ|<BRANCH"    'spid   'boid   'branchid
        .ColWidth(0) = 500
        .ColWidth(1) = (.Width - 300) * 0.12
        .ColWidth(2) = (.Width - 300) * 0.2
        .ColWidth(3) = (.Width - 300) * 0.07
        .ColWidth(4) = (.Width - 300) * 0.06
        .ColWidth(5) = (.Width - 300) * 0.07
        .ColWidth(6) = (.Width - 300) * 0.07
        .ColWidth(7) = (.Width - 300) * 0.04
        .ColWidth(8) = (.Width - 300) * 0.04
        .ColWidth(9) = (.Width - 300) * 0.04
        .ColWidth(10) = (.Width - 300) * 0.04
        .ColWidth(11) = (.Width - 300) * 0.04
        .ColWidth(12) = (.Width - 300) * 0.04
        .ColWidth(13) = (.Width - 300) * 0.04
        .ColWidth(14) = (.Width - 300) * 0.1
        .ColWidth(15) = 0       'SPID
        .ColWidth(16) = 0       'BACKORDERID
        .ColWidth(17) = 0       'BRANCHID
        .ColWidth(18) = 0       'SUPPLIER PO REF
        .ColWidth(19) = 0       'SUP PO STATUS
        .ColWidth(20) = 0       'SUP PO QTY
    
    End With

End Sub

Private Sub Timer2_Timer()
    Dim rsBranch As New ADODB.Recordset
    
        With rsBranch
            Qry = "SELECT DISTINCT backorder.BRANCHID, backorder.BRANCHNAME From backorder Group By BRANCHNAME Order By backorder.BRANCHNAME;"
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        cboBranch.AddItem IIf(IsNull(!branchName), "", !branchName)
                        cboBranch.ItemData(cboBranch.ListCount - 1) = IIf(IsNull(!branchID), 0, !branchID)
                        .MoveNext
                    Loop
                End If
            .Close
       End With
       
    Timer2.Enabled = False
End Sub

Private Sub txtBBOaloc_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        txtBBOaloc.Text = Empty
        flxBBO.TextMatrix(ActRow, 13) = txtBBOaloc.Tag
        txtBBOaloc.Tag = "skiplostfocus"
        txtBBOaloc.Visible = False
        flxBBO.SetFocus
        flxBBO.Col = 13
    End If
    
    
End Sub

Private Sub txtBBOaloc_KeyPress(KeyAscii As Integer)
    Dim numOnly As String
    
        numOnly = "1234567890"
        
        If KeyAscii > 26 Then
            If InStr(numOnly, Chr(KeyAscii)) = 0 Then KeyAscii = 0
        End If
        If KeyAscii = 13 Then
            flxBBO.SetFocus
            flxBBO.Col = 13
        End If
            
End Sub

Private Sub txtBBOaloc_LostFocus()
         
         If txtBBOaloc.Tag = "skiplostfocus" Then
             txtBBOaloc.Tag = ""
             Exit Sub
         End If
         If Val(txtBBOaloc.Text) > Val(flxBBO.TextMatrix(ActRow, 11)) Then
             MsgBox ("The Allocated Quantity greater than Available is Not Allowed!"), vbExclamation, "Not Allowed"
             txtBBOaloc.SetFocus
             Exit Sub
         ElseIf Val(txtBBOaloc.Text) > Val(flxBBO.TextMatrix(ActRow, 12)) Then
             MsgBox ("The Allocated Quantity greater than Backorder is Not Allowed!"), vbExclamation, "Not Allowed"
             txtBBOaloc.SetFocus
             Exit Sub
         End If
        
             flxBBO.TextMatrix(ActRow, 13) = txtBBOaloc.Text
             txtBBOaloc.Visible = False
        
    
End Sub

Public Sub empCell()
    Dim emp As Integer
    With flxBBO
        For emp = 1 To flxBBO.Rows - 1
            .TextMatrix(emp, 3) = ""
            .TextMatrix(emp, 4) = ""
            .TextMatrix(emp, 7) = ""
            .TextMatrix(emp, 8) = ""
        Next emp
    End With
End Sub
