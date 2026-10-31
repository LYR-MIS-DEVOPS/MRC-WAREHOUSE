VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form BranchPicking 
   Caption         =   "Picking List"
   ClientHeight    =   10320
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10320
   ScaleWidth      =   19200
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtPMID 
      Height          =   285
      Left            =   19200
      TabIndex        =   30
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtBCODE 
      Height          =   285
      Left            =   19200
      TabIndex        =   29
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtPrnCopy 
      Height          =   285
      Left            =   19320
      TabIndex        =   28
      Top             =   6120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   0
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
      Image           =   "BranchPicking.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Tag             =   "XYF"
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "Print Preview"
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
      Image           =   "BranchPicking.frx":059A
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Tag             =   "XYF"
      Top             =   9360
      Width           =   18975
      Begin VB.TextBox txtFINISHPICK 
         Height          =   285
         Left            =   17040
         TabIndex        =   21
         Tag             =   "XY"
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtSTARTPICK 
         Height          =   285
         Left            =   13440
         TabIndex        =   19
         Tag             =   "XY"
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtPROCSBY 
         Height          =   285
         Left            =   1800
         TabIndex        =   17
         Tag             =   "XY"
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label lblEnd 
         Alignment       =   1  'Right Justify
         Caption         =   "TIME FINISH PICKING"
         Height          =   255
         Left            =   15240
         TabIndex        =   20
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblStart 
         Alignment       =   1  'Right Justify
         Caption         =   "TIME START PICKING"
         Height          =   255
         Left            =   11520
         TabIndex        =   18
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblProcBy 
         Alignment       =   1  'Right Justify
         Caption         =   "PROCESSED BY"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7215
      Left            =   120
      TabIndex        =   13
      Tag             =   "XY"
      Top             =   2160
      Width           =   18975
      Begin MSFlexGridLib.MSFlexGrid flxPickList 
         Height          =   4695
         Left            =   120
         TabIndex        =   14
         Tag             =   "XY"
         Top             =   240
         Width           =   18735
         _ExtentX        =   33046
         _ExtentY        =   8281
         _Version        =   393216
         Cols            =   9
         AllowUserResizing=   1
         FormatString    =   ""
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
      Begin MSFlexGridLib.MSFlexGrid flxPickSumm 
         Height          =   1935
         Left            =   120
         TabIndex        =   25
         Tag             =   "XY"
         Top             =   5160
         Width           =   18735
         _ExtentX        =   33046
         _ExtentY        =   3413
         _Version        =   393216
         Cols            =   8
         SelectionMode   =   2
         FormatString    =   ""
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
      Begin VB.Label lblSummart 
         Caption         =   "SUMMARY"
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
         TabIndex        =   27
         Top             =   4920
         Width           =   1095
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   19320
      TabIndex        =   23
      Top             =   480
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
      Image           =   "BranchPicking.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   19320
      TabIndex        =   24
      Top             =   960
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
      Image           =   "BranchPicking.frx":10CE
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Tag             =   "XY"
      Top             =   480
      Width           =   18975
      Begin VB.ComboBox cmboRefno 
         Height          =   315
         Left            =   1080
         TabIndex        =   26
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtTOTALSRQ 
         Height          =   285
         Left            =   16800
         TabIndex        =   9
         Tag             =   "XY"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtTOTALAVQ 
         Height          =   285
         Left            =   16800
         TabIndex        =   8
         Tag             =   "XY"
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtTOTALPN 
         Height          =   285
         Left            =   16800
         TabIndex        =   7
         Tag             =   "XY"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtDATE 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Tag             =   "XY"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.ComboBox cmboBRANCH 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Tag             =   "XY"
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label lblTOTALSRQ 
         Caption         =   "TOTAL SRQ:"
         Height          =   255
         Left            =   15720
         TabIndex        =   12
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblTOTALAVQ 
         Caption         =   "TOTAL AVQ:"
         Height          =   255
         Left            =   15720
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblTOTALPN 
         Caption         =   "TOTAL PN:"
         Height          =   255
         Left            =   15720
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblDATE 
         Caption         =   "DATE:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblREFNO 
         Caption         =   "REF NO:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblBranch 
         Caption         =   "BRANCH:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   9720
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   19320
      m_MDI_Height    =   10830
   End
   Begin VB.Menu mnuPrint 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuPICKRepl 
         Caption         =   "PICK BRANCH REPLENISHMENT"
      End
      Begin VB.Menu mnuFINALRepl 
         Caption         =   "FINAL BRANCH REPLENISHMENT"
      End
   End
End
Attribute VB_Name = "BranchPicking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrREFNO() As String
Dim arrBRANCH() As String

Dim Ans As String
Dim branchCode As String
Dim bREFID As Integer
Dim formAction As String
Dim Tmp() As String
Dim MyErr As String, myMsg As String
Dim Qry As String, Qry2 As String

Dim matrixID As Integer
Dim strREFID As String

Dim i As Long, j As Long, z As Long, X As Long, c As Long
Dim ActRow As Integer, ActCol As Integer, intRow As Integer

Dim tmpRs  As New ADODB.Recordset
Dim mainRs As New ADODB.Recordset

Dim strPartial As String, strTotal As String
Dim StatSrch As Boolean, AddNew As Boolean, ProperExit As Boolean

Public mySrc As String

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
    Call cmboBranch_Change

    branchCode = arrBRANCH(cmboBranch.ListIndex, 0)
    
    Call cmboREFNODisplay
    txtBCODE.Text = branchCode 'used in frmPrn
    cmboRefno.SetFocus
End Sub

Private Sub cmboBranch_GotFocus()
Dim Tmp
    Const CB_SHOWDROPDOWN = &H14F
    Tmp = SendMessage(cmboBranch.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cmboBRANCH_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cmboBranch, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Call cmboBRANCH_Click
        cmboRefno.SetFocus
    End If
End Sub

Private Sub cmboRefno_Change()
On Error GoTo er
    If StatSrch Then
            StatSrch = False
            Exit Sub
    End If

    With cmboRefno
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmboBRAND_Change()")
End Sub

Private Sub cmboRefno_Click()
    Call cmboRefno_Change
    
    matrixID = arrREFNO(cmboRefno.ListIndex, 0)
    txtpmID.Text = matrixID
    Call ShowRecord(branchCode, cmboRefno.Text, matrixID)
End Sub

Private Sub cmboRefno_GotFocus()
Dim Tmp
    Const CB_SHOWDROPDOWN = &H14F
    Tmp = SendMessage(cmboRefno.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cmboRefno_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cmboRefno, KeyAscii, ValidChars)
    
    If KeyAscii = 13 Then
        
    End If
End Sub


Public Sub cmboDisplay()
On Error GoTo er
    'Get Data from database
    DoEvents
    Call modPMBranchIdentifying.sp_get_Branch(tmpRs)
    If tmpRs.State = 1 Then
        Call modObjTools.addToListObjectWithArrayGivenField(cmboBranch, tmpRs, arrBRANCH, "BRANCHCODE", "BRANCHNAME")
    End If
    If tmpRs.State <> 0 Then tmpRs.Close

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub

Public Sub cmboREFNODisplay()
On Error GoTo er
    'Get Data from database
    DoEvents
    Call modPMBranchPicking.sp_getRefNo(tmpRs)
    cmboRefno.Clear
    If tmpRs.State = 1 Then
        Call modObjTools.addToListObjectWithArrayGivenField(cmboRefno, tmpRs, arrREFNO, "PMATRIXID", "REFNO")
    End If
    If tmpRs.State <> 0 Then tmpRs.Close
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmboREFNODisplay()")
End Sub

Private Sub cmdTrans_Click(Index As Integer)

If Index = 0 Then 'Print
    If txtPrnCopy.Text = "" Then
        MsgBox "No available Pick List yet", vbInformation + vbOKOnly, "PICK LIST"
        Exit Sub
    ElseIf Val(txtPrnCopy.Text) = 0 Then
        mnuPICKRepl.Enabled = False
        mnuFINALRepl.Enabled = False
    ElseIf Val(txtPrnCopy.Text) > 0 Then
        mnuPICKRepl.Enabled = False
        mnuFINALRepl.Enabled = True
    End If
    
    Call PopupMenu(mnuPrint)
    
ElseIf Index = 1 Then 'Exit
    formAction = "EXIT"
    ProperExit = True
    'Call clearform
    Unload Me
End If

End Sub

Private Sub Form_Load()
  '  Call modCon.Connection
    Call cmboDisplay
    txtProcsBy.Text = EmployeesName
    txtDATE.Text = Format(Date, "mm/dd/yyyy")
    txtTotalAVQ.Text = 0
    txtTOTALPN.Text = 0
    txtTotalSRQ.Text = 0
    txtSTARTPICK.Text = Format$(Now, "hh:mm AM/PM")
    txtFINISHPICK.Text = Format$(Now, "hh:mm AM/PM")
    
    '===================BUTTON POSITION===================
    
'    If cmdTrans(0).Visible = False Then
'            'Align Save Button
'            cmdTrans(4).Left = cmdTrans(0).Left
'            cmdTrans(4).Top = cmdTrans(0).Top
'            'Align Cancel Button
'            cmdTrans(5).Left = cmdTrans(1).Left
'            cmdTrans(5).Top = cmdTrans(1).Top
'
'    End If
End Sub

Private Sub Form_Resize()
Call RXResizer1_ResizeComplete
End Sub



Private Sub mnuFINALRepl_Click()
    If Val(txtpmID.Text) = 0 Then
        MsgBox "No Records Found", vbExclamation + vbOKOnly, "PRINT"
        Exit Sub
    End If
    
    Qry = "call sp_partsmatrix.sp_brnchPicking_PickingListDtls('" & sysBranchCode & "','" & cmboRefno.Text & "'," & matrixID & ",'" & branchCode & "')"
    Qry2 = "call sp_partsmatrix.sp_brnchPicking_PickingListTtl('" & sysBranchCode & "','" & cmboRefno.Text & "'," & matrixID & ",'" & branchCode & "')"
    frmPrn.QryPrn = Qry
    frmPrn.QryPrn2 = Qry2
    
    frmPrn.ParamValue1 = SysBranch
    frmPrn.ParamValue2 = Format(SysAddress, "")
    frmPrn.ParamValue3 = Format(SysTel, "")
    frmPrn.ParamValue4 = Format(cmboBranch, "")
    frmPrn.ParamValue5 = Format(cmboRefno, "")
    frmPrn.ParamValue34 = Format(Date, "mm-dd-yyyy")
    frmPrn.ParamValue36 = Format(txtTOTALPN, "")
    frmPrn.ParamValue37 = Format(txtTotalAVQ, "")
    frmPrn.ParamValue38 = Format(txtTotalSRQ, "")
    frmPrn.ParamValue9 = Format(txtProcsBy, "")
    frmPrn.ParamValue39 = Format(txtSTARTPICK, "")
    frmPrn.ParamValue40 = Format(txtFINISHPICK, "")
    frmPrn.MySource = "Branch Picking"
    frmPrn.Show
End Sub

Private Sub mnuPICKRepl_Click()
    If Val(txtpmID.Text) = 0 Then
        MsgBox "No Records Found", vbExclamation + vbOKOnly, "PRINT"
        Exit Sub
    End If

    Qry = "call sp_partsmatrix.sp_brnchPicking_PickingListDtls('" & sysBranchCode & "','" & cmboRefno.Text & "'," & matrixID & ",'" & branchCode & "')"
    Qry2 = "call sp_partsmatrix.sp_brnchPicking_PickingListTtl('" & sysBranchCode & "','" & cmboRefno.Text & "'," & matrixID & ",'" & branchCode & "')"
    frmPrn.QryPrn = Qry
    frmPrn.QryPrn2 = Qry2
    
    frmPrn.ParamValue1 = SysBranch
    frmPrn.ParamValue2 = Format(SysAddress, "")
    frmPrn.ParamValue3 = Format(SysTel, "")
    frmPrn.ParamValue4 = Format(cmboBranch, "")
    frmPrn.ParamValue5 = Format(cmboRefno, "")
    frmPrn.ParamValue34 = Format(Date, "mm-dd-yyyy")
    frmPrn.ParamValue36 = Format(txtTOTALPN, "")
    frmPrn.ParamValue37 = Format(txtTotalAVQ, "")
    frmPrn.ParamValue38 = Format(txtTotalSRQ, "")
    frmPrn.ParamValue9 = Format(txtProcsBy, "")
    frmPrn.ParamValue39 = Format(txtSTARTPICK, "")
    frmPrn.ParamValue40 = Format(txtFINISHPICK, "")
    frmPrn.MySource = "Branch Picking"
    frmPrn.Show
End Sub

Private Sub RXResizer1_ResizeComplete()
    Call flxColChanger
    Call flxColChangerForSumm
    
    'Frame width
    Frame1.Width = Me.Width - 465
    Frame2.Width = Me.Width - 465
    Frame3.Width = Me.Width - 465
  
    'Frame1 label Left
    lblBranch.Left = 240
    lblRefNo.Left = 240
    lblDATE.Left = 240
    
    lblTOTALPN.Left = Frame1.Width - 3255
    lblTOTALSRQ.Left = Frame1.Width - 3255
    lblTOTALAVQ.Left = Frame1.Width - 3255

    'Frame1 Label Width
    lblBranch.Width = 735
    lblTOTALPN.Width = 855
    lblTOTALSRQ.Width = 855
    lblTOTALAVQ.Width = 855
    lblRefNo.Width = 735
    lblDATE.Width = 735
    
    'Frame1 Combobox Left
    cmboBranch.Left = 1200
    cmboRefno.Left = 1200
    
    'Frame1 Combobox Width
    cmboBranch.Width = 4575
    cmboRefno.Width = 1815
    
    'Frame1 Text Left
    txtTOTALPN.Left = Frame1.Width - 2175
    txtTotalSRQ.Left = Frame1.Width - 2175
    txtTotalAVQ.Left = Frame1.Width - 2175
    txtDATE.Left = 1200
    
    'Frame1 Text Width
    txtTOTALPN.Width = 1815
    txtTotalSRQ.Width = 1815
    txtTotalAVQ.Width = 1815
    txtDATE.Width = 1815
    
    'Frame4 Label Left
    lblProcBy.Left = 240
    lblStart.Left = Frame3.Width - 7455
    lblEnd.Left = Frame3.Width - 3735
    
    'Frame4 Label Width
    lblProcBy.Width = 1455
    lblStart.Width = 1815
    lblEnd.Width = 1695
    
    'Frame4 Text Left
    txtProcsBy.Left = 1800
    txtSTARTPICK.Left = Frame3.Width - 5535
    txtFINISHPICK.Left = Frame3.Width - 1935
    
    'Frame4 Text Width
    txtProcsBy.Width = 4335
    txtSTARTPICK.Width = 1695
    txtFINISHPICK.Width = 1695
    
End Sub

Public Sub flxColChanger()

With flxPickList
'    'MERGE COL
'    .TextMatrix(0, 0) = " "
'    .TextMatrix(1, 0) = " "
'    'MERGE ROW
'    .TextMatrix(1, 1) = "MODEL BRAND (PM)"
'    .TextMatrix(1, 2) = "MODEL BRAND (PM)"
'    .TextMatrix(1, 3) = "MODEL BRAND (PM)"
'    .TextMatrix(1, 4) = "MODEL BRAND (PM)"
'    .TextMatrix(1, 5) = "MODEL BRAND (PM)"
'    .TextMatrix(1, 6) = "MODEL BRAND (PM)"
'    .TextMatrix(1, 7) = "MODEL BRAND (PM)"
'    .TextMatrix(1, 8) = "MODEL BRAND (PM)"
    
    'MERGE COL =FALSE
    .TextMatrix(0, 1) = "NO."
    .TextMatrix(0, 2) = "PARTNUMBER"
    .TextMatrix(0, 3) = "DESCRIPTION"
    .TextMatrix(0, 4) = "QH"
    .TextMatrix(0, 5) = "AVQ"
    .TextMatrix(0, 6) = "SRQ"
    .TextMatrix(0, 7) = "ALQ"
    .TextMatrix(0, 8) = "LOCATION"

'    'SET MERGING
'    .MergeCells = 1
'    .MergeRow(1) = True
'    .MergeCol(1) = True
'    .MergeCol(1) = True
'    For i = 1 To 3
'        .MergeCol(i) = False
'    Next i
'    'SET THE COLUMN COLOR GREEN
'    .Row = 0
'    .Col = 0
'    .Col = 1
End With
    
    Call flxWidth
    
End Sub

Public Sub flxColChangerForSumm()

With flxPickSumm
    'MERGE COL
    .TextMatrix(0, 0) = " "
    .TextMatrix(1, 0) = " "
    'MERGE ROW
'    .TextMatrix(0, 1) = "SUMMARY"
'    .TextMatrix(0, 2) = "SUMMARY"
'    .TextMatrix(0, 3) = "SUMMARY"
'    .TextMatrix(0, 4) = "SUMMARY"
'    .TextMatrix(0, 5) = "SUMMARY"
'    .TextMatrix(0, 6) = "SUMMARY"
'    .TextMatrix(0, 7) = "SUMMARY"
    '.TextMatrix(1, 8) = "MODEL BRAND (PM)"
    
    'MERGE COL =FALSE
    .TextMatrix(0, 1) = "NO."
    .TextMatrix(0, 2) = "MODEL BRAND (PM)"
    .TextMatrix(0, 3) = "TTL PN"
    .TextMatrix(0, 4) = "TTL SRQ"
    .TextMatrix(0, 5) = "TTL AVQ"
    .TextMatrix(0, 6) = "TTL ALQ"
    .TextMatrix(0, 7) = "EFFICIENCY"
    
'    'SET MERGING
'    .MergeCells = 1
'    .MergeRow(0) = True
'    .MergeCol(1) = True
'    .MergeCol(1) = True
'    For i = 1 To 3
'        .MergeCol(i) = False
'    Next i
'    'SET THE COLUMN COLOR GREEN
'    .Row = 0
'    .Col = 0
'    .Col = 1
End With
    
    Call flxWidthForSumm
    
End Sub



Public Sub flxWidth()

    With flxPickList
        '.Clear
        .Cols = 10
       
        .FormatString = "|<NO |<PART NUMBER |<DESCRIPTION |^QH |^AVQ |^SRQ |^ALQ |<LOCATION"
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.03  'NO
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.25 'PARTNUMBER
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.35 'DESCRIPTION
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.06 'QH
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.06 'AVQ
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.06 'SRQ
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.06 'ALQ
        .ColWidth(8) = (.Width - .ColWidth(0)) * 0.1 'LOCATION
        .ColWidth(9) = 0 '(.Width - .ColWidth(0)) * 0.1 'LOCATION
    End With
End Sub

Public Sub flxWidthForSumm()

    With flxPickSumm
        '.Clear
        .Cols = 8
        '.ColAlignment(1) = flexAlignLeftTop
        .FormatString = "|<NO |<MODEL BRAND (PM) |^TOTAL PN |^TOTAL SRQ |^TOTAL AVQ |^TOTAL ALQ |^EFFICIENCY"
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.03  'NO
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.2 'MODEL BRAND (PM)
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.1 'TTL PN
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.1 'TTL SRQ
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.1 'TTL AVQ
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.1 'TOTAL ALQ
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.1 'EFFICIENCY
       ' .ColWidth(8) = 0 '(.Width - .ColWidth(0)) * 0.06 'ID

    End With
    
End Sub

Public Sub ShowRecord(MyBranchCode As String, myREFNO As String, MyID As Integer)

  Screen.MousePointer = vbHourglass
  
  DoEvents
 Call modPMBranchPicking.sp_brnchPicking_PickingListDtls(myREFNO, MyID, MyBranchCode, tmpRs)
 
    If tmpRs.State <> 1 Then Exit Sub
    
    If tmpRs.RecordCount <= 0 Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    With tmpRs
        If .RecordCount > 0 Then
            .MoveFirst
           ' intTtlPN = .RecordCount
            flxPickList.Rows = .RecordCount + 2
            txtPrnCopy.Text = !PRNCOPY
           j = 0
            For i = 2 To flxPickList.Rows - 1
                j = j + 1
                flxPickList.TextMatrix(i, 1) = j
                flxPickList.TextMatrix(i, 2) = !PartNumber
                flxPickList.TextMatrix(i, 3) = !Description
                flxPickList.TextMatrix(i, 4) = !QtyOnHand
                flxPickList.TextMatrix(i, 5) = !AVQ
                flxPickList.TextMatrix(i, 6) = !PMLACKING
                flxPickList.TextMatrix(i, 7) = !ALLOCATEDQTY
                flxPickList.TextMatrix(i, 8) = !Location
                flxPickList.TextMatrix(i, 9) = !MODELBRAND
                .MoveNext
            Next i
            
            txtTOTALPN.Text = flxPickList.Rows - 2
            
            With flxPickList
        
                .TextMatrix(1, 1) = flxPickList.TextMatrix(2, 9)
                .TextMatrix(1, 2) = flxPickList.TextMatrix(2, 9)
                .TextMatrix(1, 3) = flxPickList.TextMatrix(2, 9)
                .TextMatrix(1, 4) = flxPickList.TextMatrix(2, 9)
                .TextMatrix(1, 5) = flxPickList.TextMatrix(2, 9)
                .TextMatrix(1, 6) = flxPickList.TextMatrix(2, 9)
                .TextMatrix(1, 7) = flxPickList.TextMatrix(2, 9)
                .TextMatrix(1, 8) = flxPickList.TextMatrix(2, 9)
                'SET MERGING
                .MergeCells = 1
                .MergeRow(1) = True
                .MergeCol(1) = True
                .MergeCol(1) = True
                For j = 1 To 9
                .MergeCol(j) = False
                Next j
            End With
          
            
            
            
            For i = 2 To flxPickList.Rows - 2
                If flxPickList.TextMatrix(i, 9) <> flxPickList.TextMatrix(i + 1, 9) Then
                    
                    flxPickList.Rows = flxPickList.Rows + 1
                        For z = flxPickList.Rows - 1 To i Step -1
                            If z = i Then
                              
                                For X = 1 To flxPickList.Cols - 1
                                flxPickList.TextMatrix(z + 1, X) = flxPickList.TextMatrix(i + 2, 9)
                                    'SET MERGING
                                    
                                    flxPickList.MergeCells = 1
                                    flxPickList.MergeRow(z + 1) = True
                                    flxPickList.MergeCol(0) = True
                                    flxPickList.MergeCol(0) = True
                                    For j = 1 To 9
                                    flxPickList.MergeCol(j) = False
                                    Next j
                                    
                                Next X
                           
                            Else
                                For X = 1 To flxPickList.Cols - 1
                                flxPickList.TextMatrix(z, X) = flxPickList.TextMatrix(z - 1, X)
                                Next X
                            End If
                        Next z
                End If
            Next i
        End If
    End With
    If tmpRs.State <> 0 Then
        Screen.MousePointer = vbDefault
        tmpRs.Close
    End If
    
    
    Call modPMBranchPicking.sp_brnchPicking_PickingListTtl(myREFNO, MyID, MyBranchCode, mainRs)
    If mainRs.State <> 1 Then Exit Sub
    
    If mainRs.RecordCount <= 0 Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    With mainRs
     If .RecordCount > 0 Then
            .MoveFirst
           ' intTtlPN = .RecordCount
            flxPickSumm.Rows = .RecordCount + 1
            txtSTARTPICK.Text = Format$(!TIMESTART, "")
            txtFINISHPICK.Text = Format$(!TIMEFINISH, "")
            j = 0
            For i = 1 To flxPickSumm.Rows - 1
                j = j + 1
                flxPickSumm.TextMatrix(i, 1) = j
                flxPickSumm.TextMatrix(i, 2) = !MODELBRAND
                flxPickSumm.TextMatrix(i, 3) = !TOTALPN
                flxPickSumm.TextMatrix(i, 4) = !TOTALSRQ
                flxPickSumm.TextMatrix(i, 5) = !TOTALAVQ
                flxPickSumm.TextMatrix(i, 6) = !TOTALALQ
                
                .MoveNext
            Next i
    End If
    End With
    If mainRs.State <> 0 Then
        Screen.MousePointer = vbDefault
        mainRs.Close
    End If

    
    txtTotalAVQ.Text = 0
    For i = 1 To flxPickSumm.Rows - 1
        txtTotalAVQ.Text = Val(txtTotalAVQ) + Val(flxPickSumm.TextMatrix(i, 5))
    Next i
    
    txtTotalSRQ.Text = 0
    For i = 1 To flxPickSumm.Rows - 1
        txtTotalSRQ.Text = Val(txtTotalSRQ) + Val(flxPickSumm.TextMatrix(i, 4))
    Next i
   
End Sub




Private Sub txtFINISHPICK_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFINISHPICK, KeyAscii, ValidChars)
    
    If KeyAscii = 13 Then
        txtFINISHPICK.SetFocus
    End If
End Sub

Private Sub txtPROCSBY_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtProcsBy, KeyAscii, ValidChars)
    
    If KeyAscii = 13 Then
        txtSTARTPICK.SetFocus
    End If
End Sub



Private Sub txtSTARTPICK_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtSTARTPICK, KeyAscii, ValidChars)
    
    If KeyAscii = 13 Then
        txtFINISHPICK.SetFocus
    End If
End Sub
