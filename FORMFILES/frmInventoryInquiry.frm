VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInventoryInquiry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Retail Inventory Inquiry"
   ClientHeight    =   9510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12555
   Icon            =   "frmInventoryInquiry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9510
   ScaleWidth      =   12555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "Branch Database Update"
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Tag             =   "XY"
      Top             =   8280
      Width           =   12375
      Begin VB.Label lblBrnCon 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Suggested Order Quantity Formula"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   195
         TabIndex        =   9
         Tag             =   "XY"
         Top             =   240
         Width           =   11940
      End
      Begin VB.Label lblBrnCon2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Suggested Order Quantity Formula"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   195
         TabIndex        =   8
         Tag             =   "XY"
         Top             =   480
         Width           =   11940
      End
      Begin VB.Label lblBrnCon3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Suggested Order Quantity Formula"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   195
         TabIndex        =   7
         Tag             =   "XY"
         Top             =   720
         Width           =   11940
      End
   End
   Begin VB.Timer tmrAutoLoad 
      Interval        =   10
      Left            =   8400
      Top             =   0
   End
   Begin VB.Frame Frame2 
      Height          =   7455
      Left            =   120
      TabIndex        =   5
      Tag             =   "XY"
      Top             =   720
      Width           =   12375
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxParts 
         Height          =   7095
         Left            =   120
         TabIndex        =   10
         Tag             =   "xyf"
         Top             =   240
         Width           =   12075
         _ExtentX        =   21299
         _ExtentY        =   12515
         _Version        =   393216
         BackColor       =   16777215
         RowHeightMin    =   300
         FocusRect       =   2
         FillStyle       =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Tag             =   "XY"
      Top             =   120
      Width           =   12375
      Begin VB.ComboBox cboFind 
         Height          =   315
         Left            =   1980
         TabIndex        =   1
         Top             =   120
         Width           =   4335
      End
      Begin MSComCtl2.UpDown UpDRef 
         Height          =   315
         Left            =   1740
         TabIndex        =   2
         Top             =   120
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Max             =   1
         Enabled         =   -1  'True
      End
      Begin lvButton.lvButtons_H cmdEvaluate 
         Height          =   375
         Left            =   6420
         TabIndex        =   3
         Top             =   120
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   661
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
         ImgAlign        =   4
         Image           =   "frmInventoryInquiry.frx":57E2
         cBack           =   -2147483633
      End
      Begin VB.Label lblRefSearch 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FIND PARTS DESC."
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   195
         Width           =   1485
      End
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   9000
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   12645
      m_MDI_Height    =   9930
   End
End
Attribute VB_Name = "frmInventoryInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mainRs As New ADODB.Recordset

Dim ActRow As Long, ActCol As Long

Public Sub LoadBackUp()
Dim Qry As String

lblBrnCon.Caption = ""
lblBrnCon2.Caption = ""
lblBrnCon3.Caption = ""
Qry = "CALL sp_mrc_backup_monitoring.sp_mrc_backup_monitoring;"
    If mainRs.State = adStateOpen Then mainRs.Close
    With mainRs
        DoEvents
        .Open Qry, LYRcon, adOpenStatic
        If .RecordCount > 1 Then
            .MoveFirst
            Do While Not .EOF
                If !BranchCode = "MST" Then
                    lblBrnCon.Caption = IIf(lblBrnCon.Caption <> "", lblBrnCon.Caption & ", ", "") & "LYRMP SOBRECARY:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MRT" Then
                    lblBrnCon.Caption = IIf(lblBrnCon.Caption <> "", lblBrnCon.Caption & ", ", "") & "LYRMP GANTE:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MUD" Then
                    lblBrnCon.Caption = IIf(lblBrnCon.Caption <> "", lblBrnCon.Caption & ", ", "") & "LYRMP ULAS:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MG" Then
                    lblBrnCon.Caption = IIf(lblBrnCon.Caption <> "", lblBrnCon.Caption & ", ", "") & "LYRMP GENSAN:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MTS" Then
                    lblBrnCon.Caption = IIf(lblBrnCon.Caption <> "", lblBrnCon.Caption & ", ", "") & "LYRMP STO. TOMAS:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MP" Then
                    lblBrnCon.Caption = IIf(lblBrnCon.Caption <> "", lblBrnCon.Caption & ", ", "") & "LYRMP PANABO:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MMK" Then
                    lblBrnCon.Caption = IIf(lblBrnCon.Caption <> "", lblBrnCon.Caption & ", ", "") & "LYRMP MONKAYO:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MCD" Then
                    lblBrnCon2.Caption = IIf(lblBrnCon2.Caption <> "", lblBrnCon2.Caption & ", ", "") & "LYRMP CALINAN:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MPT" Then
                    lblBrnCon2.Caption = IIf(lblBrnCon2.Caption <> "", lblBrnCon2.Caption & ", ", "") & "MC TORIL:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MCPR" Then
                    lblBrnCon2.Caption = IIf(lblBrnCon2.Caption <> "", lblBrnCon2.Caption & ", ", "") & "MC RIZAL:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "LMSF" Then
                    lblBrnCon2.Caption = IIf(lblBrnCon2.Caption <> "", lblBrnCon2.Caption & ", ", "") & "LYRMP SANFRANS:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MLA" Then
                    lblBrnCon2.Caption = IIf(lblBrnCon2.Caption <> "", lblBrnCon2.Caption & ", ", "") & "LYRMP AGDAO:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "LMPC" Then
                    lblBrnCon2.Caption = IIf(lblBrnCon2.Caption <> "", lblBrnCon2.Caption & ", ", "") & "LYRMP COMPOSTELA:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "LMPD" Then
                    lblBrnCon2.Caption = IIf(lblBrnCon2.Caption <> "", lblBrnCon2.Caption & ", ", "") & "LYRMP DIGOS:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MWS" Then
                    lblBrnCon3.Caption = IIf(lblBrnCon3.Caption <> "", lblBrnCon3.Caption & ", ", "") & "LYRMP WHOLESALE:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MV1" Then
                    lblBrnCon3.Caption = IIf(lblBrnCon3.Caption <> "", lblBrnCon3.Caption & ", ", "") & "VAN-72:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MV2" Then
                    lblBrnCon3.Caption = IIf(lblBrnCon3.Caption <> "", lblBrnCon3.Caption & ", ", "") & "VAN-56:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MV3" Then
                    lblBrnCon3.Caption = IIf(lblBrnCon3.Caption <> "", lblBrnCon3.Caption & ", ", "") & "VAN-55:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MV4" Then
                    lblBrnCon3.Caption = IIf(lblBrnCon3.Caption <> "", lblBrnCon3.Caption & ", ", "") & "VAN-73:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                ElseIf !BranchCode = "MV5" Then
                    lblBrnCon3.Caption = IIf(lblBrnCon3.Caption <> "", lblBrnCon3.Caption & ", ", "") & "VAN-74:" & Format(!PAYMENTS, "MM/DD/YYYY") '& "(" & IIf(IsNull(!REMARKS), "", !REMARKS) & ")"
                End If
                .MoveNext
            Loop
        End If
        .Close
    End With
End Sub
Public Sub DataRefresh(MyCriteria As String)
Dim Qry As String

    DoEvents
    Qry = "CALL sp_mrcwarehouse.sp_All_Branch_Inv_Inquiry('" & MyCriteria & "');"
    
    If MyCriteria = "" Then Exit Sub
    With mainRs
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenStatic, adLockReadOnly
        
        If .RecordCount > 0 Then
            .MoveFirst
             flxParts.Clear
             flxParts.Rows = 2
             Set flxParts.DataSource = mainRs
             Call RXResizer1_ResizeComplete
        Else
            flxParts.Clear
            flxParts.Rows = 2
            Call Unload(frm_OnProcess2)
            Call RXResizer1_ResizeComplete
            Exit Sub
        End If
        .Close
    End With
End Sub
Public Sub flxColChanger()
Dim i As Long
On Error Resume Next
With flxParts
    '.Clear
    '.Rows = 2
    '.FixedCols = 1
    '               0       1             2            3           4      5       6      7        8              9
    '.FormatString = "|<PART NUMBER |<DESCRIPTION  |^NET OF VAT  |>SRP  |>WSP  |>QTY. |>SIX MSQ |>BACK ORDER |<APPLICABLE MODEL"
    
    .ColWidth(0) = 200
    .ColWidth(1) = 0 '(.Width - .ColWidth(0)) * 0.15  'PARTNUMBER
    .ColWidth(2) = (.Width - .ColWidth(0)) * 0.15 'PARTNUMBER
    .ColWidth(3) = (.Width - .ColWidth(0)) * 0.2 'DESCRIPTION
    
    For i = 4 To .Cols - 1
        .ColWidth(i) = (.Width - .ColWidth(0)) * 0.1 'branches
    Next i
    
End With
End Sub

Private Sub cboFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdEvaluate_Click    '     cmdFind_Click
End Sub

Private Sub cmdEvaluate_Click()
    Dim rsSP As New ADODB.Recordset
    
    Dim MySP() As String, Mdl() As String, Qry As String
    Dim myW As String, myW1 As String, myW2 As String, myW3 As String
    Dim MySel As String, i As Long, r As Long, rM As Integer
    Dim QHGray As Boolean, BOGray As Boolean, SQGray As Boolean
    
    myW1 = ""
    If cboFind.Text <> "" Then
        MySP = Split(cboFind.Text, " ")
        For i = LBound(MySP) To UBound(MySP)
            If myW1 = "" Then
                'myW1 = "(SPAREPARTS.DESCRIPTION) Like '%" & MySP(i) & "%'"
                
                If UpDRef.Value = 0 Then
                    myW1 = "PARTNUMBER Like \'%" & MySP(i) & "%\'"
                Else
                    myW1 = "DESCRIPTION Like \'%" & MySP(i) & "%\'"
                End If
            Else
                'myW1 = myW1 & " AND (SPAREPARTS.DESCRIPTION) Like '%" & MySP(i) & "%'"
                If UpDRef.Value = 0 Then
                    myW1 = myW1 & " AND PARTNUMBER Like \'%" & MySP(i) & "%\'"
                Else
                    myW1 = myW1 & " AND DESCRIPTION Like \'%" & MySP(i) & "%\'"
                End If
            End If
        Next i
        myW = myW1
    End If

    Call DataRefresh(myW1)
    
End Sub

Private Sub RXResizer1_ResizeComplete()
Call flxColChanger
End Sub

Private Sub tmrAutoLoad_Timer()
On Error GoTo er
    
    If tmrAutoLoad.Enabled = False Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    tmrAutoLoad.Enabled = False
    
    
    ActRow = 1
    
    'Set Search Description
    UpDRef.Value = 1
    
    DoEvents
    Call LoadBackUp
    
    If LoadOnProcess Then
    Call Unload(frm_OnProcess2)
    Set frm_OnProcess2 = Nothing
    End If
    
    Screen.MousePointer = Default

Exit Sub
er:
    tmrAutoLoad.Enabled = False
    Screen.MousePointer = vbDefault
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    Call modObjTools.ErrorMessage(Err.Description, "Error->tmrAutoLoad_Timer()")
End Sub

'Private Sub tmrLoadBackUp_Timer()
'Dim Qry As String
'    If tmrLoadBackUp.Enabled = False Then Exit Sub
'    Screen.MousePointer = vbHourglass
'
'    tmrLoadBackUp.Enabled = False
'
'    If LoadOnProcess Then
'    Call Unload(frm_OnProcess2)
'    Set frm_OnProcess2 = Nothing
'    End If
'
'    Screen.MousePointer = Default
'
'Exit Sub
'er:
'    tmrLoadBackUp.Enabled = False
'    Screen.MousePointer = vbDefault
'    If LoadOnProcess Then
'        Call Unload(frm_OnProcess2)
'        Set frm_OnProcess2 = Nothing
'    End If
'    Call modObjTools.ErrorMessage(Err.Description, "Error->tmrLoadBackUp_Timer()")
'End Sub

Private Sub UpDRef_Change()
    If UpDRef.Value = 0 Then
        lblRefSearch.Caption = "FIND PARTNUMBER"
    Else
        lblRefSearch.Caption = "FIND PARTS DESC."
    End If
End Sub

Private Sub Form_Load()
On Error GoTo er

    DoEvents
    tmrAutoLoad.Enabled = True
    
    DoEvents
    Call Load(frm_OnProcess2)
    frm_OnProcess2.Show vbModal
    
    
Exit Sub
er:
    tmrAutoLoad.Enabled = False
    Call modObjTools.ErrorMessage(Err.Description, "Error!->Form_Load()")
End Sub
