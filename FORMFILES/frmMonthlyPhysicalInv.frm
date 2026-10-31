VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMonthlyPhysicalInv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monthly Physical Inventory (AUDIT)"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8865
   ScaleWidth      =   13380
   StartUpPosition =   1  'CenterOwner
   Tag             =   "XY"
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1620
      TabIndex        =   24
      Top             =   0
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   661
      Caption         =   "&Print Preview"
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
      Image           =   "frmMonthlyPhysicalInv.frx":0000
      cBack           =   -2147483633
   End
   Begin VB.Frame frmFooter 
      Height          =   975
      Left            =   120
      TabIndex        =   19
      Tag             =   "XYF"
      Top             =   7800
      Width           =   13095
      Begin VB.TextBox txtPreparedby 
         Height          =   285
         Left            =   1320
         TabIndex        =   21
         Tag             =   "XY"
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label12 
         Caption         =   "PREPARED BY:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Tag             =   "XY"
         Top             =   375
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6255
      Left            =   120
      TabIndex        =   12
      Tag             =   "XYF"
      Top             =   1560
      Width           =   13095
      Begin MSFlexGridLib.MSFlexGrid flxParts 
         Height          =   5535
         Left            =   120
         TabIndex        =   16
         Tag             =   "XY"
         Top             =   600
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   9763
         _Version        =   393216
         Cols            =   9
         FormatString    =   $"frmMonthlyPhysicalInv.frx":059A
      End
      Begin VB.TextBox txtFind 
         Height          =   315
         Left            =   11160
         TabIndex        =   15
         Tag             =   "XY"
         Top             =   240
         Width           =   1815
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   300
         Left            =   10200
         TabIndex        =   13
         Tag             =   "XY"
         ToolTipText     =   "Find Part Number"
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "FIND"
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
         Image           =   "frmMonthlyPhysicalInv.frx":0660
         cBack           =   -2147483633
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "Add &New"
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
      Image           =   "frmMonthlyPhysicalInv.frx":0BFA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3180
      TabIndex        =   1
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "E&xit"
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
      Image           =   "frmMonthlyPhysicalInv.frx":1194
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4920
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Save"
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
      Image           =   "frmMonthlyPhysicalInv.frx":172E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6480
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Cancel"
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
      Image           =   "frmMonthlyPhysicalInv.frx":1CC8
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Tag             =   "XYF"
      Top             =   360
      Width           =   13095
      Begin VB.ComboBox cboSheets 
         Height          =   315
         Left            =   2040
         TabIndex        =   22
         Tag             =   "XY"
         Top             =   705
         Width           =   2055
      End
      Begin VB.TextBox txtBatchNo 
         Height          =   315
         Left            =   10320
         TabIndex        =   10
         Tag             =   "XY"
         Top             =   315
         Width           =   2655
      End
      Begin VB.ComboBox cboBranch 
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Tag             =   "XY"
         Top             =   320
         Width           =   5415
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   375
         Left            =   10320
         TabIndex        =   8
         Tag             =   "XY"
         Top             =   705
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin lvButton.lvButtons_H cmdXLS 
         Height          =   300
         Left            =   4200
         TabIndex        =   14
         Tag             =   "XY"
         ToolTipText     =   "Import P.O. from Excel"
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         Caption         =   "IMPORT FILE"
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
         Image           =   "frmMonthlyPhysicalInv.frx":2262
         cBack           =   -2147483633
      End
      Begin VB.Label Label3 
         Caption         =   "SELECT WORKSHEET:"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Tag             =   "XY"
         Top             =   750
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "BATCH NO.:"
         Height          =   255
         Left            =   9000
         TabIndex        =   11
         Tag             =   "XY"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "MM/DD/YYYY"
         Height          =   255
         Left            =   11880
         TabIndex        =   9
         Tag             =   "XY"
         Top             =   750
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE:"
         Height          =   255
         Left            =   9600
         TabIndex        =   7
         Tag             =   "XY"
         Top             =   750
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "BRANCH:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Tag             =   "XY"
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   11760
      TabIndex        =   17
      Text            =   "HoldBackColor"
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog Cdlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   10800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   13470
      m_MDI_Height    =   9285
   End
   Begin VB.Menu mnuStatus 
      Caption         =   "Popupmenu"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Row"
      End
   End
End
Attribute VB_Name = "frmMonthlyPhysicalInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xl As New Excel.Application
Dim xlsheet As Excel.Worksheet
Dim xlwbook As Excel.Workbook

Dim arrBCODE() As String
Dim BCODE As String
Dim formAction As String
Dim aArray(1000, 2) As String
Dim Qry As String

Dim iLastrow, iRow As Integer
Dim iLastCol, icol As Integer
Dim i, j As Integer

Private Sub SetButtonVisibility(DataEntry As Boolean)
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->VIEW SUMM
'3 ->EXIT
'4 ->PRINT
'5 ->SAVE
'6 ->CANCEL

    'hide all button
    For i = 0 To 2
        cmdTrans(i).Visible = False
    Next i

    If DataEntry Then
        cmdTrans(3).Visible = True
        cmdTrans(4).Visible = True
    Else
        For i = 0 To 2
            cmdTrans(i).Visible = True
            cmdTrans(3).Visible = False
            cmdTrans(4).Visible = False
        Next i
    End If
End Sub

Sub Clearform()

    cboBranch.Text = ""
    cboSheets.Text = ""
    txtBatchNo.Text = "<AUTO>"
    txtDate.Text = Format(Now, "mm/dd/yyyy")
    txtFind.Text = ""
    txtPreparedby.Text = ""
    
    BCODE = ""
End Sub

Public Sub CboDisplay()
Dim cboRs As New ADODB.Recordset
On Error GoTo er
    
    'Get Data from database
    DoEvents
    Qry = "Select BRANCHNAME,BRANCHCODE From branch Where Branchcategory IN ('MRC','WHOLESALE','DEPOT','WAREHOUSE')"
    If cboRs.State = adStateOpen Then cboRs.Close
    cboRs.Open Qry, LYRcon, adOpenStatic, adLockReadOnly
    If cboRs.State = 1 Then
        Call modObjTools.addToListObjectWithArrayGivenField(cboBranch, cboRs, arrBCODE, "BRANCHNAME", "BRANCHCODE")
    End If
    If cboRs.State <> 0 Then cboRs.Close
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub

Private Sub cboSheets_Click()
Dim X, z, i

    Set xlsheet = xlwbook.Sheets.Item(cboSheets.ListIndex + 1)
    xlsheet.Activate
    
    With xlsheet
       iLastCol = .Cells(1, Columns.Count).End(xlToLeft).Column
       iLastrow = .Cells.SpecialCells(xlCellTypeLastCell).row
    End With
    
    flxParts.Rows = 1
    flxParts.Rows = iLastrow
    flxParts.Cols = iLastCol + 2
    flxParts.ColWidth(0) = 500
    i = 0
    For iRow = 1 To iLastrow
        If i > 0 Then
            flxParts.TextMatrix(i, 1) = iRow - 1
            flxParts.ColWidth(i + 1) = (flxParts.Width - flxParts.ColWidth(0)) * 0.15
        End If
        For icol = 1 To iLastCol
            X = iRow - 1 'Flexgrid # of rows
            z = icol  'flexgrid # of columns
            If X = 0 And z = 1 Then
                flxParts.TextMatrix(X, z) = "NO."
                z = z + 1
                flxParts.TextMatrix(X, z) = UCase(xl.Cells(iRow, icol))
            Else
                If icol = iLastCol Then
                    z = z + 1
                    flxParts.TextMatrix(X, z) = Format(xl.Cells(iRow, icol), "mm/dd/yyyy hh:mm:ss")
                Else
                    z = z + 1
                    flxParts.TextMatrix(X, z) = UCase(xl.Cells(iRow, icol))
                End If
            End If
        Next icol
        i = i + 1
    Next iRow
End Sub

Private Sub cmdTrans_Click(Index As Integer)
On Error GoTo er
    If Index = 0 Then 'add
        Call Clearform
        Call LockData(True)
        Call SetButtonVisibility(True)
        formAction = "add"
        txtPreparedby.Text = GetEmployeeName
        txtBatchNo.Text = "<AUTO>"
    ElseIf Index = 2 Then 'exit
        formAction = "exit"
        Unload Me
    End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdTrans-Click()")
Unload Me
End Sub

Private Sub cmdXLS_Click()
Dim nos As Integer
Dim N As Integer
    
    With Me.Cdlg
        .DialogTitle = "Select Excel Spreasheet"
        .InitDir = App.Path
        .filter = "Excel Spreadsheets (*.xls)/*.xls/All Files (*.*)/*.*"
        .ShowOpen
    End With
    If Me.Cdlg.filename = "" Then Exit Sub
    Set xlwbook = xl.Workbooks.Open(Me.Cdlg.filename)
    nos = xlwbook.Sheets.Count
    
    cboSheets.Clear
    For N = 1 To nos
        cboSheets.AddItem xlwbook.Sheets(N).Name
    Next
'POexcel.Show vbModal
End Sub

Private Sub Form_Load()
    Call LockData(False)
    txtDate.Text = Format$(Now, "mm/dd/yyyy")
    txtBatchNo.Text = "<AUTO>"
    Call CboDisplay
    '===================BUTTON POSITION===================
    If cmdTrans(0).Visible = False Then
        'Align Save Button
        cmdTrans(3).Left = cmdTrans(0).Left
        cmdTrans(3).Top = cmdTrans(0).Top
        'Align Cancel Button
        cmdTrans(4).Left = cmdTrans(1).Left
        cmdTrans(4).Top = cmdTrans(1).Top
    End If
End Sub

Sub LockData(DataEntry As Boolean)
    DataEntry = Not DataEntry
    
    'frmHeader
    cboBranch.Locked = DataEntry
    cboSheets.Locked = DataEntry
    txtBatchNo.Locked = DataEntry
    txtDate.Enabled = Not DataEntry
    cmdXLS.Enabled = Not DataEntry
    
    'frmBody
    txtFind.Locked = DataEntry
    cmdFind.Enabled = Not DataEntry
    
    'frmFooter
    txtPreparedby.Locked = DataEntry
    
    If cboBranch.Locked Then
        'frmHeader
        cboBranch.BackColor = txtHoldGrayBackColor.BackColor
        cboSheets.BackColor = txtHoldGrayBackColor.BackColor
        txtBatchNo.BackColor = txtHoldGrayBackColor.BackColor
        'txtDate.BackColor = txtHoldGrayBackColor.BackColor
        'frmBody
        txtFind.BackColor = txtHoldGrayBackColor.BackColor
        'frmFooter
        txtPreparedby.BackColor = txtHoldGrayBackColor.BackColor
    Else
        'frmHeader
        cboBranch.Locked = True
        cboSheets.BackColor = txtHoldWhiteBackColor.BackColor
        txtBatchNo.Locked = True
        txtDate.Enabled = False
        'frmBody
        txtFind.BackColor = txtHoldWhiteBackColor.BackColor
        'frmFooter
        txtPreparedby.Locked = True
    End If
End Sub
