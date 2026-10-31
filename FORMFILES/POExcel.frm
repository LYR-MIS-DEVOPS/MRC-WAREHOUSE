VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form POexcel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import P.O. from Excel"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   ClipControls    =   0   'False
   Icon            =   "POExcel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboSheets 
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   2055
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   3555
      Visible         =   0   'False
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog Cdlg 
      Left            =   120
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin lvButton.lvButtons_H cmdImport 
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Import"
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
      Image           =   "POExcel.frx":058A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
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
      Image           =   "POExcel.frx":0B24
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxPO 
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4260
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   16777215
      ScrollBars      =   2
      FormatString    =   "^N O.    |^BARCODE                        |^Q T Y.               |^LOCATION                           |^DATETIME                  "
   End
   Begin lvButton.lvButtons_H cmdOpen 
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Open"
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
      Image           =   "POExcel.frx":10BE
      cBack           =   -2147483633
   End
   Begin VB.Label Label3 
      Caption         =   "Select WorkSheet:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   2820
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "TOTAL:"
      Height          =   255
      Left            =   6960
      TabIndex        =   4
      Top             =   2820
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Part Number to be copied:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "POexcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim db As DAO.Database
Dim rs As DAO.Recordset
Private filepath As String
Private sheetname As String

Dim aArray(1000, 2) As String
Dim xl As New Excel.Application
Dim xlsheet As Excel.Worksheet
Dim xlwbook As Excel.Workbook
Dim iLastrow, iRow As Integer
Dim iLastCol, icol As Integer

Private Sub cmdOpen_Click()
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
    

    'xl.ActiveWorkbook.Close True, Me.Cdlg.filename
    'xl.Quit
End Sub

Private Sub cboSheets_Click()
Dim x, z, i

    Set xlsheet = xlwbook.Sheets.Item(cboSheets.ListIndex + 1)
    xlsheet.Activate
    
    With xlsheet
       iLastCol = .Cells(1, Columns.Count).End(xlToLeft).Column
       iLastrow = .Cells.SpecialCells(xlCellTypeLastCell).row
    End With
 
    Me.flxPO.Rows = iLastrow
    Me.flxPO.Cols = iLastCol + 1
    i = 0
    For iRow = 1 To iLastrow
        If i > 0 Then Me.flxPO.TextMatrix(i, 0) = iRow - 1
        For icol = 1 To iLastCol
            x = iRow - 1 'Flexgrid # of rows
            z = icol 'flexgrid # of columns
            If icol = iLastCol Then
                Me.flxPO.TextMatrix(x, z) = Format(xl.Cells(iRow, icol), "mm/dd/yyyy hh:mm:ss")
            Else
                Me.flxPO.TextMatrix(x, z) = UCase(xl.Cells(iRow, icol))
            End If
        Next icol
        i = i + 1
    Next iRow
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set xlwbook = Nothing
    Set xl = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdImport_Click()
    Dim i As Integer, j As Integer
    Dim x As Integer, z As Integer
    Dim Qry As String
    
    cmdImport.Enabled = False
    cmdOpen.Enabled = False
    cmdCancel.Enabled = False
    With frmMonthlyPhysicalInv
        pBar.Visible = True
        pBar.Value = 0
        pBar.Max = flxPO.Rows - 1
        .flxParts.Rows = 1
        .flxParts.Rows = flxPO.Rows
'        .flxParts.Cols = flxPO.Cols
        For i = 1 To flxPO.Rows
            Qry = "Select PARTNUMBER, DESCRIPTION, REMARKS AS CLASSIFICATION From spareparts Where BCODEVALUE = '" & flxPO.TextMatrix(i, 1) & "' "
            
            If rsAnyUse.State = adStateOpen Then rsAnyUse.Close
            rsAnyUse.Open Qry, LYRcon, adOpenStatic, adLockReadOnly
                
            If rsAnyUse.RecordCount > 0 Then
                rsAnyUse.MoveFirst
                .flxParts.TextMatrix(i, 1) = i
                .flxParts.TextMatrix(i, 2) = flxPO.TextMatrix(i, 1)
                .flxParts.TextMatrix(i, 3) = rsAnyUse!PARTNUMBER
                .flxParts.TextMatrix(i, 4) = rsAnyUse!Description
                .flxParts.TextMatrix(i, 5) = flxPO.TextMatrix(i, 2)
                .flxParts.TextMatrix(i, 6) = rsAnyUse!CLASSIFICATION
                .flxParts.TextMatrix(i, 7) = flxPO.TextMatrix(i, 3)
                .flxParts.TextMatrix(i, 8) = flxPO.TextMatrix(i, 4)
            Else
                .flxParts.TextMatrix(i, 1) = i
                .flxParts.TextMatrix(i, 2) = flxPO.TextMatrix(i, 1)
                .flxParts.TextMatrix(i, 3) = ""
                .flxParts.TextMatrix(i, 4) = ""
                .flxParts.TextMatrix(i, 5) = flxPO.TextMatrix(i, 2)
                .flxParts.TextMatrix(i, 6) = ""
                .flxParts.TextMatrix(i, 7) = flxPO.TextMatrix(i, 3)
                .flxParts.TextMatrix(i, 8) = flxPO.TextMatrix(i, 4)
            End If
'            For j = 1 To .flxParts.Cols - 1
'                x = i - 1
'                z = j
'                .flxParts.TextMatrix(x, z) = flxPO.TextMatrix(x, z)
'
'            Next j
'            If x > 0 Then pBar.Value = pBar.Value + 1
            pBar.Value = pBar.Value + 1
            DoEvents
        Next i
    End With
'    With SPordering
'        pBar.Visible = True
'        pBar.Value = 0
'        pBar.Max = flxPO.Rows - 1
'        .flxParts.Rows = 1
'        For i = 1 To flxPO.Rows - 1
'            .flxParts.AddItem flxPO.TextMatrix(i, 0) & vbTab & flxPO.TextMatrix(i, 1) & vbTab & vbTab & flxPO.TextMatrix(i, 2)
'            pBar.Value = pBar.Value + 1
'            DoEvents
'        Next i
'    End With

    pBar.Visible = False
    Unload Me
    DoEvents
    
    'Call SPordering.ProcessImportedPO
    
End Sub

'Private Sub cmdOpen_Click()
'    Dim x As Integer, TtlQty As Long
'    Dim zQty As Integer
'    Dim InvldRow As String
'    Dim con As ADODB.Connection
'    Dim rs As ADODB.Recordset
'    Dim strSQL As String
'
'    On Error GoTo LocalErr
'    With Cdlg
'        .CancelError = True
'        .filter = "Microsoft Excel (*.xls)|*.xls|All Files(*.*)|*.*"
'        .ShowOpen
'        filepath = .filename
'        If filepath = "" Then
'            MsgBox "Please select a File Name to continue!", vbExclamation + vbOKOnly, "Import P.O."
'            Exit Sub
'        End If
'    End With
'    sheetname = "Sheet1"
'    'con.Provider = "Microsoft.Jet.OLEDB.4.0"
'    'con.ConnectionString = "Data Source = " & filepath & ";" & _
'                            "Extended Properties = Excel 8.0;"
'    'con.Open
'
'    'strSQL = "SELECT * FROM [" & filepath & "$]"
'    'rs.Open strSQL, con, adOpenStatic, adLockOptimistic
'    Set db = OpenDatabase(filepath, False, False, "Excel 8.0;HDR=yes;")
'    Set rs = db.OpenRecordset(sheetname)
'    If rs.RecordCount > 0 Then
'        rs.MoveFirst
'        Screen.MousePointer = 11
'        x = 1
'        TtlQty = 0
'        zQty = 0
'        flxPO.Rows = 1
'        While rs.EOF <> True
'            If IsNull(rs.Fields(0)) Or IsNull(rs.Fields(1)) Then
'                zQty = zQty + 1
'                InvldRow = IIf(InvldRow = "", "Row " & x + 1, InvldRow & ", " & x + 1)
'            Else
'                If rs.Fields(0) = "" Or Val(rs.Fields(1)) = 0 Then
'                    zQty = zQty + 1
'                    InvldRow = IIf(InvldRow = "", "Row " & x + 1, InvldRow & ", " & x + 1)
'                Else
'                    flxPO.AddItem x & vbTab & rs.Fields(0) & vbTab & rs.Fields(1)
'                    TtlQty = TtlQty + rs.Fields(1)
'                    x = x + 1
'                End If
'            End If
'            rs.MoveNext
'        Wend
'        Screen.MousePointer = 0
'        lblTotal.Caption = TtlQty
'        If zQty > 0 Then
'            MsgBox zQty & IIf(zQty > 1, " Records", " Record") & " found Invalid!" & vbCrLf & _
'            "Record is either blank Part number or zero quantity." & vbCrLf & "Please check source at the following rows:" & vbCrLf & _
'            InvldRow, vbExclamation + vbOKOnly, "Import P.O."
'        End If
'    Else
'        MsgBox "No records found in file!", vbExclamation + vbOKOnly, "Import P.O."
'    End If
'    Exit Sub
'LocalErr:
'    Screen.MousePointer = 0
'    If Err.Number = 32755 Then Exit Sub
'    If Err.Number = 3125 Then
'        MsgBox "No P.O. Sheet found in File!" & vbCrLf & "Please use POSheet Worksheet w/ Part Number and Qty.", vbCritical + vbOKOnly, "Error"
'    Else
'        MsgBox Err.Description & vbCrLf & "Please contact System Administrator", vbCritical + vbOKOnly, "Error"
'    End If
'End Sub

Private Sub Form_Load()
    Call cmdOpen_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set rs = Nothing
    Set db = Nothing
End Sub
