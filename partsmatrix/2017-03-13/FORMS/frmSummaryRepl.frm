VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSummaryRepl 
   Caption         =   "View Summary"
   ClientHeight    =   9360
   ClientLeft      =   120
   ClientTop       =   255
   ClientWidth     =   14520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9360
   ScaleWidth      =   14520
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10080
      Top             =   240
   End
   Begin VB.Frame frmeSummary 
      Height          =   8415
      Left            =   120
      TabIndex        =   1
      Tag             =   "xyf"
      Top             =   720
      Width           =   14295
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   375
         Left            =   5640
         TabIndex        =   14
         Tag             =   "xyf"
         ToolTipText     =   "Execute"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
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
         Image           =   "frmSummaryRepl.frx":0000
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cbBranch 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Tag             =   "xyf"
         Top             =   240
         Width           =   4215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxviewSummary 
         Height          =   7575
         Left            =   120
         TabIndex        =   3
         Tag             =   "xyf"
         Top             =   720
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   13361
         _Version        =   393216
         SelectionMode   =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox txtFind 
         Height          =   330
         Left            =   12240
         TabIndex        =   2
         Tag             =   "xyf"
         Top             =   240
         Width           =   1935
      End
      Begin lvButton.lvButtons_H btnFind 
         Height          =   315
         Left            =   10800
         TabIndex        =   9
         Tag             =   "xyf"
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   "Find"
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
         Image           =   "frmSummaryRepl.frx":059A
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dFrom 
         Height          =   315
         Left            =   1560
         TabIndex        =   11
         Tag             =   "xyf"
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   53870593
         CurrentDate     =   39322
      End
      Begin MSComCtl2.DTPicker dTo 
         Height          =   315
         Left            =   3840
         TabIndex        =   12
         Tag             =   "xyf"
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   53870593
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdPrev 
         Height          =   375
         Left            =   3360
         TabIndex        =   15
         Tag             =   "xyf"
         ToolTipText     =   "Hide Date End"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
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
         Image           =   "frmSummaryRepl.frx":0B34
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   375
         Left            =   3360
         TabIndex        =   13
         ToolTipText     =   "Show Date End"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
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
         Image           =   "frmSummaryRepl.frx":10CE
         cBack           =   -2147483633
      End
      Begin VB.Label lblBranch 
         Alignment       =   1  'Right Justify
         Caption         =   "BRANCH"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Tag             =   "xyf"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Date Covered:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Tag             =   "xyf"
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin lvButton.lvButtons_H cmdShow 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "Show &Form"
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
      Image           =   "frmSummaryRepl.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "E&xit Report"
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
      Image           =   "frmSummaryRepl.frx":1C02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   375
      Left            =   1920
      TabIndex        =   8
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
      Image           =   "frmSummaryRepl.frx":219C
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   8040
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14640
      m_MDI_Height    =   9870
   End
   Begin VB.Label lblviewSummary 
      BackStyle       =   0  'Transparent
      Caption         =   "PARTS MATRIX SUMMARY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Tag             =   "xyf"
      Top             =   480
      Width           =   4095
   End
   Begin VB.Menu mnuSum 
      Caption         =   "Summary"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmSummaryRepl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Strings Variables
Dim Qry As String
Dim qohQry As String
Dim strTtlExcess As String, strExcess As String
Dim strTtlLacking As String, strLacking As String
Dim strTtlEmpty As String, strEmpty As String
Dim arrBRANCH() As String
Dim branchCode As String

Dim strPartial As String, strTotal As String

Public mySrc As String

'Recordset Variables
Dim mainRs As New ADODB.Recordset
Dim qohRs As New ADODB.Recordset

'Integer Variables
Dim ActRow As Integer, TCol As Integer
Dim fRow As Integer, fCol As Integer
Dim i As Integer, j As Integer
Dim TtlQOH As Integer, TtlPMATRIXID As Integer

Dim StatSrch As Boolean


Public Sub CboDisplay()
On Error GoTo er

   
        'Get Data from database
        DoEvents
        Call modPMBranchIdentifying.sp_get_BranchSumm(mainRs)
        If mainRs.State = 1 Then
           
            Call modObjTools.fillArray(modGlobalVar.AllArrBranch, mainRs)

        End If
        If mainRs.State <> 0 Then mainRs.Close
        
        'Load Data to combo box
        Call SetCboListData
        
   
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub

Sub SetCboListData()
On Error GoTo er

    If cbBranch.Text = "" Then
        DoEvents
        Call modObjTools.addToListObjectUsingArrayWithCols(cbBranch, arrBRANCH, modGlobalVar.AllArrBranch, 4, 1)
     End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->SetCboListData()")
End Sub

Private Sub cbBranch_Change()
On Error GoTo er
    If StatSrch Then
            StatSrch = False
            Exit Sub
    End If

    With cbBranch
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->cbBranch-Change()")
End Sub

Private Sub cbBranch_GotFocus()
Dim Tmp
    Const CB_SHOWDROPDOWN = &H14F
    Tmp = SendMessage(cbBranch.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub


Private Sub cbBranch_Click()
   
    
    Call cbBranch_Change
    flxviewSummary.Clear 'used to refresh the flexgrid
    Call viewSummFlxChanger 'used to refresh the flexgrid
    DoEvents
    
    branchCode = arrBRANCH(cbBranch.ListIndex, 2)
    
    Call ShowRecord(branchCode)
    
End Sub

Public Sub ShowRecord(MyBranchCode As String)
On Error GoTo er

    Screen.MousePointer = vbHourglass
    
    DoEvents
    Call modPMBranchIdentifying.sp_brnchIdentify_summary(MyBranchCode, mainRs)
    
    If mainRs.State <> 1 Then Exit Sub
        
        If mainRs.RecordCount <= 0 Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        With mainRs
            If .RecordCount > 0 Then
                .MoveFirst
                flxviewSummary.Rows = .RecordCount + 1
                flxviewSummary.Cols = .Fields.Count + 1
        
                For i = 1 To flxviewSummary.Rows - 1
                    flxviewSummary.TextMatrix(i, 1) = !MODELBRAND
                    flxviewSummary.TextMatrix(i, 2) = !TOTALPN
                    flxviewSummary.TextMatrix(i, 3) = !TOTALSRQ
                    flxviewSummary.TextMatrix(i, 4) = !TOTALALQ
                    flxviewSummary.TextMatrix(i, 5) = !RefNo
                    flxviewSummary.TextMatrix(i, 6) = !DATEPREPARED
                    flxviewSummary.TextMatrix(i, 7) = !DATEFINAL
                    flxviewSummary.TextMatrix(i, 8) = !PROCESSEDBY
                    flxviewSummary.TextMatrix(i, 9) = !PMATRIXID
                    flxviewSummary.TextMatrix(i, 10) = !REFID
                    mainRs.MoveNext
                Next i
        
            End If
          
        End With
        
        If mainRs.State <> 0 Then
            Screen.MousePointer = vbDefault
            mainRs.Close
        End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->SummaryShowRecord()")
End Sub

Public Sub viewSummFlxChanger()

    If mySrc = "Data Entry" Then
        With flxviewSummary
            .Cols = 12
            '                        1            2          3           4          5          6          7         8           9            10
            .FormatString = "|<MODEL BRAND |^TOTAL PN |^COMPLETE |^MAINTENANCE |^PLASTIC |^UNIVERSAL |^OTHERS |^GENUINE |^REPLACEMENT |^LASTUPDATE"
            
            .ColWidth(0) = 200
            .ColWidth(1) = (.Width - .ColWidth(0)) * 0.23  'MODEL BRAND
            .ColWidth(2) = (.Width - .ColWidth(0)) * 0.07 'TOTAL PN
            .ColWidth(3) = (.Width - .ColWidth(0)) * 0.08 'COMPLETE
            .ColWidth(4) = (.Width - .ColWidth(0)) * 0.1 'MAINTENANCE
            .ColWidth(5) = (.Width - .ColWidth(0)) * 0.07 'PLASTIC
            .ColWidth(6) = (.Width - .ColWidth(0)) * 0.09 'UNIVERSAL
            .ColWidth(7) = (.Width - .ColWidth(0)) * 0.07 'OTHERS
            .ColWidth(8) = (.Width - .ColWidth(0)) * 0.07 'GENUINE
            .ColWidth(9) = (.Width - .ColWidth(0)) * 0.1 'REPLACEMENT
            .ColWidth(10) = 0 '(.Width - .ColWidth(0)) * 0.1 'LASTUPDATE
            .ColWidth(11) = 0   '(.Width - .ColWidth(0)) * 0.1 'ID
           
        End With
        
    ElseIf mySrc = "Branch Quantity" Then
    
        With flxviewSummary
            .Cols = 12
           '                      1         2          3          4           5           6        7         8        9         10
            .FormatString = " |<BRANCH |^MC BRAND |^MC MODEL |^TOTAL PN |^TOTAL STDQ |^TOTAL PG |^EXCESS |^LACKING |^EMPTY |^UPDATED BY"
        
            .ColWidth(0) = 200
            .ColWidth(1) = (.Width - .ColWidth(0)) * 0.15 'BRANCH
            .ColWidth(2) = (.Width - .ColWidth(0)) * 0.11 'BRAND
            .ColWidth(3) = (.Width - .ColWidth(0)) * 0.11 'MODEL
            .ColWidth(4) = (.Width - .ColWidth(0)) * 0.11 'TOTAL PN
            .ColWidth(5) = (.Width - .ColWidth(0)) * 0.11 'TOTAL STDQ
            .ColWidth(6) = (.Width - .ColWidth(0)) * 0.1 'TOTAL PARTSGROUP
            .ColWidth(7) = (.Width - .ColWidth(0)) * 0.1 'EXCESS
            .ColWidth(8) = (.Width - .ColWidth(0)) * 0.1 'LACKING
            .ColWidth(9) = (.Width - .ColWidth(0)) * 0.1 'EMPTY
            .ColWidth(10) = 0 '(.Width - .ColWidth(0)) * 0.11 'BCODE
            .ColWidth(11) = 0 '(.Width - .ColWidth(0)) * 0.03 'PMATRIXID
            lblviewSummary.Caption = "Branch Quantity Summary"
        End With
        
    ElseIf mySrc = "Branch Identifying" Then
    
        With flxviewSummary
            .Cols = 11
          '                           1            2           3           4           5           6               7              8
            .FormatString = " |<MODEL BRAND |^TOTAL PN |^TOTAL SRQ |^TOTAL ALQ |^REFERENCE NO. |^DATE PREPARED |^DATE FINAL |^PROCESSED BY"
        
            .ColWidth(0) = 200
            .ColWidth(1) = (.Width - .ColWidth(0)) * 0.15 'MODEL BRAND
            .ColWidth(2) = (.Width - .ColWidth(0)) * 0.1 'TOTAL PN
            .ColWidth(3) = (.Width - .ColWidth(0)) * 0.1 'TOTAL SRQ
            .ColWidth(4) = (.Width - .ColWidth(0)) * 0.1 'TOTAL ALQ
            .ColWidth(5) = (.Width - .ColWidth(0)) * 0.1 'REFERENCE NO.
            .ColWidth(6) = (.Width - .ColWidth(0)) * 0.1 'DATE PREPARED
            .ColWidth(7) = (.Width - .ColWidth(0)) * 0.1 'DATE FINAL
            .ColWidth(8) = (.Width - .ColWidth(0)) * 0.2 'PROCESSED BY
            .ColWidth(9) = 0 '(.Width - .ColWidth(0)) * 0.1 'PMATRIXID
            .ColWidth(10) = 0 '(.Width - .ColWidth(0)) * 0.03 'REFID
            lblviewSummary.Caption = "Branch Identifying Summary"
            lblBranch.Visible = True
            cbBranch.Visible = True
        End With
    
    ElseIf mySrc = "Replenishment StockTransfer" Then
    
        With flxviewSummary
            .Cols = 13
          '                           1            2           3           4           5           6       7         8          9        10
            .FormatString = " |^REF DATE |^FINAL DATE |^STF NO |<RECEIVING BRANCH |<ADDRESS |^REPL REF# |^STATUS |^REMARKS |^CHECKER |^DRIVER"
        
            .ColWidth(0) = 200
            .ColWidth(1) = (.Width - .ColWidth(0)) * 0.08 'REF DATE
            .ColWidth(2) = (.Width - .ColWidth(0)) * 0.08 'FINAL DATE
            .ColWidth(3) = (.Width - .ColWidth(0)) * 0.08 'STF NO
            .ColWidth(4) = (.Width - .ColWidth(0)) * 0.25 'RECEIVING BRANCH
            .ColWidth(5) = (.Width - .ColWidth(0)) * 0.25 'ADDRESS
            .ColWidth(6) = (.Width - .ColWidth(0)) * 0.1 'REPL REF#
            .ColWidth(7) = (.Width - .ColWidth(0)) * 0.1 'STATUS
            .ColWidth(8) = (.Width - .ColWidth(0)) * 0.1 'REMARKS
            .ColWidth(9) = (.Width - .ColWidth(0)) * 0.17 'CHECKER
            .ColWidth(10) = (.Width - .ColWidth(0)) * 0.17 'DRIVER
            .ColWidth(11) = 0 '(.Width - .ColWidth(0)) * 0.03 'BRANCHCODE
            .ColWidth(12) = 0 '(.Width - .ColWidth(0)) * 0.03 'BRANCHSPOID
            lblviewSummary.Caption = "Replenishment Stock Transfer Summary"
            Label1.Visible = True
            dFrom.Visible = True
            dTo.Visible = True
            cmdNext.Visible = True
            cmdPrev.Visible = True
            cmdRefresh.Visible = True
        End With
        
    End If
End Sub

Public Sub viewSummary()
    
On Error GoTo er
    If mySrc = "Data Entry" Then
        
        Screen.MousePointer = vbHourglass
        
        DoEvents
        Call modPMDataEntry.sp_dataentry_summary(mainRs)
    
        If mainRs.State <> 1 Then Exit Sub
        
        If mainRs.RecordCount <= 0 Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        With mainRs
            If .RecordCount > 0 Then
                .MoveFirst
                flxviewSummary.Rows = .RecordCount + 1
                flxviewSummary.Cols = .Fields.Count + 1
        
                For i = 1 To flxviewSummary.Rows - 1
                    flxviewSummary.TextMatrix(i, 1) = !PARTSMATRIX
                    flxviewSummary.TextMatrix(i, 2) = !TOTALPN
                    flxviewSummary.TextMatrix(i, 3) = !COMPLETE
                    flxviewSummary.TextMatrix(i, 4) = !MAINTENANCE
                    flxviewSummary.TextMatrix(i, 5) = !PLASTIC
                    flxviewSummary.TextMatrix(i, 6) = !UNIVERSAL
                    flxviewSummary.TextMatrix(i, 7) = !Others
                    flxviewSummary.TextMatrix(i, 8) = !GENUINE
                    flxviewSummary.TextMatrix(i, 9) = !REPLACEMENT
                    flxviewSummary.TextMatrix(i, 10) = Format$(!LASTUPDATE, "mm-dd-yyyy")
                    flxviewSummary.TextMatrix(i, 11) = !PMATRIXID
                    mainRs.MoveNext
                Next i
        
            End If
          
        End With
        
        If mainRs.State <> 0 Then
            Screen.MousePointer = vbDefault
            mainRs.Close
        End If
        
    ElseIf mySrc = "Branch Quantity" Then

        Screen.MousePointer = vbHourglass
        
        DoEvents
        Call modPMBranchSQTY.sp_brnchQty_summary(mainRs)
        
        If mainRs.State <> 1 Then Exit Sub
        
        If mainRs.RecordCount <= 0 Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        With mainRs
            If .RecordCount > 0 Then
                .MoveFirst
                flxviewSummary.Rows = .RecordCount + 1
                flxviewSummary.Cols = .Fields.Count + 1
                
             
                For i = 1 To flxviewSummary.Rows - 1
                    TtlQOH = !PMTOTALQOH
                    TtlPMATRIXID = !PMTOTALPN
                    strExcess = Format$(!PMTOTALEXCESS, 0)
                    If Val(strExcess) <> 0 Then
                        
                        If Val(strExcess) <= TtlQOH Then
                            strTtlExcess = strExcess / TtlQOH * 100 & "%"
                        ElseIf strExcess >= TtlQOH Then
                            strTtlExcess = TtlQOH / strExcess * 100 & "%"
                        End If
                        
                    Else
                        strTtlExcess = "0.0%"
                    End If
                    
                    strLacking = Format$(!PMTOTALLACKING, 0)
                    If strLacking <> 0 Then
                    
                        If Val(strLacking) <= !PMTOTALSTDQ Then
                            strTtlLacking = strLacking / Format$(!PMTOTALSTDQ, 0) * 100 & "%"
                        ElseIf Val(strLacking) >= !PMTOTALSTDQ Then
                            strTtlLacking = Format$(!PMTOTALSTDQ, 0) / strLacking * 100 & "%"
                        End If
                        
                    Else
                        strTtlLacking = "0.0%"
                    End If
                    
                    strEmpty = Format$(!PMTOTALEMPTY, 0)
                    If strEmpty <> 0 Then
                    
                        If strEmpty <= TtlPMATRIXID Then
                            strTtlEmpty = strEmpty / TtlPMATRIXID * 100 & "%"
                        ElseIf strEmpty >= TtlPMATRIXID Then
                            strTtlEmpty = TtlPMATRIXID / strEmpty * 100 & "%"
                        End If
                        
                    Else
                        strTtlEmpty = "0.0%"
                    End If
                    
                    flxviewSummary.TextMatrix(i, 1) = !BNAME
                    flxviewSummary.TextMatrix(i, 2) = !PMBRAND
                    flxviewSummary.TextMatrix(i, 3) = !PMMODEL
                    flxviewSummary.TextMatrix(i, 4) = Format$(!PMTOTALPN, 0)
                    flxviewSummary.TextMatrix(i, 5) = Format$(!PMTOTALSTDQ, 0)
                    flxviewSummary.TextMatrix(i, 6) = !partsGroup
                    flxviewSummary.TextMatrix(i, 7) = Format$(strTtlExcess, "0.0%")
                    flxviewSummary.TextMatrix(i, 8) = Format$(strTtlLacking, "0.0%")
                    flxviewSummary.TextMatrix(i, 9) = Format$(strTtlEmpty, "0.0%")
                    flxviewSummary.TextMatrix(i, 10) = !BCODE
                    flxviewSummary.TextMatrix(i, 11) = !PMATRIXID
                  ' flxviewSummary.TextMatrix(i, 12) =
                    .MoveNext
               
                Next i
            End If
            
        End With
        
        If mainRs.State <> 0 Then
            Screen.MousePointer = vbDefault
            mainRs.Close
        End If
        
    ElseIf mySrc = "Branch Identifying" Then
        
        Call CboDisplay
    End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->viewSummary()")
End Sub

Private Sub btnFind_Click()
    Dim i As Integer, j As Integer

    If txtFind.Text = "" Then Exit Sub
    With flxviewSummary
        For i = fRow To .Rows - 1
            For j = fCol To TCol - 1
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = TCol - 1: .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = TCol - 1, i + 1, i)
                    fCol = IIf(j = TCol - 1, 1, j + 1)
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

Private Sub cbBranch_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cbBranch, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Call cbBranch_Click
    End If
End Sub

Private Sub cmdExit_Click()
    If mySrc = "Data Entry" Then
        DataEntry.Show
        Unload Me
    ElseIf mySrc = "Branch Quantity" Then
        BranchSQTY.Show
        Unload Me
    ElseIf mySrc = "Branch Identifying" Then
        BranchIdentifying.Show
        Unload Me
    ElseIf mySrc = "Replenishment StockTransfer" Then
        ReplStockTransfer.Show
        Unload Me
    End If
    
End Sub

Private Sub cmdNext_Click()
    cmdNext.Tag = "2"
    cmdNext.Visible = False
    cmdPrev.Left = cmdNext.Left
    cmdPrev.Top = cmdNext.Top
    cmdPrev.Visible = True
    dTo.Visible = True
    cmdRefresh.Visible = True
    cmdRefresh.Left = cmdPrev.Left + cmdPrev.Width + dTo.Width + 100
    dTo.Value = dFrom.Value
End Sub

Private Sub cmdPrev_Click()
    cmdNext.Tag = "1"
    cmdNext.Visible = True
    cmdPrev.Visible = False
    dTo.Visible = False
    cmdRefresh.Left = cmdNext.Left + cmdNext.Width
    Call RefreshRecord(1)
End Sub
Public Sub RefreshRecord(dCode As Integer)
    If mySrc = "Replenishment StockTransfer" Then
        If dCode = 1 Then
                DoEvents
                Call modPMTransfer.sp_StockTransfer_summary(Format$(dFrom.Value, "yyyy-mm-dd"), Format$(dTo.Value, "yyyy-mm-dd"), dCode, mainRs)
            Else
                DoEvents
                Call modPMTransfer.sp_StockTransfer_summary(Format$(dFrom.Value, "yyyy-mm-dd"), Format$(dTo.Value, "yyyy-mm-dd"), dCode, mainRs)
            End If
        If mainRs.State <> 1 Then Exit Sub
    
    If mainRs.RecordCount <= 0 Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    With mainRs
        If .RecordCount > 0 Then
            .MoveFirst
            flxviewSummary.Rows = .RecordCount + 1
            flxviewSummary.Cols = .Fields.Count + 1
    
            For i = 1 To flxviewSummary.Rows - 1
                flxviewSummary.TextMatrix(i, 1) = Format(!TRANSDATE, "")
                flxviewSummary.TextMatrix(i, 2) = Format(!DATEFINAL, "")
                flxviewSummary.TextMatrix(i, 3) = !TRANSFERID
                flxviewSummary.TextMatrix(i, 4) = !RECBRANCH
                flxviewSummary.TextMatrix(i, 5) = Format(!Address, "")
                flxviewSummary.TextMatrix(i, 6) = !PONO
                flxviewSummary.TextMatrix(i, 7) = !Status
                flxviewSummary.TextMatrix(i, 8) = Format(!Remarks, "")
                flxviewSummary.TextMatrix(i, 9) = Format(!CHECKER, "")
                flxviewSummary.TextMatrix(i, 10) = Format(!DriverName, "")
                flxviewSummary.TextMatrix(i, 11) = !branchCode
                flxviewSummary.TextMatrix(i, 12) = !BRANCHSPOID
                mainRs.MoveNext
            Next i
    
        End If
      
    End With
    
    If mainRs.State <> 0 Then
        Screen.MousePointer = vbDefault
        mainRs.Close
    End If
    
    End If
End Sub
Private Sub cmdPrint_Click()
           
    If mySrc = "Data Entry" Then
        Qry = "Call sp_partsmatrix.sp_dataentry_summary('" & sysBranchCode & "');"
        frmPrn.QryPrn = Qry
        frmPrn.ParamValue1 = SysBranch
        frmPrn.ParamValue2 = Format(lblviewSummary.Caption, "")
        frmPrn.ParamValue3 = Format(SysAddress, "")
        frmPrn.MySource = "Data Entry Summary"
        frmPrn.Show
    Else
        MsgBox "Not yet available", vbInformation + vbOKOnly, "Print"
    End If

End Sub

Private Sub cmdRefresh_Click()
If cmdPrev.Visible = False Then
    Call RefreshRecord(1)
Else
    Call RefreshRecord(2)
End If
End Sub

Private Sub cmdShow_Click()
    If mySrc = "Data Entry" Then
    
        With flxviewSummary
            Call DataEntry.ShowRecord(Val(.TextMatrix(ActRow, 11)))
        End With
        Unload Me
    ElseIf mySrc = "Branch Quantity" Then
    
        With flxviewSummary
            Call BranchSQTY.ShowRecord(Val(.TextMatrix(ActRow, 11)), .TextMatrix(ActRow, 10))
        End With
        Unload Me
        
    ElseIf mySrc = "Branch Identifying" Then
    
        With flxviewSummary
            BranchIdentifying.mySrc = "From Summary"
            Call BranchIdentifying.getREFID(Val(.TextMatrix(ActRow, 10)))
            Call BranchIdentifying.ShowRecord(branchCode, Val(.TextMatrix(ActRow, 9)), "TRUE")
        End With
        Unload Me
    ElseIf mySrc = "Replenishment StockTransfer" Then
    
        With flxviewSummary
            Call ReplStockTransfer.getSTFID(Val(.TextMatrix(ActRow, 3)))
            Call ReplStockTransfer.ShowRecord(.TextMatrix(ActRow, 11), .TextMatrix(ActRow, 6), Val(.TextMatrix(ActRow, 12)))
        End With
        Unload Me
        
    End If
    
End Sub





Private Sub flxviewSummary_EnterCell()
    ActRow = flxviewSummary.Row
End Sub

Private Sub Form_Load()
    
'*********** Code Update: 03/14/2017 ****************
    Timer1.Enabled = True


'********** End of Code Update: 03/14/2017 ****************
    
    ActRow = 1
    cbBranch.Visible = False
    lblBranch.Visible = False
End Sub

Private Sub ShowCommands(CmdType As Integer)
    Dim TmpLeft As Integer
    
    Select Case CmdType
        Case 1  'Show Form, Print Summary, Exit Report
            cmdExit.Left = cmdPrint.Left
            cmdPrint.Visible = False
        Case 2  'Show Form, Print Summary, Exit Report
           cmdShow.Visible = True
           cmdPrint.Visible = True
           cmdExit.Visible = True
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mainRs.State = adStateOpen Then mainRs.Close
End Sub

Private Sub Form_Resize()
    Call RXResizer1_ResizeComplete
End Sub



Private Sub RXResizer1_ResizeComplete()
        
    Call viewSummFlxChanger
    
'    'Label Top
'    lblviewSummary.Top = 600
'
'    'Label Left
'    lblviewSummary.Left = 120
'
'    'frmeSummary(Frame) Width
'    frmeSummary.Width = Me.Width - 540
'
'    'frmeSummary(Frame) Height
'    frmeSummary.Height = Me.Height - 1545
'
'    'FrmeSummary(Frame) Button Left
'    btnFind.Left = frmeSummary.Width - 3495
'
'    'FrmeSummary(Frame) Button Width
'    btnFind.Width = 1335
'
'    'FrmeSummary(Frame) Text Left
'    txtFind.Left = frmeSummary.Width - 2055
'
'    'FrmeSummary(Frame) Text Width
'    txtFind.Width = 1935
'
'    'FrmeSummary(Frame) FlexGrid Width
'    flxviewSummary.Width = frmeSummary.Width - 240
'
'    'FrmeSummary(Frame) FlexGrid Height
'    flxviewSummary.Height = frmeSummary.Height - 840
'
'    'FrmeSummary(Frame) Label Left
'    lblBranch.Left = 120
'
'    'FrmeSummary(Frame) Label Width
'    lblBranch.Width = 975
'
'    'FrmeSummary(Frame) ComboBox Left
'    cbBranch.Left = 1200
'
'    'FrmeSummary(Frame) ComboBox Width
'    cbBranch.Width = 4215
'
    
End Sub



Private Sub Timer1_Timer()
    If mySrc = "Data Entry" Then
        TCol = 12
        Call ShowCommands(2)
    ElseIf mySrc = "Branch Quantity" Then
        TCol = 11
        Call ShowCommands(1)
    ElseIf mySrc = "Branch Identifying" Then
        TCol = 9
        Call ShowCommands(1)
    ElseIf mySrc = "Replenishment StockTransfer" Then
        TCol = 13
        dFrom.Value = Format(Date, "mm/dd/yyyy")
        dTo.Value = Format(Date, "mm/dd/yyyy")
        Call ShowCommands(1)
        Call RefreshRecord(1)
    End If
    
    Call viewSummary
    Call viewSummFlxChanger
    Timer1.Enabled = False
End Sub

Private Sub txtFind_Change()
fRow = 1
fCol = 1
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFind, KeyAscii, ValidChars)
    
    If KeyAscii = 13 Then
        Call btnFind_Click
    End If
End Sub
