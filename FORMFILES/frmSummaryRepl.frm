VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSummaryRepl 
   Caption         =   "View Summary"
   ClientHeight    =   9150
   ClientLeft      =   120
   ClientTop       =   255
   ClientWidth     =   14505
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   14505
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSConfirmTab 
      Height          =   1575
      Left            =   14760
      TabIndex        =   25
      Tag             =   "XYF"
      Top             =   1800
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   2778
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Confirmation"
      TabPicture(0)   =   "frmSummaryRepl.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCancel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdOK"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtConfirmDate"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin MSMask.MaskEdBox txtConfirmDate 
         Height          =   375
         Left            =   2120
         TabIndex        =   26
         Tag             =   "XYF"
         Top             =   420
         Width           =   1575
         _ExtentX        =   2778
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
      Begin lvButton.lvButtons_H cmdOK 
         Height          =   375
         Left            =   840
         TabIndex        =   27
         Tag             =   "XYF"
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "&CONFIRM"
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
         Image           =   "frmSummaryRepl.frx":001C
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   375
         Left            =   2040
         TabIndex        =   28
         Tag             =   "XYF"
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "&CANCEL"
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
         Image           =   "frmSummaryRepl.frx":05B6
         cBack           =   -2147483633
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Change Confirmation Date:"
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
         Left            =   120
         TabIndex        =   29
         Tag             =   "XYF"
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Timer tmrLoadData 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5880
      Top             =   120
   End
   Begin VB.Frame frmeSummary 
      Height          =   7335
      Left            =   120
      TabIndex        =   1
      Tag             =   "XY"
      Top             =   1800
      Width           =   14295
      Begin MSComctlLib.ProgressBar pBar 
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Tag             =   "XY"
         Top             =   7680
         Visible         =   0   'False
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxviewSummary 
         Height          =   6975
         Left            =   120
         TabIndex        =   2
         Tag             =   "XY"
         Top             =   240
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   12303
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
   End
   Begin lvButton.lvButtons_H cmdShow 
      Height          =   375
      Left            =   120
      TabIndex        =   3
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
      Image           =   "frmSummaryRepl.frx":0B50
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   3840
      TabIndex        =   4
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
      Image           =   "frmSummaryRepl.frx":10EA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
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
      Image           =   "frmSummaryRepl.frx":1684
      cBack           =   -2147483633
   End
   Begin VB.Frame frmStockTransfer 
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Tag             =   "XY"
      Top             =   840
      Visible         =   0   'False
      Width           =   7215
      Begin MSComCtl2.DTPicker dFrom 
         Height          =   315
         Left            =   1560
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   96141313
         CurrentDate     =   39322
      End
      Begin MSComCtl2.DTPicker dTo 
         Height          =   315
         Left            =   3840
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   96141313
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdPrev 
         Height          =   375
         Left            =   3360
         TabIndex        =   16
         ToolTipText     =   "Hide Date End"
         Top             =   360
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
         Image           =   "frmSummaryRepl.frx":1C1E
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   375
         Left            =   5640
         TabIndex        =   17
         ToolTipText     =   "Execute"
         Top             =   360
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
         Image           =   "frmSummaryRepl.frx":21B8
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   375
         Left            =   3360
         TabIndex        =   14
         ToolTipText     =   "Show Date End"
         Top             =   360
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
         Image           =   "frmSummaryRepl.frx":2752
         cBack           =   -2147483633
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
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Frame frmBranchIdentfy 
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Tag             =   "XY"
      Top             =   840
      Visible         =   0   'False
      Width           =   7215
      Begin VB.ComboBox cbBranch 
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label lblBranch 
         Caption         =   "BRANCH:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame frmSQTY 
      Height          =   975
      Left            =   120
      TabIndex        =   20
      Tag             =   "XY"
      Top             =   840
      Visible         =   0   'False
      Width           =   7215
      Begin VB.ComboBox cbSQTYBrand 
         Height          =   315
         Left            =   5280
         TabIndex        =   24
         Tag             =   "XY"
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox cboSQTYBranch 
         Height          =   315
         Left            =   840
         TabIndex        =   22
         Tag             =   "XY"
         Top             =   360
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label lblSQTYBrand 
         Caption         =   "BRAND:"
         Height          =   255
         Left            =   4680
         TabIndex        =   23
         Tag             =   "XY"
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblSQTYBranch 
         Caption         =   "BRANCH:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Tag             =   "XY"
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Frame frmFind 
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Tag             =   "XY"
      Top             =   840
      Width           =   14295
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   12120
         TabIndex        =   11
         Tag             =   "XY"
         Top             =   300
         Width           =   1935
      End
      Begin lvButton.lvButtons_H btnFind 
         Height          =   435
         Left            =   10920
         TabIndex        =   10
         Tag             =   "XY"
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
         Caption         =   "Find Next"
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
         Image           =   "frmSummaryRepl.frx":2CEC
         cBack           =   -2147483633
      End
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   8040
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14745
      m_MDI_Height    =   9720
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
      Tag             =   "XY"
      Top             =   600
      Width           =   4095
   End
   Begin VB.Menu mnuSum 
      Caption         =   "Summary"
      Visible         =   0   'False
      Begin VB.Menu mnuConfirm 
         Caption         =   "Confirm SP"
      End
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
Dim arrBRAND() As String
Dim BranchCode As String
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
'************* CODE UPDATED: 04/21/17 *********************
If mySrc = "Branch Identifying" Then
    'Get Data from database
    DoEvents
    Call modPMBranchIdentifying.sp_get_BranchSumm(mainRs)
    If mainRs.State = 1 Then
        Call modObjTools.addToListObjectWithArrayGivenField(cbBranch, mainRs, arrBRANCH, "BRANCHCODE", "BRANCHNAME")
    End If
    If mainRs.State <> 0 Then mainRs.Close
    
ElseIf mySrc = "Branch Quantity" Then
    'Get Data from database
    DoEvents
    Call modPMBranchSQTY.sp_WHbrnchQty_getBranch(mainRs)
    If mainRs.State = 1 Then
       Call modObjTools.addToListObjectWithArrayGivenField(cboSQTYBranch, mainRs, arrBRANCH, "BRANCHCODE", "BRANCHNAME")
    End If
    If mainRs.State <> 0 Then mainRs.Close
End If
'********* END CODE UPDATED: 04/21/17 *********************
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub

Public Sub CboBrandDisplay()
On Error GoTo er
    'Get Data from database
    DoEvents
    Call modPMBranchSQTY.get_Brand(BranchCode, mainRs)
    If mainRs.State = 1 Then
        Call modObjTools.addToListObjectWithArrayGivenField(cbSQTYBrand, mainRs, arrBRAND, "nulls", "PMBRAND")
    End If
    If mainRs.State <> 0 Then mainRs.Close
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboBrandDisplay()")
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

Private Sub cbBranch_Click()
    frm_OnProcess2.Show
    flxviewSummary.Clear 'used to refresh the flexgrid
    Call viewSummFlxChanger 'used to refresh the flexgrid
    DoEvents
    
'********** CODE UPDATED: 04/11/17 ***************
    pBar.Visible = True
'****** END CODE UPDATED: 04/11/17 ***************

    BranchCode = arrBRANCH(cbBranch.ListIndex, 0)
    
    'Call ShowRecord(branchCode)
    
    Call Unload(frm_OnProcess2)
End Sub

'Public Sub ShowRecord(MyBranchCode As String)
'On Error GoTo er
'    Screen.MousePointer = vbHourglass
'
'    DoEvents
'    Call modPMBranchIdentifying.sp_brnchIdentify_summary(MyBranchCode, mainRs)
'
'    If mainRs.State <> 1 Then Exit Sub
'
''************* CODE UPDATED: 3/14/17 *****************
'    pBar.Visible = True
'    pBar.Value = 0
''********* END CODE UPDATED: 3/14/17 *****************
'
'    If mainRs.RecordCount <= 0 Then
'
''************* CODE UPDATED: 3/14/17 *****************
'        pBar.Visible = False
'        pBar.Value = 0
'        MsgBox "No Records Available", vbInformation + vbOKOnly, "Summary"
''********* END CODE UPDATED: 3/14/17 *****************
'
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'
'    With mainRs
'        If .RecordCount > 0 Then
'            .MoveFirst
'            flxviewSummary.Rows = .RecordCount + 1
'            flxviewSummary.Cols = .Fields.Count + 1
'        '************* CODE UPDATED: 3/14/17 *****************
'            pBar.Max = .RecordCount
'        '********* END CODE UPDATED: 3/14/17 *****************
'            For i = 1 To flxviewSummary.Rows - 1
'                flxviewSummary.TextMatrix(i, 1) = !branchName
'                flxviewSummary.TextMatrix(i, 2) = !MODELBRAND
'                flxviewSummary.TextMatrix(i, 3) = !TOTALPN
'                flxviewSummary.TextMatrix(i, 4) = !TOTALALQ
'                flxviewSummary.TextMatrix(i, 5) = !RefNo
'                flxviewSummary.TextMatrix(i, 6) = !DATEPREPARED
'                flxviewSummary.TextMatrix(i, 7) = !DATEFINAL
'                flxviewSummary.TextMatrix(i, 8) = IIf(IsNull(!PROCESSEDBY), "", !PROCESSEDBY)
'                flxviewSummary.TextMatrix(i, 9) = IIf(IsNull(!Status), "", !Status)
'                flxviewSummary.TextMatrix(i, 10) = !REFID
'                flxviewSummary.TextMatrix(i, 11) = !BCODE
'                mainRs.MoveNext
'            '************* CODE UPDATED: 3/14/17 *****************
'                pBar.Value = pBar.Value + 1
'            '********* END CODE UPDATED: 3/14/17 *****************
'            Next i
'
'            '************* CODE UPDATED: 06/09/17 *****************
'            With flxviewSummary
'            For i = 1 To .Rows - 1
'                .TextMatrix(i, 6) = Format(.TextMatrix(i, 6), "mm/dd/yyyy")
'                .TextMatrix(i, 7) = Format(.TextMatrix(i, 7), "mm/dd/yyyy")
'
'                If .TextMatrix(i, 9) = "NEW" Then
'                        .Col = 1: .Row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
'                        .Col = 2
'                End If
'                If .TextMatrix(i, 9) = "ALLOCATED" Then
'                        .Col = 1: .Row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
'                        .Col = 2
'                End If
'            Next i
'            End With
'            '********* END CODE UPDATED: 06/09/17 *****************
'
'        End If
'
'    End With
'
'    If mainRs.State <> 0 Then
'        Screen.MousePointer = vbDefault
'        mainRs.Close
'    End If
'
''************* CODE UPDATED: 3/14/17 *****************
'    pBar.Visible = False
'    pBar.Value = 0
''********* END CODE UPDATED: 3/14/17 *****************
'Exit Sub
'er:
'    DoEvents
'    Call modObjTools.ErrorMessage(Err.Description, "Error->SummaryShowRecord()")
'End Sub

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
            .ColWidth(10) = (.Width - .ColWidth(0)) * 0.1 'LASTUPDATE
            .ColWidth(11) = 0 '(.Width - .ColWidth(0)) * 0.1 'ID
            lblviewSummary.Caption = "PARTS MATRIX DATA ENTRY SUMMARY"
        End With
        
    ElseIf mySrc = "Branch Quantity" Then
    
        With flxviewSummary
            .Cols = 14
           '                       1         2          3          4           5            6            7          8        9       10         11
            .FormatString = " |<BRANCH |^MC BRAND |^MC MODEL |^TOTAL PN |^TOTAL STDQ |^TOTAL AMOUNT |^TOTAL PG |^EXCESS |^LACKING |^EMPTY |^UPDATED BY"
        
            .ColWidth(0) = 200
            .ColWidth(1) = (.Width - .ColWidth(0)) * 0.15 'BRANCH
            .ColWidth(2) = (.Width - .ColWidth(0)) * 0.11 'BRAND
            .ColWidth(3) = (.Width - .ColWidth(0)) * 0.11 'MODEL
            .ColWidth(4) = (.Width - .ColWidth(0)) * 0.1 'TOTAL PN
            .ColWidth(5) = (.Width - .ColWidth(0)) * 0.1 'TOTAL STDQ
            .ColWidth(6) = (.Width - .ColWidth(0)) * 0.1 'TOTAL AMOUNT
            .ColWidth(7) = (.Width - .ColWidth(0)) * 0.1 'TOTAL PARTSGROUP
            .ColWidth(8) = (.Width - .ColWidth(0)) * 0.1 'EXCESS
            .ColWidth(9) = (.Width - .ColWidth(0)) * 0.1 'LACKING
            .ColWidth(10) = (.Width - .ColWidth(0)) * 0.1 'EMPTY
            .ColWidth(11) = (.Width - .ColWidth(0)) * 0.2 'UPDATEDBY
            .ColWidth(12) = 0 '(.Width - .ColWidth(0)) * 0.11 'BCODE
            .ColWidth(13) = 0 '(.Width - .ColWidth(0)) * 0.03 'PMATRIXID
            lblviewSummary.Caption = "BRANCH QUANTITY SUMMARY"
    '**************** CODE UPDATED: 04/19/17 ****************************
            frmSQTY.Visible = True
            lblSQTYBranch.Visible = True
            lblSQTYBrand.Visible = True
            cboSQTYBranch.Visible = True
            cbSQTYBrand.Visible = True
    '************ END CODE UPDATED: 04/19/17 ****************************
        End With
        
    ElseIf mySrc = "Branch Replenishment Order" Then
    
        With flxviewSummary
            .Cols = 7
          '                           1            2           3           4               5          6
            .FormatString = " |<BRANCH NAME |<REFERENCE NO. |<ORDER DATE |^TOTAL PN |^TOTAL ORDER |^STATUS"
        
            .ColWidth(0) = 200
            .ColWidth(1) = (.Width - .ColWidth(0)) * 0.15 'BRANCH NAME
            .ColWidth(2) = (.Width - .ColWidth(0)) * 0.1 'REFERENCE NO
            .ColWidth(3) = (.Width - .ColWidth(0)) * 0.1 'ORDER DATE
            .ColWidth(4) = (.Width - .ColWidth(0)) * 0.1 'TOTAL PN
            .ColWidth(5) = (.Width - .ColWidth(0)) * 0.1 'TOTAL ORDER
            .ColWidth(6) = (.Width - .ColWidth(0)) * 0.1 'STATUS
            '.ColWidth(7) = (.Width - .ColWidth(0)) * 0.1 'DATE FINAL
            '.ColWidth(8) = (.Width - .ColWidth(0)) * 0.2 'PROCESSED BY
            '.ColWidth(9) = (.Width - .ColWidth(0)) * 0.1 'STATUS
            '.ColWidth(10) = 0 '(.Width - .ColWidth(0)) * 0.1 'REFID
            '.ColWidth(11) = 0 '(.Width - .ColWidth(0)) * 0.1 'BCODE
            lblviewSummary.Caption = "BRANCH REPLENISHMENT ORDER SUMMARY"
'************* CODE UPDATED: 3/14/17 *****************
            frmStockTransfer.Visible = True
            Label1.Visible = True
            Dfrom.Visible = True
            Dto.Visible = True
            cmdNext.Visible = True
            cmdRefresh.Visible = True
            cmdPrev.Visible = True
'********* END CODE UPDATED: 3/14/17 *****************
        End With
    
    ElseIf mySrc = "Replenishment StockTransfer" Then
    
        With flxviewSummary
            .Cols = 17
          '                           1        2               3           4           5               6          7         8         9        10       11         12       13
            .FormatString = " |^REF DATE |^FINAL DATE |^DELIVERY DATE |^STF NO |<RECEIVING BRANCH |<ADDRESS |^REPL REF# |^TTL PN |^TTL QTY |^STATUS |^REMARKS |^CHECKER |^DRIVER"
        
            .ColWidth(0) = 200
            .ColWidth(1) = (.Width - .ColWidth(0)) * 0.08 'REF DATE
            .ColWidth(2) = (.Width - .ColWidth(0)) * 0.08 'FINAL DATE
            .ColWidth(3) = (.Width - .ColWidth(0)) * 0.1 'DELIVERY DATE
            .ColWidth(4) = (.Width - .ColWidth(0)) * 0.08 'STF NO
            .ColWidth(5) = (.Width - .ColWidth(0)) * 0.23 'RECEIVING BRANCH
            .ColWidth(6) = (.Width - .ColWidth(0)) * 0.23 'ADDRESS
            .ColWidth(7) = (.Width - .ColWidth(0)) * 0.1 'REPL REF#
            .ColWidth(8) = (.Width - .ColWidth(0)) * 0.1 'TTL PN
            .ColWidth(9) = (.Width - .ColWidth(0)) * 0.1 'TTL QTY
            .ColWidth(10) = (.Width - .ColWidth(0)) * 0.1 'STATUS
            .ColWidth(11) = (.Width - .ColWidth(0)) * 0.1 'REMARKS
            .ColWidth(12) = (.Width - .ColWidth(0)) * 0.15 'CHECKER
            .ColWidth(13) = (.Width - .ColWidth(0)) * 0.15 'DRIVER
            .ColWidth(14) = 0 '(.Width - .ColWidth(0)) * 0.03 'BRANCHCODE
            .ColWidth(15) = 0 '(.Width - .ColWidth(0)) * 0.03 'BRANCHSPOID
            .ColWidth(16) = 0 '(.Width - .ColWidth(0)) * 0.03 'TRANSFERID
            lblviewSummary.Caption = "REPLENISHMENT STOCK TRANSFER SUMMARY"
        '************* CODE UPDATED: 3/14/17 *****************
            frmStockTransfer.Visible = True
            Label1.Visible = True
            Dfrom.Visible = True
            Dto.Visible = True
            cmdNext.Visible = True
            cmdRefresh.Visible = True
            cmdPrev.Visible = True
        '********* END CODE UPDATED: 3/14/17 *****************
        End With
        
    ElseIf mySrc = "Branch Picking" Then
    
        With flxviewSummary
            .Cols = 12
          '                       1            2             3          4            5              6            7              8           9
            .FormatString = " |<BRANCH |<REFERENCE NO. |^TOTAL PN |^TOTAL ALQ  |^TIME START |^TIME FINISH |^DATE PRINTED |<PROCESSED BY |^STATUS"
        
            .ColWidth(0) = 200
            .ColWidth(1) = (.Width - .ColWidth(0)) * 0.2 'BRANCH NAME
            .ColWidth(2) = (.Width - .ColWidth(0)) * 0.15 'REFERENCE NO.
            .ColWidth(3) = (.Width - .ColWidth(0)) * 0.08 'TOTAL PN
            .ColWidth(4) = (.Width - .ColWidth(0)) * 0.08 'TOTAL ALQ
            .ColWidth(5) = (.Width - .ColWidth(0)) * 0.1 'TIME START
            .ColWidth(6) = (.Width - .ColWidth(0)) * 0.1 'TIME FINISH
            .ColWidth(7) = (.Width - .ColWidth(0)) * 0.1 'DATE PRINTED
            .ColWidth(8) = (.Width - .ColWidth(0)) * 0.2 'PROCESSED BY
            .ColWidth(9) = (.Width - .ColWidth(0)) * 0.1 'STATUS
            .ColWidth(10) = 0 '(.Width - .ColWidth(0)) * 0.1 'PICK ID
            .ColWidth(11) = 0 '(.Width - .ColWidth(0)) * 0.2 'BRANCH CODE
            lblviewSummary.Caption = "BRANCH PICKING SUMMARY"
        End With
        
    End If
End Sub

Public Sub viewSummary()
    
On Error GoTo er
    
    frm_OnProcess2.Show
    
    If mySrc = "Data Entry" Then
        
        Screen.MousePointer = vbHourglass
        
        DoEvents
        Call modPMDataEntry.sp_dataentry_summary(mainRs)
        
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
            '************* CODE UPDATED: 3/14/17 *****************
                pBar.Max = .RecordCount
            '********* END CODE UPDATED: 3/14/17 *****************
                flxviewSummary.Rows = .RecordCount + 1
                'flxviewSummary.Cols = .Fields.Count + 1
        
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
        
    ElseIf mySrc = "Branch Quantity" Then

'*************** CODE UPDATED: 04/19/17 **********************
        Call CboDisplay
        DoEvents
        Call modPMBranchSQTY.sp_brnchQty_summary(1, mainRs)
        
        Call BranchSQTYSUMMARY
'*********** END CODE UPDATED: 04/19/17 **********************
'*************** CODE UPDATED: 04/29/17 **********************
    ElseIf mySrc = "Branch Identifying" Then
        Call CboDisplay
        
    ElseIf mySrc = "Branch Picking" Then
        Screen.MousePointer = vbHourglass
        pBar.Visible = True
        pBar.Value = 0
        
        DoEvents
        Call modPMBranchPicking.sp_brnchPicking_PickingListSumm(mainRs)
            With mainRs
                If .RecordCount = 0 Then
                    flxviewSummary.Clear
                    flxviewSummary.Rows = 2
                    Call RXResizer1_ResizeComplete
                    Screen.MousePointer = vbDefault
                    Call Unload(frm_OnProcess2)
                    Exit Sub
                End If
            End With
            Set flxviewSummary.DataSource = mainRs
            Screen.MousePointer = vbDefault
            
            '************* CODE UPDATED: 06/09/17 *****************
            With flxviewSummary
            For i = 1 To .Rows - 1
                .TextMatrix(i, 5) = Format(.TextMatrix(i, 5), "hh:mm AM/PM")
                .TextMatrix(i, 6) = Format(.TextMatrix(i, 6), "hh:mm AM/PM")
                .TextMatrix(i, 7) = Format(.TextMatrix(i, 7), "mm/dd/yyyy")
                
                If .TextMatrix(i, 9) = "PENDING" Then
                        .Col = 1: .row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                        .Col = 2
                End If
                If .TextMatrix(i, 9) = "CANCELLED" Then
                        .Col = 1: .row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
                        .Col = 2
                End If
            Next i
            End With
            '********* END CODE UPDATED: 06/09/17 *****************
    End If
'*********** END CODE UPDATED: 04/29/17 **********************
'************* CODE UPDATED: 3/14/17 *****************
    pBar.Visible = False
    pBar.Value = 0
    Call Unload(frm_OnProcess2)
'********* END CODE UPDATED: 3/14/17 *****************
Exit Sub
er:
    Call Unload(frm_OnProcess2)
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
                    .Col = 1: .row = i
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

Private Sub cbBranch_KeyDown(KeyCode As Integer, Shift As Integer)
'************** CODE UPDATED: 04/07/17 **************************
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
'********** END CODE UPDATED: 04/07/17 **************************
End Sub

Private Sub cbBranch_KeyPress(KeyAscii As Integer)
'************** CODE UPDATED: 04/07/17 **************************
Dim cbolistIndex As Integer
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cbBranch, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        cbolistIndex = modObjTools.getListIndex(cbBranch, cbBranch.Text)
        
        If cbolistIndex < 0 Then
            cbBranch.SetFocus
        Else
            cbBranch.ListIndex = cbolistIndex
        End If
        
        Call cbBranch_Click
    End If
'********** END CODE UPDATED: 04/07/17 **************************
End Sub

''**************** CODE UPDATED: 04/19/17 ****************************
Private Sub cboSQTYBranch_Change()
On Error GoTo er
    If StatSrch Then
            StatSrch = False
            Exit Sub
    End If

    With cboSQTYBranch
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboSQTYBranch-Change()")
End Sub
'************ END CODE UPDATED: 04/19/17 ****************************

'**************** CODE UPDATED: 04/19/17 ****************************
Private Sub cboSQTYBranch_Click()
    frm_OnProcess2.Show
    flxviewSummary.Clear 'used to refresh the flexgrid
    Call viewSummFlxChanger 'used to refresh the flexgrid
    
    DoEvents
    BranchCode = arrBRANCH(cboSQTYBranch.ListIndex, 0)
'********** CODE UPDATED: 04/19/17 ***************
    If cboSQTYBranch.Text = "<SELECT ALL>" Then
        Call modPMBranchSQTY.sp_brnchQty_summary(1, mainRs)
        BranchCode = sysBranchCode
        Call BranchSQTYSUMMARY
        cbSQTYBrand.Clear
    Else
        Call modPMBranchSQTY.sp_brnchQty_summary_wBRAND(BranchCode, cbSQTYBrand.Text, mainRs)
        Call BranchSQTYSUMMARY
        Call CboBrandDisplay
    End If
    Call Unload(frm_OnProcess2)
'****** END CODE UPDATED: 04/19/17 ***************
End Sub
'************ END CODE UPDATED: 04/19/17 ****************************


Private Sub cboSQTYBranch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboSQTYBranch_KeyPress(KeyAscii As Integer)
Dim cbolistIndex As Integer
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboSQTYBranch, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        cbolistIndex = modObjTools.getListIndex(cboSQTYBranch, cboSQTYBranch.Text)
        
        If cbolistIndex < 0 Then
            cboSQTYBranch.SetFocus
        Else
            cboSQTYBranch.ListIndex = cbolistIndex
        End If
        
        Call cboSQTYBranch_Click
    End If
End Sub

'**************** CODE UPDATED: 04/19/17 ****************************
Private Sub cbSQTYBrand_Change()
On Error GoTo er
    If StatSrch Then
            StatSrch = False
            Exit Sub
    End If

    With cbSQTYBrand
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->cbSQTYBrand-Change()")
End Sub
'************ END CODE UPDATED: 04/19/17 ****************************

'**************** CODE UPDATED: 04/19/17 ****************************
Private Sub cbSQTYBrand_Click()
    frm_OnProcess2.Show
    flxviewSummary.Clear 'used to refresh the flexgrid
    Call viewSummFlxChanger 'used to refresh the flexgrid
    
    DoEvents
    Call modPMBranchSQTY.sp_brnchQty_summary_wBRAND(BranchCode, cbSQTYBrand.Text, mainRs)
    Call BranchSQTYSUMMARY
    
    Call Unload(frm_OnProcess2)
End Sub
'************ END CODE UPDATED: 04/19/17 ****************************

Private Sub cbSQTYBrand_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cbSQTYBrand_KeyPress(KeyAscii As Integer)
Dim cbolistIndex As Integer
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cbSQTYBrand, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        cbolistIndex = modObjTools.getListIndex(cbSQTYBrand, cbSQTYBrand.Text)
        
        If cbolistIndex < 0 Then
            cbSQTYBrand.SetFocus
        Else
            cbSQTYBrand.ListIndex = cbolistIndex
        End If
        
        Call cbSQTYBrand_Click
    End If
End Sub

Private Sub cmdCancel_Click()
SSConfirmTab.Visible = False
End Sub

Private Sub cmdExit_Click()
    If mySrc = "Data Entry" Then
        DataEntry.Show
        Unload Me
    ElseIf mySrc = "Branch Quantity" Then
        BranchSQTY.Show
        Unload Me
'    ElseIf mySrc = "Branch Identifying" Then
'        BranchIdentifying.Show
'        Unload Me
    ElseIf mySrc = "Replenishment StockTransfer" Then
        ReplStockTransfer.Show
        Unload Me
    ElseIf mySrc = "Branch Picking" Then
        Unload Me
    End If
    
End Sub

Private Sub cmdNext_Click()
    cmdNext.Tag = "2"
    cmdNext.Visible = False
    cmdPrev.Left = cmdNext.Left
    cmdPrev.Top = cmdNext.Top
    cmdPrev.Visible = True
    Dto.Visible = True
    cmdRefresh.Visible = True
    cmdRefresh.Left = cmdPrev.Left + cmdPrev.Width + Dto.Width + 100
    Dto.Value = Dfrom.Value
End Sub

Private Sub cmdOK_Click()
    Dim rsTransfer As New ADODB.Recordset
    Dim ans As Integer, i As Integer, myDateFinal As String, STFID As Long
    Dim Qry As String, CanRemarks As String, MyRef As String
        
    ans = MsgBox("Do you want to Confirm ths Transfer?", vbQuestion + vbYesNo, "Validation")
    If ans = vbNo Then Exit Sub
    
    If flxviewSummary.TextMatrix(ActRow, 2) <> "" Then
        MyRef = "TR#" & Val(flxviewSummary.TextMatrix(ActRow, 4))
        'Qry = "SELECT SPITR.SPAREPARTID, SPITR.OUT From SPITR WHERE (((SPITR.REFERENCE)='" & MyRef & "'));"
        'mySQL
        Qry = "SELECT stf.TRANSFERID,stf.DELIVEREDDATE From stocktransfer as stf WHERE stf.TRANSFERID='" & Val(flxviewSummary.TextMatrix(ActRow, 4)) & "';"
        
        With rsTransfer
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    STFID = !TRANSFERID
                    If flxviewSummary.TextMatrix(ActRow, 3) = "" Then
                        LYRcon.Execute "UPDATE STOCKTRANSFER SET STOCKTRANSFER.DELIVEREDDATE = '" & Format(txtConfirmDate.Text, "yyyy-mm-dd") & "', STOCKTRANSFER.UPDATEDBY = " & UserID & ", STOCKTRANSFER.LASTUPDATE = '" & Format(Now, "yyyy-mm-dd") & "' " & _
                            "WHERE STOCKTRANSFER.TRANSFERID=" & STFID & ";"
                    Else
                        MsgBox "Spareparts Transfer Ref.# " & flxviewSummary.TextMatrix(ActRow, 4) & " is already Confirm !", vbInformation + vbOKOnly, "CONFIRM"
                        .Close
                        SSConfirmTab.Visible = False
                        Exit Sub
                     End If
                    .MoveNext
                Loop
            End If
            .Close
        End With
    Else
        MsgBox "Spareparts Transfer Ref.# " & flxviewSummary.TextMatrix(ActRow, 4) & " is not yet Final !", vbInformation + vbOKOnly, "CONFIRM"
        Exit Sub
    End If
    
    MsgBox "Spareparts Transfer Ref.# " & flxviewSummary.TextMatrix(ActRow, 4) & " was Confirm!", vbInformation + vbOKOnly, "CONFIRM"
    SSConfirmTab.Visible = False
    Call cmdRefresh_Click
End Sub

Private Sub cmdPrev_Click()
    cmdNext.Tag = "1"
    cmdNext.Visible = True
    cmdPrev.Visible = False
    Dto.Visible = False
    cmdRefresh.Left = cmdNext.Left + cmdNext.Width
    Call RefreshRecord(1)
End Sub

Public Sub RefreshRecord(dCode As Integer)
    frm_OnProcess2.Show
    If mySrc = "Replenishment StockTransfer" Then
        If dCode = 1 Then
            DoEvents
            Call modPMTransfer.sp_StockTransfer_summary(Format$(Dfrom.Value, "yyyy-mm-dd"), Format$(Dto.Value, "yyyy-mm-dd"), dCode, mainRs)
        Else
            DoEvents
            Call modPMTransfer.sp_StockTransfer_summary(Format$(Dfrom.Value, "yyyy-mm-dd"), Format$(Dto.Value, "yyyy-mm-dd"), dCode, mainRs)
        End If
    ElseIf mySrc = "Branch Replenishment Order" Then
        If dCode = 1 Then
            DoEvents
            Call modPMBranchIdentifying.sp_brnchIdentify_summary(Format$(Dfrom.Value, "yyyy-mm-dd"), Format$(Dto.Value, "yyyy-mm-dd"), dCode, mainRs)
        Else
            DoEvents
            Call modPMBranchIdentifying.sp_brnchIdentify_summary(Format$(Dfrom.Value, "yyyy-mm-dd"), Format$(Dto.Value, "yyyy-mm-dd"), dCode, mainRs)
        End If
    End If
    If mainRs.State <> 1 Then
        Call Unload(frm_OnProcess2)
        Exit Sub
    End If
'************* CODE UPDATED: 3/14/17 *****************
    pBar.Visible = True
    pBar.Value = 0
'********* END CODE UPDATED: 3/14/17 *****************

    If mainRs.RecordCount <= 0 Then
    '************* CODE UPDATED: 3/14/17 *****************
        Call Unload(frm_OnProcess2)
        pBar.Visible = False
        pBar.Value = 0
    '********* END CODE UPDATED: 3/14/17 *****************
        Exit Sub
    End If
    
    If mySrc = "Replenishment StockTransfer" Then
        With mainRs
            If .RecordCount > 0 Then
                .MoveFirst
                flxviewSummary.Rows = .RecordCount + 1
                flxviewSummary.Cols = .Fields.Count + 1
            '************* CODE UPDATED: 3/14/17 *****************
                pBar.Max = .RecordCount
            '********* END CODE UPDATED: 3/14/17 *****************
                For i = 1 To flxviewSummary.Rows - 1
                    flxviewSummary.TextMatrix(i, 1) = IIf(IsNull(!TRANSDATE), "", !TRANSDATE)
                    flxviewSummary.TextMatrix(i, 2) = IIf(IsNull(!DATEFINAL), "", !DATEFINAL)
                    flxviewSummary.TextMatrix(i, 3) = IIf(IsNull(!DELIVEREDDATE), "", !DELIVEREDDATE)
                    flxviewSummary.TextMatrix(i, 4) = IIf(IsNull(!TRANSNO), "", !TRANSNO)
                    flxviewSummary.TextMatrix(i, 5) = !RECBRANCH
                    flxviewSummary.TextMatrix(i, 6) = IIf(IsNull(!Address), "", !Address)
                    flxviewSummary.TextMatrix(i, 7) = !PONO
                    flxviewSummary.TextMatrix(i, 8) = !Status
                    flxviewSummary.TextMatrix(i, 9) = IIf(IsNull(!MODELBRAND), "", !MODELBRAND)
                    flxviewSummary.TextMatrix(i, 10) = IIf(IsNull(!CHECKER), "", !CHECKER)
                    flxviewSummary.TextMatrix(i, 11) = IIf(IsNull(!DriverName), "", !DriverName)
                    flxviewSummary.TextMatrix(i, 12) = !BranchCode
                    flxviewSummary.TextMatrix(i, 13) = !BRANCHSPOID
                    flxviewSummary.TextMatrix(i, 14) = !TRANSFERID
                    mainRs.MoveNext
                '************* CODE UPDATED: 3/14/17 *****************
                    pBar.Value = pBar.Value + 1
                '********* END CODE UPDATED: 3/14/17 *****************
                Next i
            End If
        End With
        '************* CODE UPDATED: 04/03/17 *****************
        With flxviewSummary
        For i = 1 To .Rows - 1
            .TextMatrix(i, 1) = Format(.TextMatrix(i, 1), "mm/dd/yyyy")
            .TextMatrix(i, 2) = Format(.TextMatrix(i, 2), "mm/dd/yyyy")
            .TextMatrix(i, 3) = Format(.TextMatrix(i, 3), "mm/dd/yyyy")
            
            If .TextMatrix(i, 8) = "PACKING" Or .TextMatrix(i, 8) = "FINAL" Then
                    .Col = 1: .row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                    .Col = 2
            End If
            If .TextMatrix(i, 8) = "CANCELLED" Then
                    .Col = 1: .row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
                    .Col = 2
            End If
        Next i
        End With
        '********* END CODE UPDATED: 04/03/17 *****************
    ElseIf mySrc = "Branch Identifying" Then
        With mainRs
            If .RecordCount > 0 Then
                .MoveFirst
                flxviewSummary.Rows = .RecordCount + 1
                flxviewSummary.Cols = .Fields.Count + 1
            '************* CODE UPDATED: 3/14/17 *****************
                pBar.Max = .RecordCount
            '********* END CODE UPDATED: 3/14/17 *****************
                For i = 1 To flxviewSummary.Rows - 1
                    flxviewSummary.TextMatrix(i, 1) = !branchName
                    flxviewSummary.TextMatrix(i, 2) = !MODELBRAND
                    flxviewSummary.TextMatrix(i, 3) = !TOTALPN
                    flxviewSummary.TextMatrix(i, 4) = !TOTALALQ
                    flxviewSummary.TextMatrix(i, 5) = !RefNo
                    flxviewSummary.TextMatrix(i, 6) = !DATEPREPARED
                    flxviewSummary.TextMatrix(i, 7) = !DATEFINAL
                    flxviewSummary.TextMatrix(i, 8) = IIf(IsNull(!PROCESSEDBY), "", !PROCESSEDBY)
                    flxviewSummary.TextMatrix(i, 9) = IIf(IsNull(!Status), "", !Status)
                    flxviewSummary.TextMatrix(i, 10) = !REFID
                    flxviewSummary.TextMatrix(i, 11) = !bCode
                    mainRs.MoveNext
                '************* CODE UPDATED: 3/14/17 *****************
                    pBar.Value = pBar.Value + 1
                '********* END CODE UPDATED: 3/14/17 *****************
                Next i
        
                '************* CODE UPDATED: 06/09/17 *****************
                With flxviewSummary
                For i = 1 To .Rows - 1
                    .TextMatrix(i, 6) = Format(.TextMatrix(i, 6), "mm/dd/yyyy")
                    .TextMatrix(i, 7) = Format(.TextMatrix(i, 7), "mm/dd/yyyy")
            
                    If .TextMatrix(i, 9) = "NEW" Then
                            .Col = 1: .row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                            .Col = 2
                    End If
                    If .TextMatrix(i, 9) = "ALLOCATED" Then
                            .Col = 1: .row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
                            .Col = 2
                    End If
                Next i
                End With
                '********* END CODE UPDATED: 06/09/17 *****************
            End If
        End With

    End If
    
    If mainRs.State <> 0 Then
        mainRs.Close
    End If
    
'************* CODE UPDATED: 3/14/17 *****************
    pBar.Visible = False
    pBar.Value = 0
    Call Unload(frm_OnProcess2)
'********* END CODE UPDATED: 3/14/17 *****************

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
            Call BranchSQTY.ShowRecord(Val(.TextMatrix(ActRow, 13)), .TextMatrix(ActRow, 12))
        End With
        Unload Me
        
'    ElseIf mySrc = "Branch Identifying" Then
'
'        With flxviewSummary
'            'BranchIdentifying.mySrc = "From Summary"
'            BranchIdentifying.Show
'            'Call BranchIdentifying.getREFID(Val(.TextMatrix(ActRow, 11)))
'            'Call BranchIdentifying.ShowRecord(branchCode, Val(.TextMatrix(ActRow, 10)), "FALSE", .TextMatrix(ActRow, 9))
'        End With
'        Unload Me
    ElseIf mySrc = "Replenishment StockTransfer" Then
    
        With flxviewSummary
            Call ReplStockTransfer.getSTFID(Val(.TextMatrix(ActRow, 14)), .TextMatrix(ActRow, 4))
            Call ReplStockTransfer.ShowRecord(.TextMatrix(ActRow, 12), .TextMatrix(ActRow, 7), Val(.TextMatrix(ActRow, 13)), "FALSE", .TextMatrix(ActRow, 8))
        End With
        Unload Me
    ElseIf mySrc = "Branch Picking" Then
    
        With flxviewSummary
            Call BranchPicking.ShowRecord(.TextMatrix(ActRow, 11), .TextMatrix(ActRow, 2), .TextMatrix(ActRow, 9), "Pick Summary")
        End With
        Unload Me
    End If
    
End Sub

Private Sub flxviewSummary_EnterCell()
    ActRow = flxviewSummary.row
End Sub

Private Sub flxviewSummary_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'Identify transactions without cancel features and exit procedure
    If flxviewSummary.TextMatrix(ActRow, 1) = "" Or mySrc = "Data Entry" Or mySrc = "Branch Quantity" Or mySrc = "Branch Identifying" Or mySrc = "Replenishment Stocktransfer" Or mySrc = "Branch Picking" Then Exit Sub
    If Not IsDate(flxviewSummary.TextMatrix(ActRow, 1)) Then
        Exit Sub
    End If
    If Button = vbRightButton Then
'        If mySrc = "Daily Cash Collection" Then
'            If InStr(UCase$(flxSum.TextMatrix(ActRow, 3)), "CANCELLED") > 0 Then
'                Call PopupMenu(mnuCancelReason)
'
'            Else
'                mnuCancel.Caption = "Cancel O.R."
'                Call PopupMenu(mnuSum)
'
'            End If
'        ElseIf mySrc = "Purchase Return" Then
'            Call PopupMenu(mnuPurchRet)
'        ElseIf mySrc = "Spareparts Invoice" Then
'            If Val(flxSum.TextMatrix(ActRow, 10)) > 0 Then
'                mnuCancel.Caption = "Cancel " & mySrc
'                Call PopupMenu(mnuSum)
'            End If
'        Else

'            mnuCancel.Caption = "Cancel " & mySrc
'************CODE UPDATED: 08/04/2017****************
            mnuConfirm.Caption = "Confirm " & mySrc
'********END CODE UPDATED: 08/04/2017****************
            Call PopupMenu(mnuSum)
            
'        End If
    End If
End Sub


Private Sub Form_Load()

'************* CODE UPDATED: 3/14/17 *****************
    tmrLoadData.Enabled = True
'********* END CODE UPDATED: 3/14/17 *****************
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
    
'************* CODE UPDATED: 3/14/17 *****************
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
'********* END CODE UPDATED: 3/14/17 *****************
End Sub




Private Sub tmrLoadData_Timer()
'************* CODE UPDATED: 3/14/17 *****************
    If mySrc = "Data Entry" Then
        TCol = 12
        Call ShowCommands(2)
    ElseIf mySrc = "Branch Quantity" Then
        TCol = 11
        Call ShowCommands(1)
    ElseIf mySrc = "Branch Replenishment Order" Then
        TCol = 9
        Call ShowCommands(1)
    ElseIf mySrc = "Replenishment StockTransfer" Then
        TCol = 13
        Dfrom.Value = Format(Date, "mm/dd/yyyy")
        Dto.Value = Format(Date, "mm/dd/yyyy")
        Call ShowCommands(1)
        Call RefreshRecord(1)
    End If
   
    Call viewSummary
    Call viewSummFlxChanger
    
    tmrLoadData.Enabled = False
'********* END CODE UPDATED: 3/14/17 *****************
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

'*************** CODE UPDATED: 04/19/17 **********************
Public Sub BranchSQTYSUMMARY()

Screen.MousePointer = vbDefault

pBar.Visible = True
pBar.Value = 0

If mainRs.State <> 1 Then Exit Sub

If mainRs.RecordCount <= 0 Then
    MsgBox "There is no Partsmatrix was Assign in this Branch", vbInformation + vbOKOnly, "Selecting Branch"
    pBar.Visible = False
    pBar.Value = 0
    flxviewSummary.Clear
    flxviewSummary.Rows = 2
    Call RXResizer1_ResizeComplete
    Screen.MousePointer = vbDefault
    Exit Sub
End If

With mainRs
    If .RecordCount > 0 Then
        .MoveFirst
        pBar.Max = .RecordCount
        
        flxviewSummary.Rows = .RecordCount + 1
        'flxviewSummary.Cols = .Fields.Count + 1
        
     
        For i = 1 To flxviewSummary.Rows - 1
            TtlQOH = !PMTOTALQOH
            TtlPMATRIXID = !PMTOTALPN
            strExcess = IIf(IsNull(!PMTOTALEXCESS), 0, !PMTOTALEXCESS)
            If Val(strExcess) <> 0 Then
                
                If Val(strExcess) <= TtlQOH Then
                    strTtlExcess = strExcess / TtlQOH * 100 & "%"
                ElseIf strExcess >= TtlQOH Then
                    strTtlExcess = TtlQOH / strExcess * 100 & "%"
                End If
                
            Else
                strTtlExcess = "0.0%"
            End If
            
            strLacking = IIf(IsNull(!PMTOTALLACKING), 0, !PMTOTALLACKING)
            If strLacking <> 0 Then
            
                If Val(strLacking) <= !PMTOTALSTDQ Then
                    strTtlLacking = strLacking / IIf(IsNull(!PMTOTALSTDQ), 0, !PMTOTALSTDQ) * 100 & "%"
                ElseIf Val(strLacking) >= IIf(IsNull(!PMTOTALSTDQ), 0, !PMTOTALSTDQ) Then
                    strTtlLacking = IIf(IsNull(!PMTOTALSTDQ), 0, !PMTOTALSTDQ) / strLacking * 100 & "%"
                End If
                
            Else
                strTtlLacking = "0.0%"
            End If
            
            strEmpty = IIf(IsNull(!PMTOTALEMPTY), 0, !PMTOTALEMPTY)
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
            flxviewSummary.TextMatrix(i, 4) = IIf(IsNull(!PMTOTALPN), 0, !PMTOTALPN)
            flxviewSummary.TextMatrix(i, 5) = IIf(IsNull(!PMTOTALSTDQ), 0, !PMTOTALSTDQ)
            flxviewSummary.TextMatrix(i, 6) = IIf(IsNull(!TOTALCOST), "###,##0.00", !TOTALCOST)
            flxviewSummary.TextMatrix(i, 7) = !partsGroup
            flxviewSummary.TextMatrix(i, 8) = Format$(strTtlExcess, "0.0%")
            flxviewSummary.TextMatrix(i, 9) = Format$(strTtlLacking, "0.0%")
            flxviewSummary.TextMatrix(i, 10) = Format$(strTtlEmpty, "0.0%")
            flxviewSummary.TextMatrix(i, 11) = Format$(!UPDATEDBY, "")
            flxviewSummary.TextMatrix(i, 12) = !bCode
            flxviewSummary.TextMatrix(i, 13) = !PMATRIXID
            .MoveNext
            pBar.Value = pBar.Value + 1
        Next i
    End If
    
End With

If mainRs.State <> 0 Then
    Screen.MousePointer = vbDefault
    mainRs.Close
End If

    pBar.Visible = False
    pBar.Value = 0
    
End Sub
'*********** END CODE UPDATED: 04/19/17 **********************

'********************** CODE UPDATED: 08/11/17 **********************************
Private Sub mnuconfirm_Click()
    SSConfirmTab.Visible = True
    txtConfirmDate.Text = Format(Now, "mm/dd/yyyy")
    With SSConfirmTab
        .Top = (Me.Height - .Height) \ 2
        .Left = (Me.Width - .Width) \ 2
        .Visible = True
        .ZOrder
    End With
End Sub

