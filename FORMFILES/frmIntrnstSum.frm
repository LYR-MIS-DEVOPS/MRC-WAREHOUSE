VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIntrnstSum 
   Caption         =   "Intransit Summary"
   ClientHeight    =   8115
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   13095
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrUpdateIntransit 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   7560
      Top             =   0
   End
   Begin TabDlg.SSTab tabDiscrepancy 
      Height          =   1215
      Left            =   7920
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2143
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Parameter"
      TabPicture(0)   =   "frmIntrnstSum.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdCancel"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdOK"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtDiscrepancy"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.TextBox txtDiscrepancy 
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
         IMEMode         =   3  'DISABLE
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   2175
      End
      Begin lvButton.lvButtons_H cmdOK 
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
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
         Image           =   "frmIntrnstSum.frx":001C
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   375
         Left            =   1380
         TabIndex        =   16
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
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
         Image           =   "frmIntrnstSum.frx":05B6
         cBack           =   -2147483633
      End
   End
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   8280
      Top             =   0
   End
   Begin VB.Timer tmrLoad 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7920
      Top             =   0
   End
   Begin VB.Frame Frame2 
      Height          =   6375
      Left            =   120
      TabIndex        =   10
      Tag             =   "XY"
      Top             =   1560
      Width           =   12855
      Begin MSFlexGridLib.MSFlexGrid flxIntrnst 
         Height          =   6015
         Left            =   120
         TabIndex        =   11
         Tag             =   "XY"
         Top             =   240
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   10610
         _Version        =   393216
         Cols            =   12
         FormatString    =   "|^DATE FINAL |^REF NO |<PARTNUMBER |<DESCRIPTION |^QTY |^DAYS  |^DATEREC |<REC BRANCH |^REMARKS |^LACKING |^EXCESS"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Tag             =   "XY"
      Top             =   480
      Width           =   12855
      Begin VB.ComboBox cboIntrnstType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmIntrnstSum.frx":0B50
         Left            =   11160
         List            =   "frmIntrnstSum.frx":0B5A
         TabIndex        =   18
         Tag             =   "XY"
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cboReportType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmIntrnstSum.frx":0B71
         Left            =   11160
         List            =   "frmIntrnstSum.frx":0B7B
         TabIndex        =   17
         Tag             =   "XY"
         Top             =   600
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dTO 
         Height          =   375
         Left            =   3600
         TabIndex        =   6
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   84606977
         CurrentDate     =   43614
      End
      Begin VB.ComboBox cboBranch 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   4335
      End
      Begin MSComCtl2.DTPicker dFROM 
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   84606977
         CurrentDate     =   43614
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   375
         Left            =   5400
         TabIndex        =   9
         ToolTipText     =   "Execute"
         Top             =   600
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
         Image           =   "frmIntrnstSum.frx":0B91
         cBack           =   -2147483633
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "INTRANSIT TYPE:"
         Height          =   255
         Left            =   9600
         TabIndex        =   20
         Tag             =   "XY"
         Top             =   285
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "REPORT TYPE:"
         Height          =   255
         Left            =   9600
         TabIndex        =   19
         Tag             =   "XY"
         Top             =   675
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "TO"
         Height          =   255
         Left            =   3300
         TabIndex        =   7
         Top             =   675
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "DATE COVERED:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   680
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "BRANCH:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   280
         Width           =   975
      End
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "E&xit Summary"
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
      Image           =   "frmIntrnstSum.frx":112B
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
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
      Image           =   "frmIntrnstSum.frx":1B3D
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdUpdate 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "&WH Update Intransit"
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
      Image           =   "frmIntrnstSum.frx":254F
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   6720
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   13335
      m_MDI_Height    =   8685
   End
   Begin VB.Menu mnuStatus 
      Caption         =   "Popupmenu"
      Visible         =   0   'False
      Begin VB.Menu mnuGood 
         Caption         =   "Good"
      End
      Begin VB.Menu mnuLacking 
         Caption         =   "Lacking"
      End
      Begin VB.Menu mnuExcess 
         Caption         =   "Excess"
      End
      Begin VB.Menu mnuNotRec 
         Caption         =   "Not Received"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancelled"
      End
   End
End
Attribute VB_Name = "frmIntrnstSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMain As New ADODB.Recordset

Dim strPartial As String, strTotal As String
Dim Qry As String, BCODE As String
Dim myDiscrepancy As String, ans As String
Dim arrBRANCH() As String

Dim i As Long, j As Long

Dim StatSrch As Boolean, PrtBln As Boolean

Public Sub UpdateWHIntransit()
    DoEvents
    Qry = "call sp_mrcwarehouse.sp_consolidate_receiving_intrnst()"
    
    With rsMain
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenStatic, adLockOptimistic
        
        If .RecordCount > 0 Then
            .MoveFirst
            MsgBox "" & !message & "", vbInformation + vbOKOnly, "Validation"
        End If
        .Close
    End With
    Call cmdRefresh_Click
    
End Sub
Public Sub CboDisplay()
On Error GoTo er

    DoEvents
    Qry = "call sp_mrcwarehouse.sp_intrnst_summ_LoadBranch()"
    
    With rsMain
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        
        If .State = 1 Then
            Call modObjTools.addToListObjectWithArrayGivenField(cboBranch, rsMain, arrBRANCH, "BRANCHCODE", "BRANCHNAME")
        End If
    End With
    
    'Get Data from database
    DoEvents
    If rsMain.State <> 0 Then rsMain.Close
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub


Public Sub RefreshRecord()
    If cboReportType.Text = "SUMMARY" And cboIntrnstType.Text = "OUTBOUND" Then '
        Qry = "call sp_mrcwarehouse.sp_intrnst_summ_outbound('" & BCODE & "','" & Format(dFrom.Value, "YYYY-MM-DD") & "','" & Format(dTo.Value, "YYYY-MM-DD") & "')"
        
        With rsMain
            If .State = adStateOpen Then .Close
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            
            If .RecordCount > 0 Then
                .MoveFirst
                flxIntrnst.Rows = 1
                flxIntrnst.Rows = 2
                flxIntrnst.Cols = 10
                flxIntrnst.Rows = .RecordCount + 1
           
                For i = 1 To flxIntrnst.Rows - 1
                    flxIntrnst.TextMatrix(i, 1) = IIf(IsNull(!DATEFINAL), "", !DATEFINAL)
                    flxIntrnst.TextMatrix(i, 2) = IIf(IsNull(!RefNo), "", !RefNo)
                    flxIntrnst.TextMatrix(i, 3) = IIf(IsNull(!TTLSTFPN), "", !TTLSTFPN)
                    flxIntrnst.TextMatrix(i, 4) = IIf(IsNull(!TTLSTFQTY), "", !TTLSTFQTY)
                    flxIntrnst.TextMatrix(i, 5) = IIf(IsNull(!Days), "", !Days)
                    flxIntrnst.TextMatrix(i, 6) = IIf(IsNull(!DATEREC), "", !DATEREC)
                    flxIntrnst.TextMatrix(i, 7) = IIf(IsNull(!ttlRecPN), "", !ttlRecPN)
                    flxIntrnst.TextMatrix(i, 8) = IIf(IsNull(!ttlRecQty), "", !ttlRecQty)
                    flxIntrnst.TextMatrix(i, 9) = IIf(IsNull(!RECBRANCH), "", !RECBRANCH)
                    .MoveNext
                Next i
            Else
                flxIntrnst.Rows = 1
                flxIntrnst.Rows = 2
                Call Unload(frm_OnProcess2)
                Call RXResizer1_ResizeComplete
                .Close
                Exit Sub
            End If
            .Close
        End With
    ElseIf cboReportType.Text = "DETAILS" And cboIntrnstType.Text = "OUTBOUND" Then
        Qry = "call sp_mrcwarehouse.sp_intrnst_summ_outbound_detailed('" & BCODE & "','" & Format(dFrom.Value, "YYYY-MM-DD") & "','" & Format(dTo.Value, "YYYY-MM-DD") & "')"
        
        With rsMain
            If .State = adStateOpen Then .Close
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            
            If .RecordCount > 0 Then
                .MoveFirst
                flxIntrnst.Rows = 1
                flxIntrnst.Rows = 2
                flxIntrnst.Cols = 13
                flxIntrnst.Rows = .RecordCount + 1
           
                For i = 1 To flxIntrnst.Rows - 1
                    flxIntrnst.TextMatrix(i, 1) = IIf(IsNull(!DATEFINAL), "", !DATEFINAL)
                    flxIntrnst.TextMatrix(i, 2) = IIf(IsNull(!RefNo), "", !RefNo)
                    flxIntrnst.TextMatrix(i, 3) = IIf(IsNull(!PartNumber), "", !PartNumber)
                    flxIntrnst.TextMatrix(i, 4) = IIf(IsNull(!Description), "", !Description)
                    flxIntrnst.TextMatrix(i, 5) = IIf(IsNull(!qty), "", !qty)
                    flxIntrnst.TextMatrix(i, 6) = IIf(IsNull(!Days), "", !Days)
                    flxIntrnst.TextMatrix(i, 7) = IIf(IsNull(!DATEREC), "", !DATEREC)
                    flxIntrnst.TextMatrix(i, 8) = IIf(IsNull(!RECQTY), "", !RECQTY)
                    flxIntrnst.TextMatrix(i, 9) = IIf(IsNull(!RECBRANCH), "", !RECBRANCH)
                    flxIntrnst.TextMatrix(i, 10) = IIf(IsNull(!Remarks), "", !Remarks)
                    flxIntrnst.TextMatrix(i, 11) = IIf(IsNull(!Lacking), "", !Lacking)
                    flxIntrnst.TextMatrix(i, 12) = IIf(IsNull(!Excess), "", !Excess)
                    .MoveNext
                Next i
            Else
                flxIntrnst.Rows = 1
                flxIntrnst.Rows = 2
                Call Unload(frm_OnProcess2)
                Call RXResizer1_ResizeComplete
                .Close
                Exit Sub
            End If
            .Close
        End With
    ElseIf cboReportType.Text = "SUMMARY" And cboIntrnstType.Text = "INBOUND" Then
        Qry = "call sp_mrcwarehouse.sp_intrnst_summ_inbound('" & BCODE & "','" & Format(dFrom.Value, "YYYY-MM-DD") & "','" & Format(dTo.Value, "YYYY-MM-DD") & "')"
        
        With rsMain
            If .State = adStateOpen Then .Close
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            
            If .RecordCount > 0 Then
                .MoveFirst
                flxIntrnst.Rows = 1
                flxIntrnst.Rows = 2
                flxIntrnst.Cols = 11
                flxIntrnst.Rows = .RecordCount + 1
           
                For i = 1 To flxIntrnst.Rows - 1
                    flxIntrnst.TextMatrix(i, 1) = IIf(IsNull(!DATEFINAL), "", !DATEFINAL)
                    flxIntrnst.TextMatrix(i, 2) = IIf(IsNull(!RefNo), "", !RefNo)
                    flxIntrnst.TextMatrix(i, 3) = IIf(IsNull(!TRANSFERBY), "", !TRANSFERBY)
                    flxIntrnst.TextMatrix(i, 4) = IIf(IsNull(!TTLSTFPN), "", !TTLSTFPN)
                    flxIntrnst.TextMatrix(i, 5) = IIf(IsNull(!TTLSTFQTY), "", !TTLSTFQTY)
                    flxIntrnst.TextMatrix(i, 6) = IIf(IsNull(!Days), "", !Days)
                    flxIntrnst.TextMatrix(i, 7) = IIf(IsNull(!DATEREC), "", !DATEREC)
                    flxIntrnst.TextMatrix(i, 8) = IIf(IsNull(!ttlRecPN), "", !ttlRecPN)
                    flxIntrnst.TextMatrix(i, 9) = IIf(IsNull(!ttlRecQty), "", !ttlRecQty)
                    flxIntrnst.TextMatrix(i, 10) = IIf(IsNull(!RECBRANCH), "", !RECBRANCH)
                    .MoveNext
                Next i
            Else
                flxIntrnst.Rows = 1
                flxIntrnst.Rows = 2
                Call Unload(frm_OnProcess2)
                Call RXResizer1_ResizeComplete
                .Close
                Exit Sub
            End If
            .Close
        End With
    ElseIf cboReportType.Text = "DETAILS" And cboIntrnstType.Text = "INBOUND" Then
        Qry = "call sp_mrcwarehouse.sp_intrnst_summ_inbound_detailed('" & BCODE & "','" & Format(dFrom.Value, "YYYY-MM-DD") & "','" & Format(dTo.Value, "YYYY-MM-DD") & "')"
        
        With rsMain
            If .State = adStateOpen Then .Close
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            
            If .RecordCount > 0 Then
                .MoveFirst
                flxIntrnst.Rows = 1
                flxIntrnst.Rows = 2
                flxIntrnst.Cols = 11
                flxIntrnst.Rows = .RecordCount + 1
           
                For i = 1 To flxIntrnst.Rows - 1
                    flxIntrnst.TextMatrix(i, 1) = IIf(IsNull(!DATEFINAL), "", !DATEFINAL)
                    flxIntrnst.TextMatrix(i, 2) = IIf(IsNull(!RefNo), "", !RefNo)
                    flxIntrnst.TextMatrix(i, 3) = IIf(IsNull(!TRANSFERBY), "", !TRANSFERBY)
                    flxIntrnst.TextMatrix(i, 4) = IIf(IsNull(!PartNumber), "", !PartNumber)
                    flxIntrnst.TextMatrix(i, 5) = IIf(IsNull(!Description), "", !Description)
                    flxIntrnst.TextMatrix(i, 6) = IIf(IsNull(!STFQTY), "", !STFQTY)
                    flxIntrnst.TextMatrix(i, 7) = IIf(IsNull(!Days), "", !Days)
                    flxIntrnst.TextMatrix(i, 8) = IIf(IsNull(!DATEREC), "", !DATEREC)
                    flxIntrnst.TextMatrix(i, 9) = IIf(IsNull(!RECQTY), "", !RECQTY)
                    flxIntrnst.TextMatrix(i, 10) = IIf(IsNull(!RECBRANCH), "", !RECBRANCH)
                    .MoveNext
                Next i
            Else
                flxIntrnst.Rows = 1
                flxIntrnst.Rows = 2
                Call Unload(frm_OnProcess2)
                Call RXResizer1_ResizeComplete
                .Close
                Exit Sub
            End If
            .Close
        End With
    End If

    Call RXResizer1_ResizeComplete
End Sub


Private Sub cboBranch_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboBranch
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboBranch-Change()")
End Sub

Private Sub cboBranch_Click()
BCODE = arrBRANCH(cboBranch.ListIndex, 0)
Call cboBranch_Change
End Sub

Private Sub cboBranch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboBranch_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboBranch, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Call cboBranch_Click
        dTo.SetFocus
    End If
End Sub

Private Sub cboBranch_LostFocus()
If BCODE = "MWH" Then
    cmdUpdate.Enabled = False
Else
    cmdUpdate.Enabled = False
End If
End Sub

Private Sub cmdCancel_Click()
    tabDiscrepancy.Visible = False
    myDiscrepancy = ""
    txtDiscrepancy.Text = ""
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()

If Val(txtDiscrepancy.Text) > flxIntrnst.TextMatrix(flxIntrnst.row, 5) Then
    MsgBox "Discrepancy Qty is Greater Than Intransit Qty", vbInformation + vbOKOnly, "Validation"
    txtDiscrepancy.Text = ""
    txtDiscrepancy.SetFocus
    Exit Sub
End If

If txtDiscrepancy.Text <> "" Then
    If myDiscrepancy = "Lacking" Then
        Qry = "call sp_mrcwarehouse.sp_intrnst_auditlog('" & BCODE & "','" & flxIntrnst.TextMatrix(flxIntrnst.row, 2) & "','" & flxIntrnst.TextMatrix(flxIntrnst.row, 3) & "','" & "Lacking" & "'," & Val(txtDiscrepancy.Text) & "," & Val(flxIntrnst.TextMatrix(flxIntrnst.row, 12)) & "," & Val(UserID) & ")"
        With rsMain
            If .State = adStateOpen Then .Close
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            
            If .RecordCount > 0 Then
                .MoveFirst
                MsgBox "'" & !message & "'", vbInformation + vbOKOnly, "'" & !level & "'"
            End If
            .Close
        End With
    ElseIf myDiscrepancy = "Excess" Then
        Qry = "call sp_mrcwarehouse.sp_intrnst_auditlog('" & BCODE & "','" & flxIntrnst.TextMatrix(flxIntrnst.row, 2) & "','" & flxIntrnst.TextMatrix(flxIntrnst.row, 3) & "','" & "Excess" & "'," & Val(flxIntrnst.TextMatrix(flxIntrnst.row, 11)) & "," & Val(txtDiscrepancy.Text) & "," & Val(UserID) & ")"
        With rsMain
            If .State = adStateOpen Then .Close
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            
            If .RecordCount > 0 Then
                .MoveFirst
                MsgBox "'" & !message & "'", vbInformation + vbOKOnly, "'" & !level & "'"
            End If
            .Close
        End With
    End If
End If
tabDiscrepancy.Visible = False
myDiscrepancy = ""
txtDiscrepancy.Text = ""
Call cmdRefresh_Click
End Sub

Private Sub cmdPrint_Click()
On Error GoTo er
    Screen.MousePointer = vbHourglass
    
    If cboReportType.Text = "SUMMARY" And cboIntrnstType.Text = "OUTBOUND" Then
        Qry = "call sp_mrcwarehouse.sp_intrnst_summ_outbound('" & BCODE & "','" & Format(dFrom.Value, "YYYY-MM-DD") & "','" & Format(dTo.Value, "YYYY-MM-DD") & "')"
        frmPrn.QryPrn = Qry
        frmPrn.ParamValue1 = SysBranch
        frmPrn.ParamValue2 = Format(SysAddress, "")
        frmPrn.ParamValue3 = "OUTBOUND SUMMARY"
        frmPrn.ParamValue4 = "DATE COVERED: " & Format(dFrom.Value, "MMMM DD, YYYY") & " TO " & Format(dTo.Value, "MMMM DD, YYYY")
        frmPrn.ParamValue5 = EmployeesName
        frmPrn.MySource = "OUTBOUND SUMMARY"
        frmPrn.Show
    ElseIf cboReportType.Text = "DETAILS" And cboIntrnstType.Text = "OUTBOUND" Then
        Qry = "call sp_mrcwarehouse.sp_intrnst_summ_outbound_detailed('" & BCODE & "','" & Format(dFrom.Value, "YYYY-MM-DD") & "','" & Format(dTo.Value, "YYYY-MM-DD") & "')"
        frmPrn.QryPrn = Qry
        frmPrn.ParamValue1 = SysBranch
        frmPrn.ParamValue2 = Format(SysAddress, "")
        frmPrn.ParamValue3 = "DATE COVERED: " & Format(dFrom.Value, "MMMM DD, YYYY") & " TO " & Format(dTo.Value, "MMMM DD, YYYY")
        frmPrn.ParamValue4 = EmployeesName
        frmPrn.MySource = "OUTBOUND DETAILS"
        frmPrn.Show
    ElseIf cboReportType.Text = "SUMMARY" And cboIntrnstType.Text = "INBOUND" Then
        Qry = "call sp_mrcwarehouse.sp_intrnst_summ_inbound('" & BCODE & "','" & Format(dFrom.Value, "YYYY-MM-DD") & "','" & Format(dTo.Value, "YYYY-MM-DD") & "')"
        frmPrn.QryPrn = Qry
        frmPrn.ParamValue1 = SysBranch
        frmPrn.ParamValue2 = Format(SysAddress, "")
        frmPrn.ParamValue3 = "INBOUND SUMMARY"
        frmPrn.ParamValue4 = "DATE COVERED: " & Format(dFrom.Value, "MMMM DD, YYYY") & " TO " & Format(dTo.Value, "MMMM DD, YYYY")
        frmPrn.ParamValue5 = EmployeesName
        frmPrn.MySource = "INBOUND SUMMARY"
        frmPrn.Show
    ElseIf cboReportType.Text = "DETAILS" And cboIntrnstType.Text = "INBOUND" Then
        Qry = "call sp_mrcwarehouse.sp_intrnst_summ_inbound_detailed('" & BCODE & "','" & Format(dFrom.Value, "YYYY-MM-DD") & "','" & Format(dTo.Value, "YYYY-MM-DD") & "')"
        frmPrn.QryPrn = Qry
        frmPrn.ParamValue1 = SysBranch
        frmPrn.ParamValue2 = Format(SysAddress, "")
        frmPrn.ParamValue3 = "DATE COVERED: " & Format(dFrom.Value, "MMMM DD, YYYY") & " TO " & Format(dTo.Value, "MMMM DD, YYYY")
        frmPrn.ParamValue4 = "INBOUND DETAILS"
        frmPrn.ParamValue5 = EmployeesName
        frmPrn.MySource = "INBOUND DETAILS"
        frmPrn.Show
    End If
    Screen.MousePointer = vbDefault
Exit Sub
er:
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdPrint_Click()")
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo er

    If cboReportType.Text <> "" And cboIntrnstType.Text <> "" Then
        Screen.MousePointer = vbHourglass
        tmrRefresh.Enabled = True
    
        DoEvents
        DoEvents
        Call Load(frm_OnProcess2)
        frm_OnProcess2.Show vbModal
    ElseIf cboReportType.Text = "" Then
        MsgBox "Please Select a Report Type to refresh", vbInformation + vbOKOnly, "Validation"
        Exit Sub
    ElseIf cboIntrnstType.Text = "" Then
        MsgBox "Please Select an Intransit Type to refresh", vbInformation + vbOKOnly, "Validation"
        Exit Sub
    End If
Exit Sub
er:
    tmrRefresh.Enabled = False
    Screen.MousePointer = vbDefault
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdRefresh_Click()")
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo er
    Screen.MousePointer = vbHourglass
    
    ans = MsgBox("Do you want to Update Warehouse Intransit?", vbQuestion + vbYesNo, "Validation")
    If ans = vbNo Then Exit Sub
        
    tmrUpdateIntransit.Enabled = True
    DoEvents
    Call Load(frm_OnProcess2)
    frm_OnProcess2.Show vbModal

Exit Sub
er:
    tmrUpdateIntransit.Enabled = False
    Screen.MousePointer = vbDefault
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdUpdate_Click()")
End Sub

Private Sub flxIntrnst_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        If cboReportType.Text = "DETAILS" And cboIntrnstType.Text = "OUTBOUND" And flxIntrnst.Col = 9 Then
            Call PopupMenu(mnuStatus)
        End If
    End If
End Sub

Private Sub Form_Load()
On Error GoTo er

    DoEvents
    tmrRefresh.Enabled = False

    tmrLoad.Enabled = True
    
    DoEvents
    Call Load(frm_OnProcess2)
    frm_OnProcess2.Show vbModal
    
Exit Sub
er:
    tmrLoad.Enabled = False
    Call modObjTools.ErrorMessage(Err.Description, "Error!->Form_Load()")
End Sub

Private Sub mnuCancel_Click()
ans = MsgBox("Do you want to Update Intransit Remarks?", vbQuestion + vbYesNo, "Validation")
If ans = vbNo Then Exit Sub

If flxIntrnst.TextMatrix(flxIntrnst.row, 3) <> "" Then
    Qry = "call sp_mrcwarehouse.sp_intrnst_auditlog('" & BCODE & "','" & flxIntrnst.TextMatrix(flxIntrnst.row, 2) & "','" & flxIntrnst.TextMatrix(flxIntrnst.row, 3) & "','" & "Cancelled" & "'," & Val(flxIntrnst.TextMatrix(flxIntrnst.row, 11)) & "," & Val(flxIntrnst.TextMatrix(flxIntrnst.row, 12)) & "," & Val(UserID) & ")"
    With rsMain
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        
        If .RecordCount > 0 Then
            .MoveFirst
            MsgBox "'" & !message & "'", vbInformation + vbOKOnly, "'" & !level & "'"
        End If
        .Close
    End With
    Call cmdRefresh_Click
End If
End Sub

Private Sub mnuExcess_Click()
ans = MsgBox("Do you want to Update Intransit Remarks?", vbQuestion + vbYesNo, "Validation")
If ans = vbNo Then Exit Sub

If flxIntrnst.TextMatrix(flxIntrnst.row, 3) <> "" Then
    With tabDiscrepancy
        .Top = (Me.Height - .Height) / 2
        .Left = (Me.Width - .Width) / 2
        .Visible = True
        .Caption = "Excess"
    End With
    myDiscrepancy = "Excess"
End If
End Sub

Private Sub mnuGood_Click()
ans = MsgBox("Do you want to Update Intransit Remarks?", vbQuestion + vbYesNo, "Validation")
If ans = vbNo Then Exit Sub

If flxIntrnst.TextMatrix(flxIntrnst.row, 3) <> "" Then
    Qry = "call sp_mrcwarehouse.sp_intrnst_auditlog('" & BCODE & "','" & flxIntrnst.TextMatrix(flxIntrnst.row, 2) & "','" & flxIntrnst.TextMatrix(flxIntrnst.row, 3) & "','" & "Good" & "'," & Val(flxIntrnst.TextMatrix(flxIntrnst.row, 11)) & "," & Val(flxIntrnst.TextMatrix(flxIntrnst.row, 12)) & "," & Val(UserID) & ")"
    With rsMain
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        
        If .RecordCount > 0 Then
            .MoveFirst
            MsgBox "'" & !message & "'", vbInformation + vbOKOnly, "'" & !level & "'"
        End If
        .Close
    End With
    Call cmdRefresh_Click
End If
End Sub

Private Sub mnuLacking_Click()
ans = MsgBox("Do you want to Update Intransit Remarks?", vbQuestion + vbYesNo, "Validation")
If ans = vbNo Then Exit Sub

If flxIntrnst.TextMatrix(flxIntrnst.row, 3) <> "" Then
    With tabDiscrepancy
        .Top = (Me.Height - .Height) / 2
        .Left = (Me.Width - .Width) / 2
        .Visible = True
        .Caption = "Lacking"
    End With
    myDiscrepancy = "Lacking"
End If
End Sub

Private Sub mnuNotRec_Click()
ans = MsgBox("Do you want to Update Intransit Remarks?", vbQuestion + vbYesNo, "Validation")
If ans = vbNo Then Exit Sub
If flxIntrnst.TextMatrix(flxIntrnst.row, 3) <> "" Then
    Qry = "call sp_mrcwarehouse.sp_intrnst_auditlog('" & BCODE & "','" & flxIntrnst.TextMatrix(flxIntrnst.row, 2) & "','" & flxIntrnst.TextMatrix(flxIntrnst.row, 3) & "','" & "Not Received" & "'," & Val(flxIntrnst.TextMatrix(flxIntrnst.row, 11)) & "," & Val(flxIntrnst.TextMatrix(flxIntrnst.row, 12)) & "," & Val(UserID) & ")"
    With rsMain
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        
        If .RecordCount > 0 Then
            .MoveFirst
            MsgBox "'" & !message & "'", vbInformation + vbOKOnly, "'" & !level & "'"
        End If
        .Close
    End With
    Call cmdRefresh_Click
End If
End Sub


Private Sub RXResizer1_ResizeComplete()
    If cboReportType.Text = "SUMMARY" And cboIntrnstType.Text = "OUTBOUND" Then
        With flxIntrnst
            .Cols = 10
            If BCODE = "MWH" Then
                .FormatString = "|^DATE DELIVER |^REF NO |^TOTAL STF PN |^TOTAL STF QTY |^DAYS |^DATE RECEIVE |^TOTAL REC PN |^TOTAL REC QTY |<RECEIVE BRANCH "
            Else
                .FormatString = "|^DATE FINAL |^REF NO |^TOTAL STF PN |^TOTAL STF QTY |^DAYS |^DATE RECEIVE |^TOTAL REC PN |^TOTAL REC QTY |<RECEIVE BRANCH "
            End If
            .ColWidth(0) = 200
            .ColWidth(1) = (.Width - 300) * 0.1         'DATE FINAL
            .ColWidth(2) = (.Width - 300) * 0.1         'REF NO
            .ColWidth(3) = (.Width - 300) * 0.1         'TTLPN
            .ColWidth(4) = (.Width - 300) * 0.1         'TTLQTY
            .ColWidth(5) = (.Width - 300) * 0.07        'DAYS
            .ColWidth(6) = (.Width - 300) * 0.1         'DATEREC
            .ColWidth(7) = (.Width - 300) * 0.1         'TTLRECPN
            .ColWidth(8) = (.Width - 300) * 0.1         'TTLRECQTY
            .ColWidth(9) = (.Width - 300) * 0.2        'RECBRANCH
            
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 6) = "" Then
                    .Col = 1: .row = i: .ColSel = 9: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                    .Col = 2
                End If
            Next i
        End With
    ElseIf cboReportType.Text = "DETAILS" And cboIntrnstType.Text = "OUTBOUND" Then
        With flxIntrnst
            .Cols = 13
            If BCODE = "MWH" Then
                .FormatString = "|^DATE DELIVER |^REF NO |<PARTNUMBER |<DESCRIPTION |^QTY |^DAYS  |^DATEREC |^REC QTY |<REC BRANCH |^REMARKS |^LACKING |^EXCESS"
            Else
                .FormatString = "|^DATE FINAL |^REF NO |<PARTNUMBER |<DESCRIPTION |^QTY |^DAYS  |^DATEREC |^REC QTY |<REC BRANCH |^REMARKS |^LACKING |^EXCESS"
            End If
            .ColWidth(0) = 200
            .ColWidth(1) = (.Width - 300) * 0.07        'DATE FINAL
            .ColWidth(2) = (.Width - 300) * 0.07        'REF NO
            .ColWidth(3) = (.Width - 300) * 0.12        'PARTNUMBER
            .ColWidth(4) = (.Width - 300) * 0.2        'DESCRIPTION
            .ColWidth(5) = (.Width - 300) * 0.05        'QTY
            .ColWidth(6) = (.Width - 300) * 0.05        'DAYS
            .ColWidth(7) = (.Width - 300) * 0.07        'DATE REC
            .ColWidth(8) = (.Width - 300) * 0.07        'REC QTY
            .ColWidth(9) = (.Width - 300) * 0.13        'REC BRANCH
            .ColWidth(10) = (.Width - 300) * 0.08        'REMARKS
            .ColWidth(11) = (.Width - 300) * 0.05        'LACKING
            .ColWidth(12) = (.Width - 300) * 0.05        'EXCESS
            
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 8) = "" Then 'Not yet received
                    .Col = 1: .row = i: .ColSel = 12: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                    .Col = 2
                End If
                If .TextMatrix(i, 10) = "STF Cancelled" Or .TextMatrix(i, 10) = "Cancelled" Then 'Not yet received
                    .Col = 1: .row = i: .ColSel = 12: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
                    .Col = 2
                End If
            Next i
        End With
    ElseIf cboReportType.Text = "SUMMARY" And cboIntrnstType.Text = "INBOUND" Then
        With flxIntrnst
            .Cols = 11
            .FormatString = "|^DATE FINAL |^REF NO |<TRANSFER BY |^TOTAL STF PN |^TOTAL STF QTY |^DAYS |^DATE RECEIVE |^TOTAL REC PN |^TOTAL REC QTY |<RECEIVE BRANCH "
            .ColWidth(0) = 200
            .ColWidth(1) = (.Width - 300) * 0.07         'DATE FINAL
            .ColWidth(2) = (.Width - 300) * 0.1         'REF NO
            .ColWidth(3) = (.Width - 300) * 0.2         'TRANSFER BY
            .ColWidth(4) = (.Width - 300) * 0.07         'TTLPN
            .ColWidth(5) = (.Width - 300) * 0.07         'TTLQTY
            .ColWidth(6) = (.Width - 300) * 0.05        'DAYS
            .ColWidth(7) = (.Width - 300) * 0.07         'DATEREC
            .ColWidth(8) = (.Width - 300) * 0.07         'TTLRECPN
            .ColWidth(9) = (.Width - 300) * 0.07         'TTLRECQTY
            .ColWidth(10) = (.Width - 300) * 0.2        'RECBRANCH
            
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 7) = "" Then 'Not yet received
                    .Col = 1: .row = i: .ColSel = 10: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                    .Col = 2
                End If
            Next i
        End With
    ElseIf cboReportType.Text = "DETAILS" And cboIntrnstType.Text = "INBOUND" Then
        With flxIntrnst
            .Cols = 11
            .FormatString = "|^DATE FINAL |^REF NO |<TRANSFER BY |<PART NUMBER |<DESCRIPTION |^STF QTY |^DAYS |^DATE RECEIVE |^RECEIVE QTY |<RECEIVE BRANCH "
            .ColWidth(0) = 200
            .ColWidth(1) = (.Width - 300) * 0.07         'DATE FINAL
            .ColWidth(2) = (.Width - 300) * 0.1         'REF NO
            .ColWidth(3) = (.Width - 300) * 0.12         'TRANSFER BY
            .ColWidth(4) = (.Width - 300) * 0.14         'PART NUMBER
            .ColWidth(5) = (.Width - 300) * 0.2         'DESCRIPTION
            .ColWidth(6) = (.Width - 300) * 0.05        'STF QTY
            .ColWidth(7) = (.Width - 300) * 0.05         'DAYS
            .ColWidth(8) = (.Width - 300) * 0.07         'DATE RECEIVE
            .ColWidth(9) = (.Width - 300) * 0.06         'RECEIVE QTY
            .ColWidth(10) = (.Width - 300) * 0.12        'RECBRANCH
            
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 8) = "" Then 'Not yet received
                    .Col = 1: .row = i: .ColSel = 10: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                    .Col = 2
                End If
            Next i
        End With
    End If
End Sub

Private Sub tmrLoad_Timer()
On Error GoTo er

    tmrLoad.Enabled = False
    Screen.MousePointer = vbHourglass
    frmIntrnstSum.Enabled = False
    
    DoEvents
    DoEvents
    Call CboDisplay
    dFrom.Value = DateAdd("D", -30, Date) ' DateSerial(Year(Date), Month(Date), 1)
    dTo.Value = Date
    Call RXResizer1_ResizeComplete
    
    DoEvents
    DoEvents
    
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    
    DoEvents
    frmIntrnstSum.Enabled = True
    Screen.MousePointer = vbDefault

Exit Sub
er:
    frmIntrnstSum.Enabled = True
    Screen.MousePointer = vbDefault
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    Call modObjTools.ErrorMessage(Err.Description, "Error->tmrLoad_Timer()")
End Sub

Private Sub tmrRefresh_Timer()
On Error GoTo er
    tmrRefresh.Enabled = False
    Screen.MousePointer = vbHourglass
    Warehouse.Enabled = False
    DoEvents
    DoEvents
    Call RefreshRecord
    DoEvents
    DoEvents
    
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    
    DoEvents
    DoEvents
    Warehouse.Enabled = True
    Screen.MousePointer = vbDefault
    
Exit Sub
er:
    'dFROM.Enabled = True
    'frmMain.Enabled = True
    Screen.MousePointer = vbDefault
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    Call modObjTools.ErrorMessage(Err.Description, "Error->tmrRefresh_Timer()")
End Sub

Private Sub tmrUpdateIntransit_Timer()
On Error GoTo er
    tmrRefresh.Enabled = False
    Screen.MousePointer = vbHourglass
    Warehouse.Enabled = False
    DoEvents
    DoEvents
    Call UpdateWHIntransit
    DoEvents
    DoEvents
    
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    
    DoEvents
    DoEvents
    Warehouse.Enabled = True
    Screen.MousePointer = vbDefault
    
Exit Sub
er:
    'dFROM.Enabled = True
    'frmMain.Enabled = True
    Screen.MousePointer = vbDefault
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    Call modObjTools.ErrorMessage(Err.Description, "Error->tmrUpdateIntransit_Timer()")
End Sub
