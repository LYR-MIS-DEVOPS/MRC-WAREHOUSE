VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServiceRatioSummary 
   Caption         =   "Summary"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8205
   ScaleWidth      =   13680
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab tabProcessing 
      Height          =   975
      Left            =   4440
      TabIndex        =   25
      Tag             =   "xyf"
      Top             =   2880
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   1720
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmServiceRatioSummary.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProcess"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label21"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "PLEASE WAIT...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1080
         TabIndex        =   27
         Tag             =   "xyf"
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label lblProcess 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   600
         TabIndex        =   26
         Tag             =   "xyf"
         Top             =   600
         Width           =   3975
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   240
         Picture         =   "frmServiceRatioSummary.frx":001C
         Stretch         =   -1  'True
         Tag             =   "xyf"
         Top             =   120
         Width           =   720
      End
   End
   Begin MSComctlLib.ProgressBar Pbar 
      Height          =   150
      Left            =   120
      TabIndex        =   22
      Tag             =   "xy"
      Top             =   7920
      Visible         =   0   'False
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxChk 
      Height          =   4695
      Left            =   16200
      TabIndex        =   20
      Top             =   1200
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   8281
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   12240
      Top             =   0
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   7560
      TabIndex        =   12
      Tag             =   "xyf"
      Top             =   480
      Width           =   6015
      Begin VB.TextBox txtFind 
         Height          =   375
         Left            =   3720
         TabIndex        =   14
         Tag             =   "xyf"
         Top             =   240
         Width           =   2175
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   390
         Left            =   2400
         TabIndex        =   13
         Tag             =   "xyf"
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   688
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
         Image           =   "frmServiceRatioSummary.frx":0A14
         cBack           =   -2147483633
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSum 
      Height          =   6735
      Left            =   120
      TabIndex        =   5
      Tag             =   "xyf"
      Top             =   1320
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   11880
      _Version        =   393216
      RowHeightMin    =   300
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   5400
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
      Image           =   "frmServiceRatioSummary.frx":1426
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdShow 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "Sho&w Form"
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
      Image           =   "frmServiceRatioSummary.frx":1E38
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
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
      Image           =   "frmServiceRatioSummary.frx":23D2
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdFilter 
      Height          =   390
      Left            =   3600
      TabIndex        =   3
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   688
      Caption         =   "&Filter Records"
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
      Image           =   "frmServiceRatioSummary.frx":2DE4
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCheck 
      Height          =   390
      Left            =   13080
      TabIndex        =   21
      Tag             =   "xyf"
      Top             =   0
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   688
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
      Image           =   "frmServiceRatioSummary.frx":37F6
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Tag             =   "XYF"
      Top             =   480
      Visible         =   0   'False
      Width           =   7335
      Begin VB.TextBox txtYear 
         Height          =   315
         Left            =   3135
         TabIndex        =   17
         Tag             =   "XYF"
         Top             =   240
         Width           =   720
      End
      Begin VB.ComboBox cboMonth 
         Height          =   315
         Left            =   1440
         TabIndex        =   16
         Tag             =   "XYF"
         Top             =   240
         Width           =   1695
      End
      Begin MSComCtl2.UpDown UpdYear 
         Height          =   315
         Left            =   3840
         TabIndex        =   18
         Tag             =   "XYF"
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Value           =   2015
         OrigLeft        =   11310
         OrigTop         =   480
         OrigRight       =   11565
         OrigBottom      =   795
         Max             =   2020
         Min             =   2005
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "For the month of"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   300
         Width           =   1215
      End
   End
   Begin lvButton.lvButtons_H cmdSaveExcel 
      Height          =   375
      Left            =   7200
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Save Excel"
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
      Image           =   "frmServiceRatioSummary.frx":4208
      cBack           =   -2147483633
   End
   Begin MSComDlg.CommonDialog dlgDEMO 
      Left            =   9240
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Tag             =   "xyf"
      Top             =   480
      Visible         =   0   'False
      Width           =   7335
      Begin MSComCtl2.DTPicker Dfrom 
         Height          =   300
         Left            =   1440
         TabIndex        =   7
         Tag             =   "xyf"
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         Format          =   101908481
         CurrentDate     =   42107
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   390
         Left            =   6120
         TabIndex        =   8
         Tag             =   "xyf"
         Top             =   195
         Visible         =   0   'False
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   688
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
         Image           =   "frmServiceRatioSummary.frx":47A2
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker Dto 
         Height          =   300
         Left            =   3840
         TabIndex        =   9
         Tag             =   "xyf"
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         Format          =   101580801
         CurrentDate     =   42107
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   390
         Left            =   5707
         TabIndex        =   10
         Tag             =   "xyf"
         Top             =   195
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   688
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
         Image           =   "frmServiceRatioSummary.frx":51B4
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdPrev 
         Height          =   390
         Left            =   3360
         TabIndex        =   11
         Tag             =   "xyf"
         Top             =   195
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   688
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
         Image           =   "frmServiceRatioSummary.frx":5BC6
         cBack           =   -2147483633
      End
      Begin VB.Label Label1 
         Caption         =   "Date Covered:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Tag             =   "xyf"
         Top             =   285
         Width           =   1215
      End
   End
   Begin VB.Label lblToolTip 
      Height          =   255
      Left            =   10185
      TabIndex        =   24
      Tag             =   "xyf"
      Top             =   0
      Width           =   2655
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   11520
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   13920
      m_MDI_Height    =   8775
   End
   Begin VB.Menu mnuSummary 
      Caption         =   "Summary"
      Visible         =   0   'False
      Begin VB.Menu mnuTaken 
         Caption         =   "Action Taken"
      End
   End
End
Attribute VB_Name = "frmServiceRatioSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSummary As New ADODB.Recordset
Dim DailyRange As Boolean, MonthlyRange As Boolean, FixRange As Boolean
Dim ActRow As Integer, fRow As Integer, fCol As Integer, MaxCol As Integer, ShowDate As Integer
Public MySource As String


Private Sub cboMonth_Click()
    If cboMonth.ListIndex = -1 Then Exit Sub
    
    'Identify source whose monthly reports is from Day 1 to Last day of the month
    If mySrc = "CLOSED ACCOUNTS FOR AGENT" Or mySrc = "Monthly Collection" Or mySrc = "Monthly Inventory" Then
        Dfrom.Value = DateSerial(Val(txtYear.Text), cboMonth.ListIndex + 1, 1)
        Dto.Value = DateAdd("M", 1, Dfrom.Value)
        Dto.Value = DateAdd("D", -1, Dto.Value)
    Else    'Use for reports whose date covered is from 7th day to 6th of the following month
        Dfrom.Value = DateSerial(Val(txtYear.Text), cboMonth.ListIndex + 1, 7)
        Dto.Value = DateSerial(Year(DateAdd("M", 1, Dfrom.Value)), Month(DateAdd("M", 1, Dfrom.Value)), 6)
    End If
    Call cmdRefresh_Click
End Sub

Private Sub cmdCheck_Click()
    Dim rsChk As New ADODB.Recordset
    Dim Qry As String, i As Integer, j As Integer
    
    If MySource = "BRANCH BACK ORDER" Then
        Qry = "SELECT backorder.REFERENCE, backorder.BRANCHSPOID, sporderbranch.BRANCHNAME, sporderbranch.branchID, backorder.SPAREPARTID " & _
              "From backorder INNER JOIN sporderbranch ON (backorder.REFERENCE = sporderbranch.REFERENCE) AND (backorder.BRANCHSPOID = sporderbranch.BRANCHSPOID) Order By backorder.TRANSDATE"
        'Display Query in Flx
        If rsChk.State = adStateOpen Then rsChk.Close
        rsChk.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If rsChk.RecordCount = 0 Then   'No records to show
            flxChk.Clear
            flxChk.Rows = 2
            flxChk.Cols = MaxCol
            Call RXResizer1_ResizeComplete
            Exit Sub
        End If
        Set flxChk.DataSource = rsChk   'With records
        'Find Branch in flxChk and display in flxsum
        Pbar.Value = 1
        Pbar.Max = flxSum.Rows
        Pbar.Visible = True
        For i = 1 To flxSum.Rows - 1
            For j = 1 To flxChk.Rows - 1
                If flxSum.TextMatrix(i, 2) = flxChk.TextMatrix(j, 1) And flxSum.TextMatrix(i, 10) = flxChk.TextMatrix(j, 2) And flxSum.TextMatrix(i, 9) = flxChk.TextMatrix(j, 5) Then
                    flxSum.TextMatrix(i, 7) = flxChk.TextMatrix(j, 3)
                    LYRcon.Execute "UPDATE backorder SET BRANCHNAME = '" & flxChk.TextMatrix(j, 3) & "', branchID = " & Val(flxChk.TextMatrix(j, 4)) & " " & _
                        "Where backorder.BRANCHSPOID = " & Val(flxSum.TextMatrix(i, 10)) & " AND backorder.SPAREPARTID = " & Val(flxSum.TextMatrix(i, 9)) & " AND backorder.REFERENCE = '" & flxSum.TextMatrix(i, 2) & "'"
                    Exit For
                End If
            Next j
            Pbar.Value = Pbar.Value + 1
            DoEvents
        Next i
        Pbar.Visible = False
    End If
End Sub

Private Sub cmdExit_Click()
    
    'MotoRep.tbMain.Visible = True
    Unload Me
    'MotoRep.imgBG.Visible = True
    'MotoRep.ToolbarButton.Visible = True
        
End Sub

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxSum
        For i = fRow To .Rows - 1
            For j = fCol To MaxCol - 1
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .col = 1: .row = i
                    .ColSel = MaxCol - 1: .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = MaxCol - 1, i + 1, i)
                    fCol = IIf(j = MaxCol - 1, 1, j + 1)
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

Private Sub cmdNext_Click()
    cmdNext.Visible = False
    cmdPrev.Visible = True
    Dto.Visible = True
    ShowDate = 2
End Sub

Private Sub cmdPrev_Click()
    Dto.Value = Dfrom.Value
    cmdNext.Left = cmdPrev.Left
    cmdNext.Visible = True
    cmdPrev.Visible = False
    Dto.Visible = False
    ShowDate = 1
End Sub

Private Sub cmdPrint_Click()
    Dim Qry As String
    
    
    If MySource = "BRANCH BACK ORDER" Then
        Qry = "SELECT DATE_FORMAT(backorder.TRANSDATE,'%c/%d/%Y'), backorder.REFERENCE, spareparts.PARTNUMBER, spareparts.DESCRIPTION, backorder.BACKORDERQTY, backorder.BOBALANCE, backorder.BRANCHNAME, backorder.SUPPORDERREF " & _
              "From backorder INNER JOIN spareparts ON (backorder.SPAREPARTID = spareparts.SPAREPARTID) " & _
              "WHERE backorder.TRANSDATE >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' AND backorder.TRANSDATE <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' ORDER BY backorder.TRANSDATE DESC"
    
        frmReport.MySource = "BRANCH BACK ORDER"
        frmReport.QryPrn = Qry
        frmReport.ParamValue1 = "LYR MOTORPARTS CORPORATION"
        frmReport.ParamValue2 = "GANTE AVENUE, TAGUM CITY"
        frmReport.Show vbModal
    
    ElseIf MySource = "RECEIVING DISCREPANCY" Then
        Qry = "SELECT spareparts.PARTNUMBER, spareparts.DESCRIPTION, sprecdiscrepancy.PACKINGNO, DATE_FORMAT(sprecdiscrepancy.DATEREC, '%c/%d/%Y'), sprecdiscrepancy.QTY, sprecdiscrepancy.GOOD, " & _
                "sprecdiscrepancy.DAMAGE, sprecdiscrepancy.LACKING, sprecdiscrepancy.EXCESS, sprecdiscrepancy.WRONGPN, sprecdiscrepancy.REMARKS, sprecdiscrepancy.SPAREPARTID " & _
                "From sprecdiscrepancy LEFT OUTER JOIN spareparts ON (sprecdiscrepancy.SPAREPARTID = spareparts.SPAREPARTID) " & _
                "WHERE SPRECDISCREPANCY.DATEREC  >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' AND SPRECDISCREPANCY.DATEREC  <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' ORDER BY sprecdiscrepancy.DATEREC"
    
        If flxSum.TextMatrix(flxSum.Rows - 1, 1) = "" Then
        
            MsgBox "No data to print", vbCritical, "No data"
        Else
            frmPrint.MySource = "RECEIVING DISCREPANCY PRINT"
            frmPrint.QryPrn = Qry
            frmPrint.ParamValue1 = SysBranch
            'frmPrint.ParamValue2 = cboBranch.Text
            frmPrint.ParamValue3 = SysAddress
            frmPrint.Show vbModal
        End If
    End If
End Sub

Private Sub cmdRefresh_Click()
    Dim Qry As String
    DoEvents
    'Pbar.Visible = True
    'Pbar.Value = 0
    tabProcessing.Visible = True
    DoEvents
    If MySource = "SERVICE RATIO" Then
        DoEvents
        Qry = "CALL mrcwarehouse.GetStockTransferDetailsReport('" & Format(Dfrom.Value, "YYYY-MM-DD") & "', '" & Format(Dto.Value, "YYYY-MM-DD") & "');"
    
    End If
    'Pbar.Visible = True
    'Display Query in Flx
    If rsSummary.State = adStateOpen Then rsSummary.Close
    rsSummary.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    
    If rsSummary.RecordCount = 0 Then   'No records to show
        flxSum.Clear
        flxSum.Rows = 2
        flxSum.Cols = MaxCol
        'Call RXResizer1_ResizeComplete
        tabProcessing.Visible = False
        Exit Sub
    End If
    DoEvents
    'Pbar.Max = rsSummary.RecordCount
    DoEvents
    Set flxSum.DataSource = rsSummary   'With records
    DoEvents
    ActRow = 1
    'Pbar.Visible = False
    tabProcessing.Visible = False
    Call RXResizer1_ResizeComplete
          
End Sub

Private Sub cmdSaveExcel_Click()
On Error GoTo er
    Dim XLapp As Excel.Application
    Dim wbkMVR As Excel.Workbook
    Dim wksLineA1 As Excel.Worksheet
    Dim sSaveAsFileName
    
   
    If flxSum.TextMatrix(1, 1) = "" Then
        MsgBox "No Data to Extract", vbCritical, "Error"
        Exit Sub
    End If
    
            With dlgDEMO
                .CancelError = False
                .InitDir = mstrLastDir
                .Flags = cdlOFNHideReadOnly
                '.FileName = App.Path & "\" & Format(dFrom.Value, "MM_YY") & lblConnect.Tag & IIf(Index <> 3, "_NP", "") & ".xlsx"
                .filename = ""
                .filter = "Microsoft Excel (*.xls)|*.xls|All Files(*.*)|*.*"
                .ShowOpen
                If dlgDEMO.filename = "" Then Exit Sub
                sSaveAsFileName = .filename
            End With
            
            Set XLapp = CreateObject("Excel.Application")
            XLapp.Visible = True
            XLapp.Workbooks.Add
            Set wbkMVR = XLapp.ActiveWorkbook
            Set wksLineA1 = wbkMVR.Sheets(1)
        
            If MySource = "SERVICE RATIO" Then
                    'PARTNUMBER, DESCRIPTION, REMARKS, INVDATE,
                    'TRANSNO, OUT, PRICE, SUPPLIERNAME, INVTYPE,
                    'CHARGETO, ADDRESS, FRONTLINER, Branch Name, Branch Category
                    wksLineA1.Name = sysBranchCode & " SERVICE RATIO REPORT"
                    wksLineA1.Cells(1, 1).Value = "BRANCH"
                    wksLineA1.Cells(1, 2).Value = "P.O NO."
                    wksLineA1.Cells(1, 3).Value = "STATUS"
                    wksLineA1.Cells(1, 4).Value = "PARTNUMBER"
                    wksLineA1.Cells(1, 5).Value = "DESCRIPTION"
                    wksLineA1.Cells(1, 6).Value = "ORDER QTY."
                    wksLineA1.Cells(1, 7).Value = "STF NO."
                    wksLineA1.Cells(1, 8).Value = "DATE FINAL"
                    wksLineA1.Cells(1, 9).Value = "QTY. SERVE"
                    wksLineA1.Cells(1, 10).Value = "QTY. UNSERVE"
                    wksLineA1.Cells(1, 11).Value = "SERVICE RATIO"
                    wksLineA1.Cells(1, 12).Value = "CWV"
                    wksLineA1.Cells(1, 13).Value = "SRP"
                    wksLineA1.Cells(1, 14).Value = "CLASSIFICATION"
                    wksLineA1.Cells(1, 15).Value = "SUPPLIER NAME"
                    
                                        
                    With flxSum
                        DoEvents
                        tabProcessing.Visible = True
                        cmdSaveExcel.Enabled = False
                        cmdExit.Enabled = False
                        DoEvents
                        'Pbar.Value = 0
                        'Pbar.Max = .Rows - 1
                        For i = 1 To .Rows - 1
                            For j = 1 To 15
'                                If j = 6 Then
'                                    wksLineA1.Cells(i + 1, j).Value = ""
'                                Else
                                    wksLineA1.Cells(i + 1, j).Value = .TextMatrix(i, j)
                                'End If
                            Next j
                            'Pbar.Value = Pbar.Value + 1
                            
                            DoEvents
                        Next i
                    End With
                    XLapp.DisplayAlerts = False
            End If
            
                        
                       
                wbkMVR.SaveAs filename:=sSaveAsFileName, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
                
                  
                wbkMVR.Close SaveChanges:=False
                XLapp.Quit
 
                Set XLapp = Nothing
                Set wbkMVR = Nothing
                Set wksLineA1 = Nothing
                MsgBox "Finish Saving to Excel!", vbInformation + vbOKOnly, MySource
                tabProcessing.Visible = False
                cmdSaveExcel.Enabled = True
                cmdExit.Enabled = True
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdShow_Click()
    If MySource = "PRODUCTS" Then
        Call frmproductEnt.ShowRecord(Val(flxSum.TextMatrix(ActRow, 8)))
        frmproductEnt.Show vbModal
    End If
End Sub

Private Sub Dfrom_Change()
    If ShowDate = 1 Then Dto.Value = Dfrom.Value
End Sub

Private Sub flxSum_Click()
    ActRow = flxSum.row
End Sub

Private Sub flxSum_EnterCell()
    ActRow = flxSum.row
End Sub

Private Sub flxSum_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        If MySource = "RECEIVING DISCREPANCY" Then
            mnuTaken.Caption = "Process Discrepancy"
            Call PopupMenu(mnuSummary)
        End If
    End If
End Sub

Private Sub Form_Load()
    'lblToolTip.ToolTipText = MotoRep.Caption
    Me.Caption = MySource & " SUMMARY"
    'MotoRep.sBar.Panels(6).Text = MySource
    If MySource = "SERVICE RATIO" Then
    
        cmdShow.Visible = False
        cmdPrint.Visible = False
        cmdFilter.Visible = False
        cmdSaveExcel.Visible = True
        cmdSaveExcel.Left = 0
        cmdExit.Left = cmdSaveExcel.Width
        
    End If
    tabProcessing.Caption = MySource
    Timer1.Enabled = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MySource = "MONTHLY MERCHANDISE INVENTORY" Then
    
        cmdShow.Visible = True
        cmdPrint.Visible = True
        cmdFilter.Visible = True
        
        
    End If
    'MotoRep.toolBarButton.Visible = True
    ''MotoRep.imgBG.Visible = True
End Sub

Private Sub mnuTaken_Click()
    Dim actionTaken As String
        If flxSum.TextMatrix(ActRow, 12) = "" Then
            MsgBox "No Data available", vbCritical, "Action Taken"
            Exit Sub
        End If
        
        actionTaken = InputBox("Please enter action taken for this transaction", "Action Taken")
        If actionTaken = "" Then Exit Sub
        LYRcon.Execute "UPDATE sprecdiscrepancy SET REMARKS = '" & actionTaken & "' " & _
                        "Where sprecdiscrepancy.SPAREPARTID = '" & flxSum.TextMatrix(ActRow, 12) & "'  AND sprecdiscrepancy.PACKINGNO = '" & flxSum.TextMatrix(ActRow, 3) & "' "
        Call cmdRefresh_Click
    
End Sub

Private Sub RXResizer1_ResizeComplete()
    DoEvents
    With flxSum
        
        If MySource = "SERVICE RATIO" Then
            .FormatString = " |<BRANCH |<P.O NUMBER |<STATUS |<PARTNUMBER |<DESCRIPTION |<ORDER QTY |<STF NO. |<DATE FINAL |<SERVED |<UNSERVED |<RATIO |<CWV |<SRP |<CLASSIFICATION |<SUPPLIER"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.06    'BRANCH
            .ColWidth(2) = (.Width - 300) * 0.07    'PO NUMBER
            .ColWidth(3) = (.Width - 300) * 0.09    'STATUS
            .ColWidth(4) = (.Width - 300) * 0.12    'PARTNUMBER
            .ColWidth(5) = (.Width - 300) * 0.12    'DESCRIPTION
            .ColWidth(6) = (.Width - 300) * 0.04    'ORDER QTY
            .ColWidth(7) = (.Width - 300) * 0.06    'STF NO.
            .ColWidth(8) = (.Width - 300) * 0.06     'DATE FINAL
            .ColWidth(9) = (.Width - 300) * 0.06    'QTY SERVE
            .ColWidth(10) = (.Width - 300) * 0.06   'QTY UNSERVE
            .ColWidth(11) = (.Width - 300) * 0.06    'SERVICE RATIO
            .ColWidth(12) = (.Width - 300) * 0.06   'CWV
            .ColWidth(13) = (.Width - 300) * 0.06    'SRP
            .ColWidth(14) = (.Width - 300) * 0.09    'CLASSIFICATION
            .ColWidth(15) = (.Width - 300) * 0.13    'SUPPLIER NAME
            
        End If
        
    DoEvents
    End With
    
End Sub

Private Sub Timer1_Timer()
    Dim i As Integer
    DoEvents
    If MySource = "SERVICE RATIO" Then
        FixRange = False
        DailyRange = True
        MonthlyRange = False
        MaxCol = 15
'        cmdShow.Visible = False
'        cmdFilter.Visible = False
'        cmdPrint.Visible = False
'        cmdSaveExcel.Visible = True
'        cmdSaveExcel.Left = cmdShow.Left
'        cmdExit.Left = cmdPrint.Left
    End If
       
    If DailyRange Then
        Frame1.Visible = True
        If MySource = "DETAILED STOCKTRANSFER" Then
            Frame2.Visible = True
        Else
            Frame2.Visible = False
        End If
       
        Frame3.Visible = False
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto = Date
        
    ElseIf MonthlyRange Then
        Frame1.Visible = False
        Frame3.Visible = True
        For i = 1 To 12
            cboMonth.AddItem MonthName(i)
        Next i
        'UpdYear.Value = Year(Format(Date, "YYYY-MM-DD"))
        txtYear.Text = Year(Now())
        Dfrom.Value = DateSerial(Year(Date), Month(Now()), 1)
        Dto.Value = DateAdd("d", -1, DateAdd("m", 1, Dfrom.Value))
        cboMonth.ListIndex = Month(Dfrom.Value) - 1
        
    ElseIf FixRange Then
        Frame1.Visible = False
        Frame3.Visible = False
    
    End If
    
    DoEvents
    If FrDate <> DateValue("1/1/1977") Then
        Dfrom.Value = FrDate
        Dto.Value = ToDate
    Else
        FrDate = Dfrom.Value
        ToDate = Dto.Value
    End If
    DoEvents
    'Call cmdRefresh_Click
    Timer1.Enabled = False
End Sub

Private Sub UpdYear_DownClick()
    If MySource = "MONTHLY MERCHANDISE INVENTORY" Then
        txtYear.Text = Val(txtYear.Text) - 1
    End If
End Sub

Private Sub UpdYear_UpClick()
    If MySource = "MONTHLY MERCHANDISE INVENTORY" Then
        txtYear.Text = Val(txtYear.Text) + 1
    End If
End Sub
