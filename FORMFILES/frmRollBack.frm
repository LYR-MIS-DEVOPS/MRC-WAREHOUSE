VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRollBack 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rollback Transaction"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   16785
   Icon            =   "frmRollBack.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   16785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab ssTabRBhist 
      Height          =   6735
      Left            =   3480
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      ForeColor       =   255
      TabCaption(0)   =   "ROLL BACK HISTORY"
      TabPicture(0)   =   "frmRollBack.frx":57E2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblClose"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame Frame4 
         Height          =   5535
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   10335
         Begin MSFlexGridLib.MSFlexGrid flxRBTransHist 
            Height          =   5175
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   9128
            _Version        =   393216
            Cols            =   5
            FormatString    =   "|<DATE |<REASON |<STATUS |<RESPONSIBLE"
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   120
         TabIndex        =   14
         Tag             =   "xyf"
         Top             =   360
         Width           =   10335
         Begin VB.TextBox txtTransName 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            TabIndex        =   18
            Tag             =   "xyf"
            Top             =   240
            Width           =   4575
         End
         Begin VB.TextBox txtReference 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8040
            TabIndex        =   15
            Tag             =   "xyf"
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label4 
            Caption         =   "TRANSACTION NAME:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Tag             =   "xyf"
            Top             =   285
            Width           =   1935
         End
         Begin VB.Label Label3 
            Caption         =   "REFERENCE:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6705
            TabIndex        =   16
            Tag             =   "xyf"
            Top             =   285
            Width           =   1455
         End
      End
      Begin VB.Label lblClose 
         Alignment       =   2  'Center
         Caption         =   " ( X )"
         Height          =   255
         Left            =   10080
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Tag             =   "xyf"
      Top             =   0
      Width           =   16575
      Begin VB.TextBox txtFind 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   14280
         TabIndex        =   11
         Tag             =   "xyf"
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cboType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmRollBack.frx":57FE
         Left            =   960
         List            =   "frmRollBack.frx":580B
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker Dfrom 
         Height          =   360
         Left            =   4200
         TabIndex        =   4
         Tag             =   "xyf"
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   80936961
         CurrentDate     =   42107
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   390
         Left            =   8880
         TabIndex        =   5
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
         Image           =   "frmRollBack.frx":5829
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker Dto 
         Height          =   360
         Left            =   6600
         TabIndex        =   6
         Tag             =   "xyf"
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
         _Version        =   393216
         Format          =   80871425
         CurrentDate     =   42107
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   390
         Left            =   8460
         TabIndex        =   7
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
         Image           =   "frmRollBack.frx":623B
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdPrev 
         Height          =   390
         Left            =   6120
         TabIndex        =   8
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
         Image           =   "frmRollBack.frx":6C4D
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   390
         Left            =   12960
         TabIndex        =   12
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
         Image           =   "frmRollBack.frx":765F
         cBack           =   -2147483633
      End
      Begin VB.Label Label1 
         Caption         =   "DATE COVERED:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2740
         TabIndex        =   10
         Tag             =   "xyf"
         Top             =   285
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "STATUS:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Tag             =   "xyf"
         Top             =   285
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   16575
      Begin MSFlexGridLib.MSFlexGrid flxRBhist 
         Height          =   7695
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   16335
         _ExtentX        =   28813
         _ExtentY        =   13573
         _Version        =   393216
         Cols            =   9
         HighLight       =   2
         AllowUserResizing=   3
         FormatString    =   "|^DATE |<TRANSACTION NAME |^REFERENCE |^DRAS # |<STATUS |<REASON |<RESPONSIBLE |<APPROVED BY"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Timer tmAutoLoad 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   600
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   16875
      m_MDI_Height    =   9585
   End
   Begin VB.Menu mnuStatusMgr 
      Caption         =   "PopUpMenu"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuApproved 
         Caption         =   "Approved"
      End
      Begin VB.Menu mnuDenied 
         Caption         =   "Disapproved"
      End
   End
   Begin VB.Menu mnuStatusPrn 
      Caption         =   "PopUpMenu2"
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
   End
   Begin VB.Menu mnuStatusHist 
      Caption         =   "PopUpMenu3"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuHist 
         Caption         =   "View History"
      End
   End
End
Attribute VB_Name = "frmRollBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpRs As New ADODB.Recordset
Dim Qry As String, PrnVal(0 To 12) As String, ans As String
Dim i As Long, j As Long
Dim fRow As Long, fCol As Long
Dim ActRow As Long, ActCol As Long
Dim MaxCol As Integer, ShowDate As Integer


Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxRBhist
        For i = fRow To .Rows - 1
            For j = fCol To MaxCol - 1
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .row = i
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
    cmdRefresh.Left = Dto.Left + Dto.Width + 50
    ShowDate = 2
End Sub

Private Sub cmdPrev_Click()
    Dto.Value = Dfrom.Value
    cmdNext.Left = cmdPrev.Left
    cmdRefresh.Left = Dto.Left
    cmdNext.Visible = True
    cmdPrev.Visible = False
    Dto.Visible = False
    ShowDate = 1
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo ConError
    
    'Display Query in Flx
    Qry = "call sp_mrcwarehouse.sp_rollback_list('" & cboType.Text & "'," & Val(ShowDate) & ",'" & Format(Dfrom.Value, "YYYY-MM-DD") & "','" & Format(Dto.Value, "YYYY-MM-DD") & "')"
    If tmpRs.State = adStateOpen Then tmpRs.Close
    tmpRs.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    
    With tmpRs
        If tmpRs.RecordCount = 0 Then   'No records to show
            flxRBhist.Clear
            flxRBhist.Rows = 2
            flxRBhist.Cols = MaxCol
            Call RXResizer1_ResizeComplete
            Exit Sub
        Else
            .MoveFirst
            flxRBhist.Rows = 2
            flxRBhist.Rows = .RecordCount + 1
            
            For i = 1 To flxRBhist.Rows - 1
                flxRBhist.TextMatrix(i, 1) = IIf(IsNull(!RBDATE), "", Format(!RBDATE, "mm/dd/yyyy"))
                flxRBhist.TextMatrix(i, 2) = IIf(IsNull(!TransNAme), "", UCase(!TransNAme))
                flxRBhist.TextMatrix(i, 3) = IIf(IsNull(!TRANSREFERENCE), "", !TRANSREFERENCE)
                flxRBhist.TextMatrix(i, 4) = IIf(IsNull(!DRASNO), "", !DRASNO)
                flxRBhist.TextMatrix(i, 5) = IIf(IsNull(!RBSTATUS), "", !RBSTATUS)
                flxRBhist.TextMatrix(i, 6) = IIf(IsNull(!REASON), "", !REASON)
                flxRBhist.TextMatrix(i, 7) = IIf(IsNull(!RESPONSIBLE), "", !RESPONSIBLE)
                flxRBhist.TextMatrix(i, 8) = IIf(IsNull(!ApprovedBy), "", !ApprovedBy)
                flxRBhist.TextMatrix(i, 9) = !RBID
                .MoveNext
            Next i
        End If
    End With
    ActRow = 1
Exit Sub
    
ConError:
    MsgBox Err.Description & Err.Number
End Sub

Private Sub flxRBhist_Click()
    If ActCol = 0 Or ActRow = 0 Then Exit Sub
End Sub

Private Sub flxRBhist_EnterCell()
    ActRow = flxRBhist.row
    ActCol = flxRBhist.Col
End Sub

Private Sub flxRBhist_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If flxRBhist.TextMatrix(ActRow, 5) = "FOR APPROVAL" Then
            PopupMenu mnuStatusMgr, vbPopupMenuRightButton
            Exit Sub
        ElseIf flxRBhist.TextMatrix(ActRow, 5) = "PENDING" Then
            PopupMenu mnuStatusPrn, vbPopupMenuRightButton
        Else
            PopupMenu mnuStatusHist, vbPopupMenuRightButton
        End If
    End If
End Sub

Private Sub Form_Load()
    tmAutoLoad.Enabled = True
    Warehouse.sBar.Panels(6).Text = Me.Caption
End Sub

Private Sub lblClose_Click()
ssTabRBhist.Visible = False
txtReference.Text = ""
txtTransName.Text = ""
End Sub

Private Sub mnuApproved_Click()
Dim ans As Integer, MyDRAS As String
Dim ChkExists As String
On Error GoTo er

'    MyDRAS = InputBox("Please DRAS Number:", "ROLLBACK")
'
'    chkExists = GetFieldValue("SELECT RBID FROM rollbacktrans WHERE DRAS = '" & MyDRAS & "'")
'    If chkExists <> "" Then
'        MsgBox "This Descripancy Report and Adjustment Slip number is already exists!" & vbCrLf & "Please Enter unusable DRAS Number!", vbInformation + vbOKOnly, "Validation"
'        Exit Sub
'    End If

    ChkExists = GetFieldValue("Select EMPID From Employees Where EMPID = " & Val(UserID) & " and `POSITION` IN ('Admin. Officer', 'CEO', 'I.T. Personnel') and USERTYPE = 'Admin'")
    If ChkExists = "" Or IsNull(ChkExists) Then
        MsgBox "You are not Allowed to Approved a Roll Back Transaction.", vbInformation + vbOKOnly, "Validation"
        Exit Sub
    End If
    
    DoEvents
    ans = MsgBox("Are you sure to Confirm this Rollback?", vbQuestion + vbYesNo, "Question")
    If ans = vbNo Then Exit Sub
    
    DoEvents
    Qry = "Update sporder SET PRINTCOPY = 0, UPDATEDBY = " & Val(UserID) & ", LASTUPDATE = Now(), CONFIRMNO = "" Where REFERENCE = '" & flxRBhist.TextMatrix(ActRow, 3) & "' "
    Call LYRcon.Execute(Qry)
    
    DoEvents
    Qry = "Update spitr SET TRANSTYPE = 'ORDER TEMP', UPDATEDBY = " & Val(UserID) & ", LASTUPDATE = Now() Where REFERENCE = '" & flxRBhist.TextMatrix(ActRow, 3) & "' "
    Call LYRcon.Execute(Qry)
    
    DoEvents
    Qry = "Update rollbacktrans SET RBSTATUS = 'APPROVED', APPROVEDBY = " & Val(UserID) & ", LASTUPDATE = Now() Where RBID = " & Val(flxRBhist.TextMatrix(ActRow, 9)) & " "
    Call LYRcon.Execute(Qry)
    
    DoEvents
    Qry = "Insert into rollbackhist(RBID, RBDATE, TRANSNAME, TRANSREFERENCE, REASON, RBSTATUS, RESPONSIBLE, LASTUPDATE)Values " & _
          "('" & flxRBhist.TextMatrix(ActRow, 9) & "','" & Format(Now(), "YYYY-MM-DD") & "','Purchase Order','" & flxRBhist.TextMatrix(ActRow, 3) & "','" & flxRBhist.TextMatrix(ActRow, 6) & "','APPROVED'," & Val(UserID) & ",Now());"
    Call LYRcon.Execute(Qry)
    
    LYRcon.Execute "Insert Into transhistory(`DATE`, TRANSNAME, REFERENCE, `STATUS`, RESPONSIBLE ) " & _
        "VALUES('" & Format(Now(), "yyyy-mm-dd hh:mm:ss") & "','ROLLBACK','" & flxRBhist.TextMatrix(ActRow, 3) & "','APPROVED'," & Val(UserID) & ");"
    
    
    Call cmdRefresh_Click
Exit Sub
er:
 Call modObjTools.ErrorMessage(Err.Description, "Error->mnuApproved_Click()")
End Sub

Private Sub mnuDenied_Click()
Dim ans As Integer
Dim ChkExists As String
On Error GoTo er

    ChkExists = ""
    
    ChkExists = GetFieldValue("Select EMPID From Employees Where EMPID = " & Val(UserID) & " and `POSITION` IN ('Admin. Officer', 'CEO', 'I.T. Personnel') and USERTYPE = 'Admin'")
    If ChkExists = "" Or IsNull(ChkExists) Then
        MsgBox "You are not Allowed to Disapproved a Roll Back Transaction.", vbInformation + vbOKOnly, "Validation"
        Exit Sub
    End If

    DoEvents
    ans = MsgBox("Are you sure to Disapproved this Rollback?", vbQuestion + vbYesNo, "Question")
    If ans = vbNo Then Exit Sub
    
    DoEvents
    Qry = "Update rollbacktrans SET RBSTATUS = 'DISAPPROVED', APPROVEDBY = " & Val(UserID) & ", LASTUPDATE = Now() Where RBID = " & Val(flxRBhist.TextMatrix(ActRow, 9)) & " "
    Call LYRcon.Execute(Qry)
    
    DoEvents
    Qry = "Insert into rollbackhist(RBID, RBDATE, TRANSNAME, TRANSREFERENCE, REASON, RBSTATUS, RESPONSIBLE, LASTUPDATE)Values " & _
          "('" & flxRBhist.TextMatrix(ActRow, 9) & "','" & Format(Now(), "YYYY-MM-DD") & "','" & flxRBhist.TextMatrix(ActRow, 2) & "','" & flxRBhist.TextMatrix(ActRow, 3) & "','" & flxRBhist.TextMatrix(ActRow, 6) & "','DISAPPROVED'," & Val(UserID) & ",Now());"
    Call LYRcon.Execute(Qry)
    
    LYRcon.Execute "Insert Into transhistory(`DATE`, TRANSNAME, REFERENCE, `STATUS`, RESPONSIBLE ) " & _
        "VALUES('" & Format(Now(), "yyyy-mm-dd hh:mm:ss") & "','ROLLBACK','" & flxRBhist.TextMatrix(ActRow, 3) & "','DISAPPROVED'," & Val(UserID) & ");"
    
    Call cmdRefresh_Click
Exit Sub
er:
 Call modObjTools.ErrorMessage(Err.Description, "Error->mnuDenied_Click()")
End Sub

Private Sub mnuHist_Click()
    'ssTabRBhist.Visible = True
    txtTransName.Text = flxRBhist.TextMatrix(ActRow, 2)
    txtReference.Text = flxRBhist.TextMatrix(ActRow, 3)
    With ssTabRBhist
        .Top = (Me.Height - .Height) \ 2
        .Left = (Me.Width - .Width) \ 2
        .Visible = True
        .ZOrder
    End With
    
    Qry = "call sp_mrcwarehouse.sp_rollback_hist('" & txtReference.Text & "')"
    If tmpRs.State = adStateOpen Then tmpRs.Close
    tmpRs.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    
    With tmpRs
        If tmpRs.RecordCount = 0 Then   'No records to show
            flxRBTransHist.Clear
            flxRBTransHist.Rows = 2
            flxRBTransHist.Cols = 5
            Call RXResizer1_ResizeComplete
            Exit Sub
        Else
            .MoveFirst
            flxRBTransHist.Rows = 2
            flxRBTransHist.Rows = .RecordCount + 1
            
            For i = 1 To flxRBTransHist.Rows - 1
                flxRBTransHist.TextMatrix(i, 1) = IIf(IsNull(!LASTUPDATE), "", Format(!LASTUPDATE, "mm/dd/yyyy hh:mm:ss"))
                flxRBTransHist.TextMatrix(i, 2) = IIf(IsNull(!REASON), "", UCase(!REASON))
                flxRBTransHist.TextMatrix(i, 3) = IIf(IsNull(!RBSTATUS), "", !RBSTATUS)
                flxRBTransHist.TextMatrix(i, 4) = IIf(IsNull(!RESPONSIBLE), "", !RESPONSIBLE)
                .MoveNext
            Next i
        End If
    End With
    
End Sub

Private Sub mnuPrint_Click()
Dim ChkExists As String
    ChkExists = GetFieldValue("Select RBID From rollbacktrans Where RESPONSIBLE = " & Val(UserID) & " ")
    If ChkExists = "" Or IsNull(ChkExists) Then
        MsgBox "Your Account is not Responsible to Print this Transaction", vbInformation + vbOKOnly, "Validation"
        Exit Sub
    End If
    
    DoEvents
    ans = MsgBox("Do you want to Print this Transaction for Approval?", vbQuestion + vbYesNo, "Question")
    If ans = vbNo Then Exit Sub
    PrnVal(0) = "TYPE: DISCR. REPORT AND ADJ. SLIP"
    PrnVal(1) = "DATE: " & Format(Now(), "MM/DD/YYYY")
    PrnVal(2) = "REF#: " & flxRBhist.TextMatrix(ActRow, 4)
    PrnVal(3) = ""
    PrnVal(4) = ""
    PrnVal(5) = ""
    PrnVal(6) = ""
    PrnVal(7) = GetFieldValue("Select CONCAT(FIRSTNAME, ' ', LASTNAME) as EMPLOYEE From employees Where EMPID = " & Val(UserID) & " ")
    PrnVal(8) = GetFieldValue("Select `VALUE` From settings where REMARKS = 'Approved By' and description = 'SPTRANSFER' ")
    PrnVal(9) = ""
    PrnVal(10) = "SUBMITTED BY"
    PrnVal(11) = "APPROVED BY"
    PrnVal(12) = IIf(flxRBhist.TextMatrix(ActRow, 6) <> "", "REASON: " & flxRBhist.TextMatrix(ActRow, 6), "")
    Call PrintMyForm(PrnVal, "DISCR REPORT AND ADJ SLIP", flxRBhist)
    
    DoEvents
    Qry = "Update rollbacktrans SET RBSTATUS = 'FOR APPROVAL', RESPONSIBLE = " & Val(UserID) & ", LASTUPDATE = Now() Where RBID = " & Val(flxRBhist.TextMatrix(ActRow, 9)) & " "
    Call LYRcon.Execute(Qry)
    
    DoEvents
    Qry = "Insert into rollbackhist(RBID, RBDATE, TRANSNAME, TRANSREFERENCE, REASON, RBSTATUS, RESPONSIBLE, LASTUPDATE)Values " & _
          "('" & flxRBhist.TextMatrix(ActRow, 9) & "','" & Format(Now(), "YYYY-MM-DD") & "','Purchase Order','" & flxRBhist.TextMatrix(ActRow, 3) & "','" & flxRBhist.TextMatrix(ActRow, 6) & "','FOR APPROVAL'," & Val(UserID) & ",Now());"
    Call LYRcon.Execute(Qry)
    
    LYRcon.Execute "Insert Into transhistory(`DATE`, TRANSNAME, REFERENCE, `STATUS`, RESPONSIBLE ) " & _
        "VALUES('" & Format(Now(), "yyyy-mm-dd hh:mm:ss") & "','ROLLBACK','" & flxRBhist.TextMatrix(ActRow, 3) & "','FOR APPROVAL'," & Val(UserID) & ");"
    
    Call cmdRefresh_Click
End Sub

Private Sub RXResizer1_ResizeComplete()
    With flxRBhist
        '.Clear
        .Rows = 2
        .Cols = 10
        .WordWrap = True
        .AllowUserResizing = flexResizeBoth
        '               0     1          2              3           4      5             6             7
        .FormatString = " |^DATE |<TRANSACTION NAME |^REFERENCE |^DRAS # |^STATUS |<REASON |<RESPONSIBLE |<APPROVED BY "
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.08
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.18
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.12
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.12
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.08
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.3
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.15
        .ColWidth(8) = (.Width - .ColWidth(0)) * 0.15
        .ColWidth(9) = 0
    End With

    With flxRBTransHist
        '.Clear
        .Rows = 2
        .Cols = 5
        .WordWrap = True
        .AllowUserResizing = flexResizeBoth
        '               0     1        2      3           4
        .FormatString = " |^DATE |<REASON |^STATUS |<RESPONSIBLE"
        
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.21
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.6
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.15
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.3
    End With
End Sub

Private Sub tmAutoLoad_Timer()
On Error GoTo er

    Screen.MousePointer = vbHourglass
    frmRollBack.Enabled = False
    
    Dfrom.Value = DateSerial(Year(Now()), Month(Now()), 1)
    Dto.Value = Now()
    ShowDate = 2
    MaxCol = 9
    cboType.Text = "<ALL>"
    Call cmdRefresh_Click

    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    
    DoEvents
    frmRollBack.Enabled = True
    Screen.MousePointer = vbDefault
    tmAutoLoad.Enabled = False
    
Exit Sub
er:
    frmRollBack.Enabled = True
    Screen.MousePointer = vbDefault
    If LoadOnProcess Then
        Call Unload(frm_OnProcess2)
        Set frm_OnProcess2 = Nothing
    End If
    Call modObjTools.ErrorMessage(Err.Description, "Error->tmAutoLoad_Timer()")
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Call cmdFind_Click
End Sub
