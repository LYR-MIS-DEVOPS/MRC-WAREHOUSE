VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form AllocBBO 
   Caption         =   "Allocation of Received BBO"
   ClientHeight    =   8895
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14160
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8895
   ScaleWidth      =   14160
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtTransferNo 
      Height          =   285
      Left            =   3000
      TabIndex        =   32
      Top             =   9120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin TabDlg.SSTab tabTransfer 
      Height          =   1695
      Left            =   14160
      TabIndex        =   22
      Top             =   3600
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2990
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "AllocBBO.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdClose"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdOk"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtPickby"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtCheckBy"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.TextBox txtCheckBy 
         Height          =   285
         Left            =   1560
         TabIndex        =   28
         Tag             =   "xyf"
         Top             =   840
         Width           =   4455
      End
      Begin VB.TextBox txtPickby 
         Height          =   285
         Left            =   1560
         TabIndex        =   24
         Tag             =   "xyf"
         Top             =   450
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   9120
         TabIndex        =   23
         Tag             =   "xyf"
         Top             =   1080
         Width           =   4215
      End
      Begin lvButton.lvButtons_H cmdOk 
         Height          =   375
         Left            =   2640
         TabIndex        =   30
         Tag             =   "xyf"
         Top             =   1200
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
         Image           =   "AllocBBO.frx":001C
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdClose 
         Height          =   375
         Left            =   4320
         TabIndex        =   31
         Tag             =   "xyf"
         Top             =   1200
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
         Image           =   "AllocBBO.frx":05B6
         cBack           =   -2147483633
      End
      Begin VB.Label Label8 
         Caption         =   "CHECKED BY:"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Tag             =   "xyf"
         Top             =   870
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "PICKED BY:"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Tag             =   "xyf"
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "APPROVED BY:"
         Height          =   255
         Left            =   7800
         TabIndex        =   25
         Tag             =   "xyf"
         Top             =   1080
         Width           =   1935
      End
   End
   Begin VB.TextBox txtBranchID 
      Height          =   375
      Left            =   360
      TabIndex        =   21
      Top             =   9120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin lvButton.lvButtons_H cmdBBO 
      Height          =   375
      Left            =   12600
      TabIndex        =   17
      Tag             =   "xyf"
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Process BBO"
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
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   240
      TabIndex        =   12
      Tag             =   "xyf"
      Top             =   7680
      Width           =   13815
      Begin VB.TextBox txtAprove 
         Height          =   285
         Left            =   9360
         TabIndex        =   16
         Tag             =   "xyf"
         Top             =   360
         Width           =   4215
      End
      Begin VB.TextBox txtPrepBy 
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Tag             =   "xyf"
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label5 
         Caption         =   "APPROVED BY:"
         Height          =   255
         Left            =   8040
         TabIndex        =   15
         Tag             =   "xyf"
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "PREPARED BY:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Tag             =   "xyf"
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxAllocRec 
      Height          =   5775
      Left            =   240
      TabIndex        =   11
      Tag             =   "xyf"
      Top             =   1920
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   10186
      _Version        =   393216
      Cols            =   12
      RowHeightMin    =   300
      BackColorBkg    =   -2147483633
      FormatString    =   $"AllocBBO.frx":0B50
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Tag             =   "xyf"
      Top             =   600
      Width           =   13815
      Begin VB.TextBox txtAlocDate 
         Height          =   285
         Left            =   10920
         TabIndex        =   10
         Tag             =   "xyf"
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtAlocRef 
         Height          =   285
         Left            =   10920
         Locked          =   -1  'True
         TabIndex        =   7
         Tag             =   "xyf"
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtBranch 
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   5
         Tag             =   "xyf"
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblAddress 
         Height          =   255
         Left            =   1920
         TabIndex        =   20
         Tag             =   "xyf"
         Top             =   600
         Width           =   6255
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "ALLOCATION DATE:"
         Height          =   255
         Left            =   9240
         TabIndex        =   9
         Tag             =   "xyf"
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ALLOCATION NO."
         Height          =   255
         Left            =   9240
         TabIndex        =   8
         Tag             =   "xyf"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "NAME OF BRANCH"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Tag             =   "xyf"
         Top             =   240
         Width           =   1575
      End
   End
   Begin lvButton.lvButtons_H cmdEdit 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "Edit"
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
      Image           =   "AllocBBO.frx":0C25
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   6720
      TabIndex        =   1
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "Exit"
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
      Image           =   "AllocBBO.frx":11BF
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   375
      Left            =   8400
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "AllocBBO.frx":1759
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   10080
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
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
      Image           =   "AllocBBO.frx":1CF3
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdSummary 
      Height          =   375
      Left            =   3360
      TabIndex        =   18
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&View Summary"
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
      Image           =   "AllocBBO.frx":228D
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   375
      Left            =   5040
      TabIndex        =   19
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "Print Transfer"
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
      Image           =   "AllocBBO.frx":2827
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdNew 
      Height          =   375
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "Add New"
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
      Image           =   "AllocBBO.frx":2DC1
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   12000
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14400
      m_MDI_Height    =   9465
   End
   Begin VB.Menu mnu 
      Caption         =   "hide"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuRow 
      Caption         =   "AllocRec"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "AllocBBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ActRow As Integer, AddNew As Boolean, ans As Integer
Dim Qry As String

Private Sub cmdCancel_Click()
    ans = MsgBox("Do you want to Cancel this transaction?", vbQuestion + vbYesNo, "Cancel")
            If ans = vbNo Then Exit Sub
    
    Call BtnSaveCancel(False)
    Call LockData(True)
    flxAllocRec.Enabled = False
    txtAlocDate.Locked = True
    txtPrepBy.Locked = True
    txtAprove.Locked = True
    Call ClearEntries
End Sub

Private Sub cmdClose_Click()
    txtPickby.Text = ""
    txtCheckBy.Text = ""
    tabTransfer.Visible = False
End Sub

Private Sub cmdEdit_Click()
    
    Qry = GetFieldValue("SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUID = 51;")
    If Qry <> 1 Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    
    Qry = GetFieldValue("SELECT stocktransfer.PONO From stocktransfer Where stocktransfer.PONO = 'BBO#:" & Val(txtAlocRef.Text) & "' AND (stocktransfer.`STATUS` = 'INTRANSIT' OR stocktransfer.`STATUS` = 'FINAL');")
    If Qry <> "" Then
        MsgBox "Allocation of BBO#:" & txtAlocRef.Text & " was already INTRANSIT", vbCritical, "EDIT"
        Exit Sub
    End If
    
    If txtbranch.Text = "" Then
        MsgBox "You don't have data to edit", vbCritical, "EDIT"
        Exit Sub
    End If
    ans = MsgBox("Edit BPO Back Order Allocation!", vbQuestion + vbYesNo, "Edit")
    If ans = vbNo Then Exit Sub
    Call BtnSaveCancel(True)
    Frame1.Enabled = False
    Frame2.Enabled = True
    flxAllocRec.Enabled = True
    cmdBBO.Enabled = True
    txtPrepBy.Locked = False
    txtAprove.Locked = False
    AddNew = False
End Sub

Private Sub cmdExit_Click()
    ans = MsgBox("Exit BPO Back Order Allocation!", vbQuestion + vbYesNo, "Exit")
    If ans = vbNo Then Exit Sub
    Unload Me
End Sub

Private Sub cmdBBO_Click()
    If AddNew = False Then
        findBBO.txtAllocBranch.Text = txtbranch.Text
    End If
    findBBO.Show vbModal
End Sub

Private Sub cmdNew_Click()
    Qry = GetFieldValue("SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUID = 51;")
    If Qry <> 1 Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
        
    ans = MsgBox("Add new BPO Back Order Allocation!", vbQuestion + vbYesNo, "Add New")
    If ans = vbNo Then Exit Sub
    Call BtnSaveCancel(True)
    Call ClearEntries
    Call LockData(True)
    txtAlocRef.Text = "<AUTO>"
    txtAlocDate.Text = Format(Date, "MM-DD-YYYY")
    txtAlocDate.Locked = False
    txtPrepBy.Locked = False
    txtAprove.Locked = False
    AddNew = True
    
End Sub

Private Sub cmdOK_Click()

Dim rsBBOTrans      As New ADODB.Recordset
Dim MyRef           As String
Dim wCost           As Variant
Dim AllocRef        As String
Dim MyDate          As String
Dim QryRec          As String
Dim TransID         As Integer

        AllocRef = "BBO#:" & txtAlocRef.Text
        MyDate = Format(Now, "YYYY-MM-DD HH:MM:SS")
        If txtPickby.Text = "" And txtCheckBy.Text = "" Then
            MsgBox "Pick by and Check by is empty", vbCritical, "EMPTY"
            Exit Sub
        ElseIf txtPickby.Text = "" Then
            MsgBox "Pick by is empty", vbCritical, "EMPTY"
            Exit Sub
        ElseIf txtCheckBy.Text = "" Then
            MsgBox "Check by is empty", vbCritical, "EMPTY"
            Exit Sub
        Else
        
            ans = MsgBox("Do you want to save BBO Transfer?", vbQuestion + vbYesNo, "PRINT")
            If ans = vbNo Then Exit Sub
                
                'Save Stocktransfer
                LYRcon.Execute "INSERT INTO stocktransfer(RECBRANCH, TRANSDATE, DATEFINAL, ADDRESS, PONO, PREPBY, APRVBY, REMARKS, CHECKER, `STATUS`, UPDATEDBY, LASTUPDATE) " & _
                            "VALUES('" & txtbranch.Text & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "yyyy-mm-dd") & "', " & _
                            "'" & Replace(lblAddress.Caption, "'", "`") & "', '" & AllocRef & "', '" & txtPrepBy.Text & "', " & _
                            "'" & txtAprove.Text & "', '" & txtPickby.Text & "', '" & txtCheckBy.Text & "', 'INTRANSIT', '" & UserID & "', '" & MyDate & "');"
                 
                QryRec = "SELECT STOCKTRANSFER.TRANSFERID From STOCKTRANSFER " & _
                "WHERE STOCKTRANSFER.RECBRANCH='" & txtbranch.Text & "' AND STOCKTRANSFER.LASTUPDATE='" & Format(MyDate, "yyyy-mm-dd hh:mm:ss") & "';"
                
                TransID = GetFieldValue(QryRec)
                txtTransferNo.Text = IIf(Len(str(TransID)) >= 4, TransID, Format(TransID, "0000"))
                
                'Save Transfer SPITR
                MyRef = "TR#" & txtTransferNo.Text
'                Qry = "SELECT spitr.REFERENCE From spitr Where spitr.REFERENCE = '" & MyRef & "' AND spitr.remarks = '" & AllocRef & "' ;"
'                With rsBBOTrans
'
'                    .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
'                        If .RecordCount = 0 Then
'
                        LYRcon.Execute "DELETE From spitr Where spitr.REMARKS = '" & AllocRef & "' AND spitr.REFERENCE = '" & MyRef & "';"
                           
                            For i = 1 To flxAllocRec.Rows - 1
                                If flxAllocRec.TextMatrix(i, 9) <> 0 Then
                                    
                                    wCost = GetFieldValue("SELECT spareparts.cost From spareparts Where spareparts.SPAREPARTID = " & Val(flxAllocRec.TextMatrix(i, 9)) & ";")
                                                                        
                                    LYRcon.Execute "INSERT INTO spitr( REFERENCE, TRANSTYPE, TRANSDATE, SPAREPARTID, `OUT`, PRICE, COST, REMARKS, UPDATEDBY, LASTUPDATE, SEQNO, POREF ) " & _
                                    "VALUES('" & MyRef & "', 'OUT', '" & Format(Now, "yyyy-mm-dd") & "', " & Val(flxAllocRec.TextMatrix(i, 9)) & ", " & Val(flxAllocRec.TextMatrix(i, 4)) & ", " & Val(flxAllocRec.TextMatrix(i, 11)) & ", " & Val(wCost) & ", '" & AllocRef & "', " & UserID & ", '" & MyDate & "', " & i & ", '" & flxAllocRec.TextMatrix(i, 5) & "');"
                                    
                                    'Update SPAREPARTS QUANTITY ON-HAND
                                    LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = SPAREPARTS.QTYONHAND-" & Val(flxAllocRec.TextMatrix(i, 4)) & ", SPAREPARTS.UPDATEDBY = " & UserID & ", SPAREPARTS.LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
                                    "WHERE SPAREPARTS.SPAREPARTID=" & Val(flxAllocRec.TextMatrix(i, 9)) & ";"
                                    
                                    LYRcon.Execute "UPDATE backorder SET  backorder.BOBALANCE = backorder.BOBALANCE-" & Val(flxAllocRec.TextMatrix(i, 4)) & " " & _
                                                    "Where backorder.REFERENCE = '" & flxAllocRec.TextMatrix(i, 5) & "' AND backorder.SPAREPARTID = " & Val(flxAllocRec.TextMatrix(i, 9)) & ";"
                                
                                End If
                            Next i
                                
'                        End If
'                    .Close
'                End With
                MsgBox "Allocation of BBO Succesfully Transfer!", vbInformation, "PRINT"
                Call PrintTransBBO
        End If
        txtPickby.Text = ""
        txtCheckBy.Text = ""
        tabTransfer.Visible = False
End Sub
Public Sub PrintTransBBO()
        frmPrintPO.MySource = "SP PACKING"
        frmPrintPO.QryPrn = "SELECT spitr.RECEIVE, spitr.OUT, spitr.ORDERBAL, spareparts.PARTNUMBER, spareparts.DESCRIPTION, spitr.SPAREPARTID, spitr.Price " & _
                            "From spareparts INNER JOIN spitr ON (spareparts.SPAREPARTID = spitr.SPAREPARTID) Where spitr.REFERENCE = 'TR#" & txtTransferNo.Text & "';"
            
        frmPrintPO.ParamValue1 = SysBranch
        frmPrintPO.ParamValue2 = SysAddress
        frmPrintPO.ParamValue3 = txtbranch.Text
        frmPrintPO.ParamValue4 = lblAddress.Caption
        frmPrintPO.ParamValue5 = "BBO#:" & txtAlocRef.Text
        frmPrintPO.ParamValue6 = txtTransferNo.Text
        frmPrintPO.ParamValue7 = Format(Now, "mm-dd-yyyy")
        'frmPrintPO.ParamValue8 = IIf(txtStatus.Text <> "INTRANSIT", "INTRANSIT", txtStatus.Text)
        frmPrintPO.ParamValue8 = "INTRANSIT"
        'frmPrintPO.ParamValue8 = GetFieldValue(Qrystatus)
        frmPrintPO.ParamValue9 = "(BBO TRANSFER) " & txtPickby.Text
        frmPrintPO.ParamValue10 = txtPrepBy.Text
        frmPrintPO.ParamValue11 = txtAprove.Text
        frmPrintPO.Show
End Sub
Private Sub cmdPrint_Click()
        
    Qry = GetFieldValue("SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUID = 51;")
    If Qry <> 1 Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
        
    Qry = GetFieldValue("SELECT stocktransfer.PONO From stocktransfer Where stocktransfer.PONO = 'BBO#:" & Val(txtAlocRef.Text) & "' AND (stocktransfer.`STATUS` = 'INTRANSIT' OR stocktransfer.`STATUS` = 'FINAL');")
    If Qry <> "" Then
        MsgBox "Allocation of BBO#:" & txtAlocRef.Text & " was already INTRANSIT", vbCritical, "EDIT"
        Exit Sub
    End If
    
    tabTransfer.Caption = "PRINT TRANSFER"
    If txtbranch.Text = "" Then
        MsgBox "You don't have data to print", vbCritical, "PRINT"
        Exit Sub
    End If
    ans = MsgBox("Print Allocation of BPO Back Order?", vbQuestion + vbYesNo, "PRINT")
    If ans = vbNo Then Exit Sub
   
    flxAllocRec.Enabled = False
    With tabTransfer
        .Visible = True
        .Top = cmdPrint.Top + cmdPrint.Height
        .Left = cmdPrint.Left
    End With
    Call LockData(True)
End Sub

Private Sub cmdSave_Click()
    Dim LstUpt As String
        
        myMsg = ValidateAllocBBO
        If myMsg <> "" Then
                MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
        ans = MsgBox("Do you want to Save this transaction?", vbQuestion + vbYesNo, "Save")
        If ans = vbNo Then Exit Sub
        
            LstUpt = Format(Now, "yyyy-mm-dd hh:mm:ss")
            If AddNew Then
                LYRcon.Execute "INSERT INTO backorderalloc(ALLOCDATE, BRANCHID, PREPAREDBY, APPROVEDBY, LASTUPDATE ) " & _
                                "VALUES('" & Format(txtAlocDate.Text, "yyyy-mm-dd") & "', '" & txtBranchID.Text & "', '" & txtPrepBy.Text & "', '" & txtAprove.Text & "', '" & LstUpt & "');"
                Qry = "SELECT backorderalloc.ALLOCID From backorderalloc Where backorderalloc.BRANCHID = '" & txtBranchID.Text & "' AND backorderalloc.LASTUPDATE = '" & LstUpt & "'"
                txtAlocRef.Text = GetFieldValue(Qry)
            Else
                LYRcon.Execute "UPDATE backorderalloc SET PREPAREDBY = '" & txtPrepBy.Text & "', APPROVEDBY = '" & txtAprove.Text & "', LASTUPDATE = '" & LstUpt & "'"
            End If
            
            LYRcon.Execute "DELETE From boallocdtls Where boallocdtls.ALLOCID = '" & txtAlocRef.Text & "'"
            With flxAllocRec
                For i = 1 To .Rows - 1
                    LYRcon.Execute "INSERT INTO boallocdtls(ALLOCID, SPID, BACKORDERID, ALLOCQTY, AVLQTY, RECREF, RECDATE, SPCOST) " & _
                                    "VALUES(" & txtAlocRef.Text & ", " & Val(.TextMatrix(i, 9)) & ", " & Val(.TextMatrix(i, 10)) & ", " & Val(.TextMatrix(i, 4)) & ", " & Val(.TextMatrix(i, 3)) & ", '" & .TextMatrix(i, 7) & "', '" & .TextMatrix(i, 8) & "', '" & Val(Replace(.TextMatrix(i, 11), ",", "")) & "');"
                Next i
            End With
            MsgBox "Allocation of Receive Branch Back Order is successfully saved!", vbInformation, "Save"
            Call BtnSaveCancel(False)
    
End Sub

Private Sub cmdSummary_Click()
    FrmSummary.mySrc = "AllocBBO"
    FrmSummary.Show
    Unload Me
End Sub

Private Sub flxAllocRec_Click()
    ActRow = flxAllocRec.row
End Sub

Private Sub flxAllocRec_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Call PopupMenu(mnuRow)
    End If
End Sub

Private Sub Form_Load()
    
    Call LockData(False)
    flxAllocRec.Enabled = False
    txtAlocDate.Locked = True
    txtPrepBy.Locked = True
    txtAprove.Locked = True
End Sub

Private Sub BtnSaveCancel(Bln As Boolean)
    
    If Bln Then
        
        cmdNew.Visible = False
        cmdEdit.Visible = False
        cmdSummary.Visible = False
        cmdPrint.Visible = False
        cmdExit.Visible = False
        cmdSave.Visible = True
        cmdCancel.Visible = True
        cmdSave.Left = 0
        cmdCancel.Left = cmdSave.Width
        
    Else
    
        cmdNew.Visible = True
        cmdEdit.Visible = True
        cmdSummary.Visible = True
        cmdPrint.Visible = True
        cmdExit.Visible = True
        cmdSave.Visible = False
        cmdSave.Visible = False
    End If
End Sub

Private Sub mnuDelete_Click()
    Dim ans As Integer
    
    ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If ans = vbYes Then
        If ActRow = 1 And flxAllocRec.Rows = 2 Then
            flxAllocRec.Rows = 1
            flxAllocRec.Rows = 2
        Else
            flxAllocRec.RemoveItem ActRow
        End If
    End If
End Sub

Public Sub ShowRecord(MyID As Long)
    
    Dim rsAlloc As New ADODB.Recordset
    Dim rsAllocDtls As New ADODB.Recordset
    
'    Qry = "SELECT backorderalloc.ALLOCID, backorderalloc.ALLOCDATE, backorderalloc.BRANCHNAME, backorderalloc.PREPAREDBY, backorderalloc.APPROVEDBY " & _
'            "From backorderalloc Where backorderalloc.ALLOCID = " & MyID & ";"
    
    Qry = "SELECT DISTINCT backorderalloc.ALLOCID, backorderalloc.ALLOCDATE, backorderalloc.PREPAREDBY, backorderalloc.APPROVEDBY, backorderalloc.BRANCHID, branch.BRANCHNAME, Branch.ADDRESS " & _
            "From backorderalloc INNER JOIN branch ON (backorderalloc.BRANCHID = branch.BRANCHID)Where backorderalloc.ALLOCID = " & MyID & ";"
    
    With rsAlloc
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                Call ClearEntries
                txtbranch.Text = !branchName
                lblAddress.Caption = !Address
                txtBranchID.Text = !branchID
                txtAlocRef.Text = !AllocID
                txtAlocDate.Text = Format(!ALLOCDATE, "MM/DD/YYYY")
                txtPrepBy.Text = IIf(!PreparedBy = Empty, "", !PreparedBy)
                txtAprove.Text = IIf(!ApprovedBy = Empty, "", !ApprovedBy)
            End If
        .Close
    End With
    
    Qry = "SELECT spareparts.PARTNUMBER, spareparts.DESCRIPTION, boallocdtls.SPCOST, boallocdtls.AVLQTY, boallocdtls.ALLOCQTY, backorder.REFERENCE, backorder.TRANSDATE, boallocdtls.RECREF, " & _
            "boallocdtls.RECDATE, boallocdtls.SPID, boallocdtls.BACKORDERID From boallocdtls INNER JOIN spareparts ON (boallocdtls.SPID = spareparts.SPAREPARTID) " & _
            "INNER JOIN backorder ON (boallocdtls.BACKORDERID = backorder.BACKORDERID) Where boallocdtls.ALLOCID = " & MyID & ";"

    With rsAllocDtls
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                flxAllocRec.Rows = 1
                .MoveFirst
                Do While Not .EOF
                    flxAllocRec.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & Val(!AVLQTY) & vbTab & Val(!ALLOCQTY) & vbTab & !REFERENCE & vbTab & Format(!TRANSDATE, "MM-DD-YYYY") & vbTab & !RECREF & vbTab & !RECDATE & vbTab & Val(!SPid) & vbTab & Val(!BACKORDERID) & vbTab & Format(!SPcost, "#,###,##0.00")
                .MoveNext
                Loop
            End If
        .Close
    End With
    
    
End Sub

Private Sub RXResizer1_ResizeComplete()

    With flxAllocRec
    '   |<PARTNUMBER                 |<DISCRIPTION                                  |^QOH          |^ALLOC.         |<P.O. REF#                    |<P.O. DATE              |<REC. REF#                     |<REC DATE            |
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.2 'PARTNUMBER             1
        .ColWidth(2) = (.Width - 300) * 0.2 'DISCRIPTION            2
        .ColWidth(3) = (.Width - 300) * 0.05 'QOH                   3
        .ColWidth(4) = (.Width - 300) * 0.05 'ALLOC.                4
        .ColWidth(5) = (.Width - 300) * 0.15 'PO REF#               5
        .ColWidth(6) = (.Width - 300) * 0.09 'PO DATE               6
        .ColWidth(7) = (.Width - 300) * 0.15 'REC REF#              7
        .ColWidth(8) = (.Width - 300) * 0.09 'REC DATE              8
        .ColWidth(9) = 0                     'SPID                  9
        .ColWidth(10) = 0                    'boid                  10
        .ColWidth(11) = 0                    'cost                  11
    End With

End Sub

Private Sub LockData(Bln As Boolean)
        
        Frame1.Enabled = Bln
        Frame2.Enabled = Bln
        cmdBBO.Enabled = Bln
        flxAllocRec.Enabled = Bln
        
End Sub

Function ValidateAllocBBO() As String
    Dim vMsg As String, Qry As String, RRno As String
    Dim i As Integer, cnt As Integer
    
    vMsg = ""
    If txtAlocDate.Text = "" Then vMsg = "Incomplete entry on field Date"
    If txtPrepBy.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field PREPARED BY"
    If txtAprove.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field APPROVED BY"
    ValidateSPRec = vMsg
End Function

Private Sub ClearEntries()
    txtbranch.Text = ""
    txtAlocRef.Text = ""
    txtAlocDate.Text = ""
    txtAprove.Text = ""
    txtPrepBy.Text = ""
    lblAddress.Caption = ""
    flxAllocRec.Rows = 1
    flxAllocRec.Rows = 2
End Sub
