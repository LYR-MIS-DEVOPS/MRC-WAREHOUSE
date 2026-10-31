VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form BranchReplOrder 
   Caption         =   "Branch Replenishment Order"
   ClientHeight    =   9150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   13440
   WindowState     =   2  'Maximized
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   3480
      TabIndex        =   35
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "Print Picking List"
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
      Image           =   "Identifying.frx":0000
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtTOTALAMOUNT 
      Height          =   285
      Left            =   10800
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtBranchSPOID 
      Height          =   285
      Left            =   12120
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Height          =   6255
      Left            =   120
      TabIndex        =   12
      Tag             =   "XYF"
      Top             =   2160
      Width           =   13215
      Begin MSComctlLib.ProgressBar pBar 
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   6000
         Visible         =   0   'False
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxOrd 
         Height          =   5895
         Left            =   120
         TabIndex        =   27
         Tag             =   "XYF"
         Top             =   240
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   10398
         _Version        =   393216
         Rows            =   3
         Cols            =   16
         FixedRows       =   2
         FixedCols       =   0
         FormatString    =   "|<PARTNUMBER|<DESCRIPTION|<MODELBRAND|^QH|^RO|^AVQ|^QOH|^BO|^OQ|^ALQ|^PRICE|^TOTAL|^LOCATION"
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
   Begin MSFlexGridLib.MSFlexGrid flxSRQZero 
      Height          =   1335
      Left            =   19320
      TabIndex        =   18
      Top             =   7440
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   2355
      _Version        =   393216
   End
   Begin VB.TextBox tmpSTATUS 
      Height          =   375
      Left            =   19320
      TabIndex        =   17
      Top             =   7920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox tmpREFNO 
      Height          =   375
      Left            =   19320
      TabIndex        =   16
      Top             =   7440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtpmID 
      Height          =   285
      Left            =   13680
      TabIndex        =   15
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtALQ 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   13680
      TabIndex        =   8
      Tag             =   "XYF"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   13680
      TabIndex        =   14
      Text            =   "HoldBackColor"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   13680
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Process R.O."
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
      Image           =   "Identifying.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Edit R.O."
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
      Image           =   "Identifying.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   2
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
      Image           =   "Identifying.frx":10CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   5160
      TabIndex        =   3
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "Identifying.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   13680
      TabIndex        =   4
      Top             =   1200
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
      Image           =   "Identifying.frx":1C02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   13680
      TabIndex        =   5
      Top             =   1680
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
      Image           =   "Identifying.frx":219C
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   30
      Tag             =   "XYF"
      Top             =   8400
      Width           =   13215
      Begin VB.TextBox txtPreparedby 
         Height          =   285
         Left            =   1560
         TabIndex        =   31
         Tag             =   "XYF"
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label lblUP 
         Caption         =   "PREPARED BY:"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Tag             =   "XYF"
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Tag             =   "XYF"
      Top             =   360
      Width           =   13215
      Begin VB.TextBox txtBranch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         TabIndex        =   28
         Tag             =   "XYF"
         Top             =   360
         Width           =   5415
      End
      Begin VB.TextBox txtSTATUS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11040
         TabIndex        =   26
         Tag             =   "XYF"
         Top             =   1180
         Width           =   1935
      End
      Begin VB.TextBox txtordtype 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11040
         TabIndex        =   24
         Tag             =   "XYF"
         Top             =   750
         Width           =   1935
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         TabIndex        =   22
         Tag             =   "XYF"
         Top             =   800
         Width           =   5415
      End
      Begin VB.CheckBox chkFinal 
         Caption         =   "FINAL"
         Height          =   255
         Left            =   3720
         TabIndex        =   20
         Tag             =   "XYF"
         Top             =   1300
         Width           =   855
      End
      Begin MSMask.MaskEdBox txtorddate 
         Height          =   375
         Left            =   11040
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   300
         Width           =   1935
         _ExtentX        =   3413
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
      Begin VB.TextBox txtRefNo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   1230
         Width           =   2295
      End
      Begin VB.Label lblAddress 
         Caption         =   "ADDRESS"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Tag             =   "XYF"
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblSTATUS 
         Caption         =   "STATUS:"
         Height          =   255
         Left            =   9840
         TabIndex        =   25
         Tag             =   "XYF"
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblOrderType 
         Caption         =   "ORDER TYPE:"
         Height          =   255
         Left            =   9840
         TabIndex        =   23
         Tag             =   "XYF"
         Top             =   795
         Width           =   1215
      End
      Begin VB.Label lblRefNo 
         Caption         =   "REF. NO"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Tag             =   "XYF"
         Top             =   1280
         Width           =   855
      End
      Begin VB.Label lblDateFinal 
         Caption         =   "ORDER DATE:"
         Height          =   255
         Left            =   9840
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblBranch 
         Caption         =   "BRANCH"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   400
         Width           =   855
      End
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   13680
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   13680
      m_MDI_Height    =   9720
   End
   Begin VB.Menu mnuStatus 
      Caption         =   "Popupmenu"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "AddRow"
      End
   End
End
Attribute VB_Name = "BranchReplOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Qry As String, ans As String
Dim ActRow As Integer, ActCol As Integer
Dim i As Long, j As Long
Dim formAction As String
Dim branchID As Integer

Sub ComputeTotal()
Dim nrw As Integer, GrnTl As Double, SPTotal As Double
    With flxOrd
        For nrw = 2 To .Rows - 2
            SPTotal = Val(flxOrd.TextMatrix(nrw, 10)) * Val(Replace(flxOrd.TextMatrix(nrw, 11), ",", ""))
            flxOrd.TextMatrix(nrw, 12) = Format(SPTotal, "###,###,##0.#0")
            GrnTl = GrnTl + Val(Replace(flxOrd.TextMatrix(nrw, 12), ",", ""))
        Next nrw
    End With
    txtTotalAmount.Text = Format(GrnTl, "###,###,##0.#0")
End Sub

Sub CreateTempTable()
On Error GoTo er
        LYRcon.Execute "Drop Temporary Table IF  Exists sp_partsmatrix.TMP_SPORDER_DTLS;"

         DoEvents
       LYRcon.Execute " Create Temporary Table IF Not Exists sp_partsmatrix.TMP_SPORDER_DTLS( " & _
        "    ALQ Int(11)            " & _
        "   ,RO Int(11)             " & _
        "   ,BO Int(11)             " & _
        "   ,AMOUNT DECIMAL(10,2)   " & _
        "   ,SPAREPARTID INT(11)    " & _
        ");"
        
        Qry = "Insert into sp_partsmatrix.TMP_SPORDER_DTLS(ALQ,RO,BO,AMOUNT,SPAREPARTID) VALUES "
        For i = 2 To flxOrd.Rows - 1
            If i > 2 Then Qry = Qry & ","
            Qry = Qry & "(" & Val(flxOrd.TextMatrix(i, 10)) & "," & Val(flxOrd.TextMatrix(i, 5)) & "," & Val(flxOrd.TextMatrix(i, 8)) & ",'" & Val(Replace(Trim(flxOrd.TextMatrix(i, 12)), ",", "")) & "'," & Val(flxOrd.TextMatrix(i, 14)) & ") "
        Next i
        
       
       Call LYRcon.Execute(Qry)
       
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CreateTempTable()")
End Sub

Private Sub SetButtonVisibility(DataEntry As Boolean)
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->VIEW SUMM
'3 ->EXIT
'4 ->SAVE
'5 ->CANCEL


    'hide all button
    For i = 0 To 4
        cmdTrans(i).Visible = False
    Next i

    If DataEntry Then
        cmdTrans(5).Visible = True
        cmdTrans(6).Visible = True
    Else
    For i = 0 To 4
        cmdTrans(i).Visible = True
        cmdTrans(5).Visible = False
        cmdTrans(6).Visible = False
    Next i
    
    
    End If
End Sub


Sub LockData(DataEntry As Boolean)
    DataEntry = Not DataEntry

    txtbranch.Locked = DataEntry
    txtAddress.Locked = DataEntry
    txtRefNo.Locked = DataEntry
    'txtorddate.Locked = DataEntry
    txtordtype.Locked = DataEntry
    txtStatus.Locked = DataEntry
    txtPreparedby.Locked = DataEntry
    
    If txtbranch.Locked = True Then
        txtbranch.BackColor = txtHoldGrayBackColor.BackColor
        txtAddress.BackColor = txtHoldGrayBackColor.BackColor
        txtorddate.BackColor = txtHoldGrayBackColor.BackColor
        txtRefNo.BackColor = txtHoldGrayBackColor.BackColor
        txtordtype.BackColor = txtHoldGrayBackColor.BackColor
        txtStatus.BackColor = txtHoldGrayBackColor.BackColor
        txtPreparedby.BackColor = txtHoldGrayBackColor.BackColor
    End If
End Sub


Function GETSUMofALQ(mySPID As Long) As Long
    Dim rsALQ As New ADODB.Recordset
    Dim sumALQ As Long
    
    Qry = "SELECT SUM(spbranchdtl.ALLOCATEDQTY) As SUMofALQ From spbranchdtl " & _
            "Where spbranchdtl.TRANSTYPE = 'ORDER TEMP' AND spbranchdtl.REMARKS = 'PACKING'  AND spbranchdtl.SPAREPARTID = " & mySPID & ";"
    With rsALQ
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            sumALQ = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            sumALQ = 0
        End If
        .Close
    End With
    GETSUMofALQ = sumALQ
    
End Function

Public Sub ShowRequest(myREFNO As String, myOrderID As Integer)
Dim rsPOdetails As New ADODB.Recordset
Dim rsSPdetails As New ADODB.Recordset
Dim strchkFinal As String

    Qry = "call sp_partsmatrix.sp_brnchReplOrd_ShowReqHeader('" & sysBranchCode & "', '" & myREFNO & "', " & myOrderID & ")"
    With rsPOdetails
        Call LockData(True)
        chkFinal.Enabled = False
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            txtbranch.Text = !branchName
            txtAddress.Text = !BAddress
            txtRefNo.Text = !REFERENCE
            txtStatus.Text = IIf(IsNull(!Status), "", !Status)
            txtBranchSPOID.Text = !BRANCHSPOID
            txtordtype.Text = "REPLENISH"
            txtorddate.Text = Format(!ORDERDATE, "mm/dd/yyyy")
            txtPreparedby.Text = EmployeesName
            branchID = !branchID
        End If
    End With
        
    
    Qry = "call sp_partsmatrix.sp_brnchReplOrd_ShowReqDetails('" & sysBranchCode & "'," & branchID & ", '" & myREFNO & "', " & myOrderID & ")"
    With rsSPdetails
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxOrd.Rows = .RecordCount + 2
            For i = 2 To flxOrd.Rows - 1
                Dim reserveQty As Integer
                reserveQty = GETSUMofALQ(!SPAREPARTID)
                flxOrd.TextMatrix(i, 1) = !PartNumber
                flxOrd.TextMatrix(i, 2) = !Description
                flxOrd.TextMatrix(i, 3) = !MODELBRAND
                flxOrd.TextMatrix(i, 4) = !QH
                flxOrd.TextMatrix(i, 5) = reserveQty
                flxOrd.TextMatrix(i, 6) = (!QH - reserveQty) 'AVQ
                flxOrd.TextMatrix(i, 7) = !QOH
                flxOrd.TextMatrix(i, 9) = !Order
                If Val(flxOrd.TextMatrix(i, 6)) >= Val(!Order) Then
                    flxOrd.TextMatrix(i, 10) = Val(!Order)
                ElseIf Val(flxOrd.TextMatrix(i, 6)) < Val(!Order) Then
                    flxOrd.TextMatrix(i, 10) = Val(flxOrd.TextMatrix(i, 6))
                End If
                
                If Val(!Order) > Val(flxOrd.TextMatrix(i, 10)) Then
                    flxOrd.TextMatrix(i, 8) = (!Order - Val(flxOrd.TextMatrix(i, 10)))
                ElseIf Val(!Order) < Val(flxOrd.TextMatrix(i, 10)) Then
                    flxOrd.TextMatrix(i, 8) = (Val(flxOrd.TextMatrix(i, 10)) - !Order)
                Else
                    flxOrd.TextMatrix(i, 8) = 0
                End If
                flxOrd.TextMatrix(i, 11) = !Cost
                flxOrd.TextMatrix(i, 12) = (!Order * !Cost) 'TOTAL
                flxOrd.TextMatrix(i, 13) = !Location
                flxOrd.TextMatrix(i, 14) = !SPAREPARTID
                flxOrd.TextMatrix(i, 15) = !PMATRIXID
                .MoveNext
            Next i
        End If
    End With
    
    Call ComputeTotal
    Call ResizeCompleteHead
    
    If (txtStatus.Text = "PICKING" Or txtStatus.Text = "FINAL") And chkFinal.Value = 0 Then
        cmdTrans(0).Visible = False
        cmdTrans(1).Visible = True
    ElseIf txtStatus.Text = "VIEWED" And chkFinal.Value = 0 Then
        cmdTrans(0).Visible = True
        cmdTrans(1).Visible = False
    End If
    
    Qry = "Select IDENTIFYINGSTATUS From pmreplreference where REFNO = '" & txtRefNo.Text & "' "
    strchkFinal = GetFieldValue(Qry)
    If strchkFinal = "FINAL" Then
        chkFinal.Value = 1
    Else
        chkFinal.Value = 0
    End If

End Sub

Private Sub ResizeCompleteHead()
Dim AdCol, aCol As Integer, aCel As Integer, reAln As Integer, BrWidth As Long
    
With flxOrd
    'MERGE COL
    .TextMatrix(0, 0) = " "
    .TextMatrix(1, 0) = " "
    'MERGE ROW
    .TextMatrix(0, 1) = "SPAREPARTS"
    .TextMatrix(0, 2) = "SPAREPARTS"
    .TextMatrix(0, 3) = "SPAREPARTS"
    .TextMatrix(0, 4) = "WAREHOUSE"
    .TextMatrix(0, 5) = "WAREHOUSE"
    .TextMatrix(0, 6) = "WAREHOUSE"
    .TextMatrix(0, 7) = txtbranch.Text
    .TextMatrix(0, 8) = txtbranch.Text
    .TextMatrix(0, 9) = txtbranch.Text
    .TextMatrix(0, 10) = txtbranch.Text
    .TextMatrix(0, 11) = txtbranch.Text
    .TextMatrix(0, 12) = txtbranch.Text
    .TextMatrix(0, 13) = txtbranch.Text
    
    'MERGE COL =FALSE
    .TextMatrix(1, 1) = "PARTNUMBER"
    .TextMatrix(1, 2) = "DESCRIPTION"
    .TextMatrix(1, 3) = "MODELBRAND (PM)"
    .TextMatrix(1, 4) = "QH"
    .TextMatrix(1, 5) = "RO"
    .TextMatrix(1, 6) = "AVQ"
    .TextMatrix(1, 7) = "QH"
    .TextMatrix(1, 8) = "BO"
    .TextMatrix(1, 9) = "OQ"
    .TextMatrix(1, 10) = "ALQ"
    .TextMatrix(1, 11) = "PRICE"
    .TextMatrix(1, 12) = "TOTAL"
    .TextMatrix(1, 13) = "LOCATION"
    
    'SET MERGING
    .MergeCells = 1
    .MergeRow(0) = True
    .MergeCol(0) = True
    .MergeCol(1) = True
    For i = 1 To 7
        .MergeCol(i) = False
    Next i
    'SET THE COLUMN COLOR GREEN
    .row = 0
    .Col = 0
    .Col = 1
   ' .CellBackColor = vbGreen
End With
With flxOrd
    '.Clear
    .Cols = 16
    '.FormatString = "|^PARTNUMBER |^DESCRIPTION |^MODELBRAND (PM) |^QH |^RO |^AVQ |^QOH |^BO |^OQ |^ALQ |^PRICE |^TOTAL |<LOCATION |^PMID |<SPID"
    .ColWidth(0) = 200
    .ColWidth(1) = (.Width - .ColWidth(0)) * 0.12  'PARTNUMBER
    .ColWidth(2) = (.Width - .ColWidth(0)) * 0.16 'DESCRIPTION
    .ColWidth(3) = (.Width - .ColWidth(0)) * 0.1 'MODELBRAND (PM)
    .ColWidth(4) = (.Width - .ColWidth(0)) * 0.06 'QH
    .ColWidth(5) = (.Width - .ColWidth(0)) * 0.06 'RO
    .ColWidth(6) = (.Width - .ColWidth(0)) * 0.06 'AVQ
    .ColWidth(7) = (.Width - .ColWidth(0)) * 0.06 'QH
    .ColWidth(8) = (.Width - .ColWidth(0)) * 0.06 'BO
    .ColWidth(9) = (.Width - .ColWidth(0)) * 0.06 'OQ
    .ColWidth(10) = (.Width - .ColWidth(0)) * 0.06 'ALQ
    .ColWidth(11) = (.Width - .ColWidth(0)) * 0.06  'PRICE
    .ColWidth(12) = (.Width - .ColWidth(0)) * 0.06  'TOTAL
    .ColWidth(13) = (.Width - .ColWidth(0)) * 0.06  'LOCATION
    .ColWidth(14) = 0 ' (.Width - .ColWidth(0)) * 0.01 ' SPAREPARTID
    .ColWidth(15) = 0 ' (.Width - .ColWidth(0)) * 0.1 ' PMATRIXID
End With

End Sub



Private Sub cmdTrans_Click(Index As Integer)
  Dim rsTrans As New ADODB.Recordset
  Dim allocNum As Integer, printCount As Integer
    
    Select Case Index
    
        Case 0 'PROCESS RO
            If Val(chkFinal.Value) > 0 Then
                MsgBox "P.O. Already Process!", vbExclamation + vbOKOnly, "Edit"
                Exit Sub
            End If
            If Val(txtBranchSPOID.Text) = 0 Then
                 MsgBox "No Record to be process!", vbExclamation + vbOKOnly, "Edit"
                Exit Sub
            End If
            Call SetButtonVisibility(True)
            formAction = "add"
            chkFinal.Enabled = True
         Case 1 'EDIT R.O
            If Val(chkFinal.Value) > 0 And (txtStatus.Text = "PICKING" Or txtStatus.Text = "INTRANSIT" Or txtStatus.Text = "PACKING" Or txtStatus.Text = "FINAL") Then
                MsgBox "This is not allowed because the Branch Replenish Order is already Final!", vbExclamation + vbOKOnly, "Edit"
                Exit Sub
            End If
            If Val(txtBranchSPOID.Text) = 0 Then
                 MsgBox "No Record Found!", vbExclamation + vbOKOnly, "Edit"
                Exit Sub
            End If

            Call SetButtonVisibility(True)
            formAction = "edit"
            chkFinal.Enabled = True
        Case 2 'SUMMARY
            FrmSummary.mySrc = "Branch Purchase Order"
            FrmSummary.Show
            Unload Me
        
        Case 3 'Print Picking Form
            Qry = "Select REFNO,BRANCHCODE FROM PMREPLREFERENCE WHERE REFNO = '" & txtRefNo.Text & "' "
            rsTrans.Open Qry, LYRcon, adOpenKeyset, adLockReadOnly
            
            If rsTrans.RecordCount > 0 Then
                BranchPicking.mySrc = "Picking List"
                BranchPicking.cmboBranch.Text = txtbranch.Text
                BranchPicking.cmboRefno.Text = txtRefNo.Text
                BranchPicking.branchID = branchID
                BranchPicking.txtBCODE = rsTrans!BranchCode
                Call BranchPicking.ShowRecord(branchID, txtRefNo.Text, txtStatus.Text, "Picking List")
                BranchPicking.Show vbModal
            End If
            
            If rsTrans.RecordCount <= 0 Then
                MsgBox "Not Yet Processed. Please Processed the REPLENISH ORDER (R.O) to proceed in Printing Picking List." & vbCrLf & " Click only the Final if you want to Finalized the Order!", vbInformation + vbOKOnly, "Information"
                rsTrans.Close
                Exit Sub
            End If
        Case 4 'EXIT
            Unload Me
            
        Case 5 'SAVE
            ans = MsgBox("Do you want to Save Branch P.O.", vbQuestion + vbYesNo, "Save")
            If ans = vbNo Then Exit Sub
        
            If formAction = "add" Then
                Call CreateTempTable
                Call ComputeTotal
                Qry = "call sp_partsmatrix.sp_brnchReplOrd_Create(" & branchID & ",'" & txtRefNo.Text & "'," & Val(txtBranchSPOID.Text) & ", " & UserID & ", '" & Val(Replace(Trim(txtTotalAmount.Text), ",", "")) & "', " & Val(chkFinal.Value) & ")"
            ElseIf formAction = "edit" Then
                Call CreateTempTable
                Call ComputeTotal
                Qry = "call sp_partsmatrix.sp_brnchReplOrd_Update(" & branchID & ",'" & txtRefNo.Text & "'," & Val(txtBranchSPOID.Text) & ", " & UserID & ", '" & Val(Replace(Trim(txtTotalAmount.Text), ",", "")) & "', " & Val(chkFinal.Value) & ")"
            End If
            
            With rsTrans
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    If !level = "Granted" Then
                        MsgBox " '" & !message & "' ", vbInformation + vbOKOnly, "SUCCESSFUL"
                    Else
                        MsgBox " '" & !message & "' ", vbCritical + vbOKOnly, "UNSUCCESSFUL"
                        Exit Sub
                    End If
                End If
                .Close
            End With
            Call SetButtonVisibility(False)
            formAction = ""
            chkFinal.Enabled = False
            Qry = "Select STATUS From sporderbranch where REFERENCE = '" & txtRefNo.Text & "' "
            txtStatus.Text = GetFieldValue(Qry)
        Case 6 'CANCEL
            ans = MsgBox("Do you want to Cancel the Process?", vbQuestion + vbYesNo, "Cancel")
            If ans = vbNo Then Exit Sub
            Call SetButtonVisibility(False)
            formAction = ""
            
    
    End Select
End Sub

Private Sub flxOrd_Click()
On Error GoTo er

    ActRow = flxOrd.row
    ActCol = flxOrd.Col
    
    If formAction = "add" Or formAction = "edit" Then
    
    If ActCol = 0 Or ActRow = 0 Then Exit Sub
        Select Case ActCol
        Case 10
            With txtALQ
                If flxOrd.TextMatrix(ActRow, 10) <> "" Then .Tag = flxOrd.TextMatrix(ActRow, 10)
                .Top = Frame3.Top + flxOrd.CellTop + flxOrd.Top
                .Left = Frame3.Left + flxOrd.CellLeft + flxOrd.Left
                .Height = flxOrd.CellHeight
                .Width = flxOrd.CellWidth
                .Text = flxOrd.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        End Select
    End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->flxOrd-click()")

End Sub

Private Sub flxOrd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call flxOrd_Click
End Sub

Private Sub Form_Load()
    Call LockData(False)
    
    '===================BUTTON POSITION===================
    If cmdTrans(0).Visible = False Then
        'Align Save Button
        cmdTrans(5).Left = cmdTrans(0).Left
        cmdTrans(5).Top = cmdTrans(0).Top
        cmdTrans(5).Width = cmdTrans(0).Width
        'Align Cancel Button
        cmdTrans(6).Left = cmdTrans(2).Left
        cmdTrans(6).Top = cmdTrans(2).Top
        cmdTrans(6).Width = cmdTrans(2).Width
    End If
End Sub

Private Sub RXResizer1_ResizeComplete()
Call ResizeCompleteHead
End Sub

Private Sub txtALQ_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyEscape Then
'    Tmp = Split(txtALQ.Tag, "|")
'
'    For i = LBound(Tmp) To UBound(Tmp)
'        flxOrd.TextMatrix(ActRow, i + 1) = Tmp(i)
'    Next i
'
'    txtALQ.Visible = False
'    flxOrd.SetFocus
'    flxOrd.Col = 10
'End If
End Sub

Private Sub txtALQ_KeyPress(KeyAscii As Integer)
On Error GoTo er

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtALQ, KeyAscii, NumbersOnly)

    
    If KeyAscii = 13 Then
        
        flxOrd.TextMatrix(ActRow, 10) = txtALQ.Text
        
        If Val(flxOrd.TextMatrix(ActRow, 10)) > Val(flxOrd.TextMatrix(ActRow, 10)) Then
            MsgBox "Allocated Quantity is Greater Than to Available Quantity", vbInformation + vbOKOnly, "Validation"
            flxOrd.TextMatrix(ActRow, 10) = 0
            txtALQ.Text = ""
            txtALQ.Visible = False
            flxOrd.SetFocus
            flxOrd.row = ActRow
            flxOrd.Col = 10
            Exit Sub
        End If
        
        
        If txtALQ.Text <> "" Then
                
            For i = ActRow To flxOrd.Rows - 1
                If flxOrd.TextMatrix(i, 10) = "" Or Val(flxOrd.TextMatrix(i, 10)) = 0 Then
                    txtALQ.Visible = False
                    flxOrd.SetFocus
                    flxOrd.row = i
                    flxOrd.Col = 10
                Exit For
                Else
                    txtALQ.Visible = False
                    flxOrd.SetFocus
                    flxOrd.row = ActRow
                    flxOrd.Col = 10
                End If
                
            Next i
        
        Else
            flxOrd.SetFocus
            flxOrd.Col = 10
            txtALQ.Visible = False
        End If
        'Call ComputeTotalAllocation
        
    End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtALQ-KeyPress()")
End Sub
