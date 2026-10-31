VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmViewIntransit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "VIEW INTRANSIT"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid flxViewIntrn 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   10821
      _Version        =   393216
      Cols            =   5
      BackColorBkg    =   14737632
      FormatString    =   "|<BRANCHNAME |^INTRANSIT|<PENDING |^SOURCE"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   1920
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11145
      m_MDI_Height    =   7350
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "TOTAL INTRANSIT:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   6600
      Width           =   1695
   End
End
Attribute VB_Name = "frmViewIntransit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub LoadIntransit(SPid As Long, DateStart As Date, DateEnd As Date, pTYPE As String)
On Error GoTo er
Dim Qry As String
Dim rsQry As New ADODB.Recordset
Dim i As Long
    
    If pTYPE = "WAREHOUSE" Then
        Qry = "call mrcwarehouse.sp_viewIntransit(" & Val(SPid) & ",'" & Format(DateStart, "YYYY-MM-DD") & "','" & Format(DateEnd, "YYYY-MM-DD") & "')"
    Else
        Qry = "call mrcwarehouse.sp_viewIntransit_ws(" & Val(SPid) & ",'" & Format(DateStart, "YYYY-MM-DD") & "','" & Format(DateEnd, "YYYY-MM-DD") & "')"
    End If
    
    With rsQry
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxViewIntrn.Rows = 1
            flxViewIntrn.Rows = 2
            flxViewIntrn.Rows = 3
            flxViewIntrn.Rows = 4
            
            flxViewIntrn.Rows = .RecordCount + 1
            .MoveFirst
            
            For i = 1 To flxViewIntrn.Rows - 1
                flxViewIntrn.TextMatrix(i, 1) = !branchName
                flxViewIntrn.TextMatrix(i, 2) = !BRANCH_INT
                flxViewIntrn.TextMatrix(i, 3) = !PENDING
                flxViewIntrn.TextMatrix(i, 4) = !BRANCHSOURCE
                .MoveNext
            Next i
        Else
            MsgBox "No Intransit Found", vbInformation + vbOKOnly, "Validation"
            .Close
            Exit Sub
        End If
        .Close
    End With
    
    Call RXResizer1_ResizeComplete
    Call Total(SPid)
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->LoadIntransit()")
End Sub


Private Sub Total(spareId As Long)
   ' Dim i As Integer
    Dim qty As Integer
    Dim Total As Integer
  '  Dim Qry As String
  '  Dim rsTotal As ADODB.Recordset
  '  Dim cmd As ADODB.Command
   ' Qry = "SELECT SUM(ULAS_ST) +SUM(RIZAL_ST) +SUM(TORIL_ST) +SUM(VAN2_ST) + SUM(VAN56_ST)  +SUM(PANABO_ST) +SUM(SOB_ST) + SUM(GENSAN_ST) +SUM(GANTE_ST) + SUM(VAN55_ST) +SUM(WS_ST) + SUM(WARE_ST) + SUM(MNKYO_ST) +SUM(MKLM_ST) +SUM(KPLNG_ST) +  SUM(APKN_ST) +SUM(MATI_ST) + SUM(MARKET_ST) + SUM(CALINAN_ST) + SUM(MRCTORIL_ST) +SUM(MRCAGDAO_ST) + SUM(COMPOSTELA_ST) As totalInst From mrcwarehouse.POBRANCHQTY where sparepartid = " & Val(spareId) & " ;"
  '  Set rsTotal = New ADODB.Recordset
  '  Set cmd = New ADODB.Command
   

   ' With cmd
  '      .CommandText = Qry
   '     .CommandType = adCmdText
     '   .ActiveConnection = LYRcon
   '     Set rsTotal = .Execute
    'End With
    
  '  If Not rsTotal.EOF Then
     ' Display the result in the label named lblTotal
  '   lblTotal.Caption = rsTotal.Fields("totalInst").Value
   ' End If
   ' rsTotal.Close
   ' Set rsTotal = Nothing
    'Set cmd = Nothing
     
      For i = 1 To flxViewIntrn.Rows - 1
        qty = flxViewIntrn.TextMatrix(i, 2)
        Total = Total + qty
     Next
    lblTotal.Caption = Total
End Sub


Private Sub Form_Resize()
Call RXResizer1_ResizeComplete
End Sub

Private Sub RXResizer1_ResizeComplete()
    DoEvents
    With flxViewIntrn
        .Cols = 5
        .FixedCols = 1
        
        '               0       1            2         3        4
        .FormatString = "|<BRANCHNAME |^INTRANSIT|<TO RECEIVE |^SOURCE"
        
        .ColWidth(0) = 10
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.35   'BRANCH
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.11    'INTRANSIT
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.11    'PENDING
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.43     'SOURCE
        
        lblTotal.Top = .Top + .Height + 50
        lblTotal.Left = .Left + .ColWidth(0) + .ColWidth(1)
        Label1.Top = .Top + .Height + 50
        Label1.Left = .Left + .ColWidth(0)
        
    End With
End Sub

