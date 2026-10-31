VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPriceHist 
   Caption         =   "Price History"
   ClientHeight    =   8010
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11115
   Icon            =   "frmPriceHist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   11115
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Tag             =   "XYF"
      Top             =   240
      Width           =   10815
      Begin MSFlexGridLib.MSFlexGrid flxSRPPriceHis 
         Height          =   1935
         Left            =   120
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   480
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   3413
         _Version        =   393216
         Rows            =   8
         Cols            =   8
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxWSPriceHis 
         Height          =   1815
         Left            =   120
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   2760
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   3201
         _Version        =   393216
         Rows            =   8
         Cols            =   8
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxKAPriceHis 
         Height          =   1815
         Left            =   120
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   4920
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   3201
         _Version        =   393216
         Rows            =   8
         Cols            =   8
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         AllowUserResizing=   1
      End
      Begin VB.Label Label1 
         Caption         =   "KEY ACCOUNTS PRICE HISTORY"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   4680
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "SRP PRICE HISTORY"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Tag             =   "XYF"
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "WHOLESALE PRICE HISTORY"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Tag             =   "XYF"
         Top             =   2520
         Width           =   3015
      End
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   4560
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11355
      m_MDI_Height    =   8580
   End
End
Attribute VB_Name = "frmPriceHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub LoadSRPHist(pSPID As Long)
    Dim tmpRs As New ADODB.Recordset
    Dim SPid As Integer
    Dim i As Integer
    Dim Qry As String
    
    Qry = "call sp_srppricing.sp_load_srp_history(" & pSPID & ")"
    
'PAG FETCH SA DATA
With tmpRs
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        
    If .RecordCount > 0 Then


        flxSRPPriceHis.Rows = .RecordCount + 1
                .MoveFirst
            For i = 1 To flxSRPPriceHis.Rows - 1
                flxSRPPriceHis.TextMatrix(i, 1) = tmpRs!SRPDATE
                flxSRPPriceHis.TextMatrix(i, 2) = tmpRs!REFERENCE
                flxSRPPriceHis.TextMatrix(i, 3) = Format(tmpRs!Cost, "#,##0.00")
                flxSRPPriceHis.TextMatrix(i, 4) = Format(tmpRs!SRP, "#,##0.00")
                flxSRPPriceHis.TextMatrix(i, 5) = tmpRs!PRORATIO
                flxSRPPriceHis.TextMatrix(i, 6) = (tmpRs!ProMar & "%")
                flxSRPPriceHis.TextMatrix(i, 7) = (tmpRs!discount & "%")
                .MoveNext
                Next i
        
        End If
        .Close
End With
Call LoadWSHist(pSPID)
Call LoadKeyAccHist(pSPID)
Exit Sub

End Sub
Private Sub LoadWSHist(pSPID As Long)
Dim tmpRs As New ADODB.Recordset
Qry = "call sp_srppricing.sp_load_ws_history(" & pSPID & ")"

'PAG FETCH SA DATA
With tmpRs
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        
    If .RecordCount > 0 Then

        flxWSPriceHis.Rows = .RecordCount + 1
                .MoveFirst
            For i = 1 To flxWSPriceHis.Rows - 1
                flxWSPriceHis.TextMatrix(i, 1) = tmpRs!WSDATE
                flxWSPriceHis.TextMatrix(i, 2) = tmpRs!RefNo
                flxWSPriceHis.TextMatrix(i, 3) = Format(tmpRs!Cost, "#,##0.00")
                flxWSPriceHis.TextMatrix(i, 4) = Format(tmpRs!DP, "#,##0.00")
                flxWSPriceHis.TextMatrix(i, 5) = tmpRs!PRORATIO
                flxWSPriceHis.TextMatrix(i, 6) = (tmpRs!ProMar & "%")
                flxWSPriceHis.TextMatrix(i, 7) = (tmpRs!discount & "%")
                .MoveNext
                Next i
       
        End If
        .Close
End With
        Exit Sub

End Sub
Private Sub LoadKeyAccHist(pSPID As Long)
    Dim tmpRs As New ADODB.Recordset
Qry = "call sp_srppricing.sp_load_ka_history(" & pSPID & ")"

'PAG FETCH SA DATA
With tmpRs
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        
    If .RecordCount > 0 Then

        flxKAPriceHis.Rows = .RecordCount + 1
                .MoveFirst
            For i = 1 To flxKAPriceHis.Rows - 1
                flxKAPriceHis.TextMatrix(i, 1) = tmpRs!KADATE
                flxKAPriceHis.TextMatrix(i, 2) = tmpRs!RefNo
                flxKAPriceHis.TextMatrix(i, 3) = Format(tmpRs!Cost, "#,##0.00")
                flxKAPriceHis.TextMatrix(i, 4) = Format(tmpRs!DP, "#,##0.00")
                flxKAPriceHis.TextMatrix(i, 5) = tmpRs!PRORATIO
                flxKAPriceHis.TextMatrix(i, 6) = (tmpRs!ProMar & "%")
                flxKAPriceHis.TextMatrix(i, 7) = (tmpRs!discount & "%")
                .MoveNext
                Next i
            
        End If
        .Close
End With
        Exit Sub
End Sub

Private Sub RXResizer1_ResizeComplete()
On Error Resume Next

    DoEvents
    With flxSRPPriceHis
             .Cols = 8
             '               0   1       2         3       4        5         6
             .FormatString = "|^DATE |^REF. NO. |^COST |^PRICE |^MARKUP |^P-MARGIN |^M-DISCOUNT"
             
             .ColWidth(0) = 200
             .ColWidth(1) = (.Width - .ColWidth(0)) * 0.13 'DATE
             .ColWidth(2) = (.Width - .ColWidth(0)) * 0.2   'REF. NO.
             .ColWidth(3) = (.Width - .ColWidth(0)) * 0.1 'PREV. COST
             .ColWidth(4) = (.Width - .ColWidth(0)) * 0.1  'PREV. SRP
             .ColWidth(5) = (.Width - .ColWidth(0)) * 0.17  'RATIO
             .ColWidth(6) = (.Width - .ColWidth(0)) * 0.15 'MARGIN
             .ColWidth(7) = (.Width - .ColWidth(0)) * 0.15 'DISCOUNT
        End With
        
        With flxWSPriceHis
             .Cols = 8
             '               0    1      2         3     4         5          6
             .FormatString = "|^DATE |^REF. NO. |^COST |^PRICE |^MARKUP |^P-MARGIN |^M-DISCOUNT"
             
             .ColWidth(0) = 200
             .ColWidth(1) = (.Width - .ColWidth(0)) * 0.13  'DATE
             .ColWidth(2) = (.Width - .ColWidth(0)) * 0.2  'REF. NO.
             .ColWidth(3) = (.Width - .ColWidth(0)) * 0.1 'PREV. COST
             .ColWidth(4) = (.Width - .ColWidth(0)) * 0.1 'PREV. SRP
             .ColWidth(5) = (.Width - .ColWidth(0)) * 0.17 'RATIO
             .ColWidth(6) = (.Width - .ColWidth(0)) * 0.15 'MARGIN
             .ColWidth(7) = (.Width - .ColWidth(0)) * 0.15 'DISCOUNT
        End With
        
        With flxKAPriceHis
             .Cols = 8
             '               0    1      2         3     4         5          6
             .FormatString = "|^DATE |^REF. NO. |^COST |^PRICE |^MARKUP |^P-MARGIN |^M-DISCOUNT"
             
             .ColWidth(0) = 200
             .ColWidth(1) = (.Width - .ColWidth(0)) * 0.13  'DATE
             .ColWidth(2) = (.Width - .ColWidth(0)) * 0.2  'REF. NO.
             .ColWidth(3) = (.Width - .ColWidth(0)) * 0.1 'PREV. COST
             .ColWidth(4) = (.Width - .ColWidth(0)) * 0.1 'PREV. SRP
             .ColWidth(5) = (.Width - .ColWidth(0)) * 0.17 'RATIO
             .ColWidth(6) = (.Width - .ColWidth(0)) * 0.15 'MARGIN
             .ColWidth(7) = (.Width - .ColWidth(0)) * 0.15 'DISCOUNT
        End With
End Sub
Sub flxHistColsChanger()
On Error Resume Next
    With flxSRPPriceHis
         .Clear
         .Rows = 2
         .Cols = 7
         '               0   1       2         3       4        5         6
         .FormatString = "|^DATE |^REF. NO. |^COST |^PRICE |^P-RATIO |^P-MARGIN"
         
         .ColWidth(0) = 200
         .ColWidth(1) = (.Width - .ColWidth(0)) * 0.13 'DATE
         .ColWidth(2) = (.Width - .ColWidth(0)) * 0.2  'REF. NO.
         .ColWidth(3) = (.Width - .ColWidth(0)) * 0.1 'PREV. COST
         .ColWidth(4) = (.Width - .ColWidth(0)) * 0.1 'PREV. SRP
         .ColWidth(5) = (.Width - .ColWidth(0)) * 0.17 'RATIO
         .ColWidth(6) = (.Width - .ColWidth(0)) * 0.15 'MARGIN
         .ColWidth(7) = (.Width - .ColWidth(0)) * 0.15 'MARGIN
    
    End With
    
    
    With flxWSPriceHis
         .Clear
         .Rows = 2
         .Cols = 7
         '               0    1      2         3     4         5          6
         .FormatString = "|^DATE |^REF. NO. |^COST |^PRICE |^P-RATIO |^P-MARGIN"
         
         .ColWidth(0) = 200
         .ColWidth(1) = (.Width - .ColWidth(0)) * 0.15  'DATE
         .ColWidth(2) = (.Width - .ColWidth(0)) * 0.16  'REF. NO.
         .ColWidth(3) = (.Width - .ColWidth(0)) * 0.16 'PREV. COST
         .ColWidth(4) = (.Width - .ColWidth(0)) * 0.16 'PREV. SRP
         .ColWidth(5) = (.Width - .ColWidth(0)) * 0.15 'RATIO
         .ColWidth(6) = (.Width - .ColWidth(0)) * 0.15 'MARGIN
    
    End With
    
      With flxKAPriceHis
         .Clear
         .Rows = 2
         .Cols = 8
         '               0    1      2         3     4         5          6
         .FormatString = "|^DATE |^REF. NO. |^COST |^PRICE |^P-RATIO |^P-MARGIN"
         
         .ColWidth(0) = 200
         .ColWidth(1) = (.Width - .ColWidth(0)) * 0.15  'DATE
         .ColWidth(2) = (.Width - .ColWidth(0)) * 0.16  'REF. NO.
         .ColWidth(3) = (.Width - .ColWidth(0)) * 0.16 'PREV. COST
         .ColWidth(4) = (.Width - .ColWidth(0)) * 0.16 'PREV. SRP
         .ColWidth(5) = (.Width - .ColWidth(0)) * 0.15 'RATIO
         .ColWidth(6) = (.Width - .ColWidth(0)) * 0.15 'MARGIN
    
    End With

End Sub
