VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBSMSQTY 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SIX MONTHS SALE QTY PER BRANCH"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid flxBSMSQTY 
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Tag             =   "XYZ"
      Top             =   240
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8070
      _Version        =   393216
      Rows            =   14
      Cols            =   5
      RowHeightMin    =   315
      BackColor       =   16777215
      BackColorBkg    =   14737632
      HighLight       =   2
      GridLines       =   2
      FormatString    =   "|<BRANCH                                                      |^QTY. ON HAND |^SALES QTY. |<LAST DATE-IN"
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
   Begin VB.Label lblTotal2 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   4800
      TabIndex        =   4
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "TOTAL:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   5040
      Width           =   3255
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   240
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   7365
      m_MDI_Height    =   5835
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "TOTAL OF 6 MONTHS SALE QTY. :"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   5040
      Width           =   2655
   End
End
Attribute VB_Name = "frmBSMSQTY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MySource As String

Public Sub LoadRecordPO(pSPID As Long, DateStart As Date, DateEnd As Date, pTYPE As String)
On Error GoTo er
    Dim tmpRs As New ADODB.Recordset
    Dim sortby As String
    
    sortby = "BRANCH"
    
    If pTYPE = "WAREHOUSE" Then
        Qry = "call sp_mrcwarehouse.sp_get_SMSQTY_LDI(" & pSPID & ",'" & Format(DateStart, "YYYY-MM-DD") & "','" & Format(DateEnd, "YYYY-MM-DD") & "','" & sortby & "')"
    Else
        Qry = "call sp_mrcwarehouse.sp_get_SMSQTY_LDI_ws(" & pSPID & ",'" & Format(DateStart, "YYYY-MM-DD") & "','" & Format(DateEnd, "YYYY-MM-DD") & "','" & sortby & "')"
    End If
    If tmpRs.State = adStateOpen Then tmpRs.Close
    tmpRs.Open Qry, LYRcon, adOpenStatic, adLockReadOnly
    Me.Caption = "QTYONHAND PER BRANCH"
    Label2.Visible = True
    Label1.Visible = False
    
    If tmpRs.State = 0 Then Exit Sub
    
    If tmpRs.RecordCount <= 0 Then
        tmpRs.Close
        flxBSMSQTY.Rows = 1
        flxBSMSQTY.Rows = 2
        Exit Sub
    End If
    

    flxBSMSQTY.Rows = 1
    flxBSMSQTY.Rows = 2

    tmpRs.MoveFirst
           
        While Not tmpRs.EOF
            If flxBSMSQTY.TextMatrix(1, 1) = "" Then
                flxBSMSQTY.TextMatrix(1, 1) = tmpRs!branchName
                flxBSMSQTY.TextMatrix(1, 2) = tmpRs!QtyOnHand
                flxBSMSQTY.TextMatrix(1, 3) = tmpRs!SMSQTY
                flxBSMSQTY.TextMatrix(1, 4) = IIf(IsNull(tmpRs!LASTDATEIN), "", tmpRs!LASTDATEIN)
            Else
                flxBSMSQTY.AddItem _
                vbTab & tmpRs!branchName & _
                vbTab & tmpRs!QtyOnHand & _
                vbTab & tmpRs!SMSQTY & _
                vbTab & IIf(IsNull(tmpRs!LASTDATEIN), "", tmpRs!LASTDATEIN)
            End If
            
        tmpRs.MoveNext
        
    Wend
    
    If tmpRs.State <> 0 Then tmpRs.Close
    
    Call Total
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->LoadRecordPO()")
End Sub

Public Sub LoadRecord(pSPID As Long)
On Error GoTo er
    Dim tmpRs As New ADODB.Recordset
    Dim sortby As String
    
    sortby = "BRANCH"
    Label1.Visible = True
    Label2.Visible = False
    Call modSRP.sp_get_SMSQTY_LDI(pSPID, sortby, tmpRs)
    If tmpRs.State = 0 Then Exit Sub
    
    If tmpRs.RecordCount <= 0 Then
        tmpRs.Close
        flxBSMSQTY.Rows = 1
        flxBSMSQTY.Rows = 2
        Exit Sub
    End If
    

    flxBSMSQTY.Rows = 1
    flxBSMSQTY.Rows = 2

    tmpRs.MoveFirst
           
        While Not tmpRs.EOF
            If flxBSMSQTY.TextMatrix(1, 1) = "" Then
                flxBSMSQTY.TextMatrix(1, 1) = tmpRs!branchName
                flxBSMSQTY.TextMatrix(1, 2) = tmpRs!QtyOnHand
                flxBSMSQTY.TextMatrix(1, 3) = tmpRs!SMSQTY
                flxBSMSQTY.TextMatrix(1, 4) = tmpRs!LASTDATEIN
            Else
                flxBSMSQTY.AddItem _
                vbTab & tmpRs!branchName & _
                vbTab & tmpRs!QtyOnHand & _
                vbTab & tmpRs!SMSQTY & _
                vbTab & tmpRs!LASTDATEIN
            End If
            
        tmpRs.MoveNext
        
    Wend
    
    If tmpRs.State <> 0 Then tmpRs.Close
    
    Call Total
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->LoadRecord()")
End Sub


Private Sub Total()
    Dim i As Integer
    Dim qty As Integer, salesqty As Integer
    Dim Total As Integer, salesTTL As Integer
    
    For i = 1 To flxBSMSQTY.Rows - 1
        qty = flxBSMSQTY.TextMatrix(i, 2)
        salesqty = flxBSMSQTY.TextMatrix(i, 3)
        Total = Total + qty
        salesTTL = salesTTL + salesqty
    Next
    
    lblTotal.Caption = Total
    lblTotal2.Caption = salesTTL
End Sub


Private Sub RXResizer1_ResizeComplete()
        DoEvents
        With flxBSMSQTY
            .Cols = 5
            .FixedCols = 1
            
            '               0    1       2
            .FormatString = "|<BRANCH |^QTY. ON HAND |^SALES QTY. |<LAST DATE-IN"
            
            .ColWidth(0) = 10
            .ColWidth(1) = (.Width - .ColWidth(0)) * 0.4  'BRANCH
            .ColWidth(2) = (.Width - .ColWidth(0)) * 0.19 'QTY. ON HAND
            .ColWidth(3) = (.Width - .ColWidth(0)) * 0.18 'SALES QTY.
            .ColWidth(4) = (.Width - .ColWidth(0)) * 0.19 'LAST DATE-IN
        End With
End Sub
