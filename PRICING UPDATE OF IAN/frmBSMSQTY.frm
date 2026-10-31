VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBSMSQTY 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SIX MONTHS SALE QTY PER BRANCH"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid flxBSMSQTY 
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Tag             =   "XYZ"
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   8070
      _Version        =   393216
      Rows            =   14
      Cols            =   3
      RowHeightMin    =   315
      BackColor       =   16777215
      BackColorBkg    =   14737632
      HighLight       =   2
      GridLines       =   2
      FormatString    =   "|<BRANCH                                                                 |^QTY.                             "
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
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   4080
      TabIndex        =   1
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "TOTAL OF 6 MONTHS SALE QTY. :"
      Height          =   255
      Left            =   1320
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
Public Sub LoadRecord(pSPID As Integer)
On Error GoTo er
    Dim tmpRs As New ADODB.Recordset
    Dim sortby As String
    
    sortby = "BRANCH"
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
                flxBSMSQTY.TextMatrix(1, 2) = tmpRs!SMSQTY
            Else
                flxBSMSQTY.AddItem _
                vbTab & tmpRs!branchName & _
                vbTab & tmpRs!SMSQTY
            End If
            
        tmpRs.MoveNext
        
    Wend
    
    If tmpRs.State <> 0 Then tmpRs.Close
    
    Call total
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->LoadRecord()")
End Sub


Private Sub total()
    Dim i As Integer
    Dim qty As Integer
    Dim total As Integer
    
    For i = 1 To flxBSMSQTY.Rows - 1
        qty = flxBSMSQTY.TextMatrix(i, 2)
        total = total + qty
    Next
    
    lblTotal.Caption = total
End Sub
