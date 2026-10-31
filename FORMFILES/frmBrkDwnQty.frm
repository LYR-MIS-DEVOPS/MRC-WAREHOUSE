VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBrkDwnQty 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   4560
   End
   Begin MSFlexGridLib.MSFlexGrid flxSPBreakDown 
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   7223
      _Version        =   393216
      Cols            =   3
      RowHeightMin    =   300
      FillStyle       =   1
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.TextBox txtSPID 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Tag             =   "xyf"
      Top             =   4320
      Width           =   1335
      _ExtentX        =   2355
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
      Image           =   "frmBrkDwnQty.frx":0000
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmBrkDwnQty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i       As Integer
Dim Qry     As String

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    On Error GoTo LocalErr
    Qry = "SELECT BRANCH, QTY FROM mrcwarehouse.spoverallqty where SPID='" & txtSPID.Text & "';"
    If rsAnyUse.State <> 0 Then rsAnyUse.Close
    Set rsAnyUse = LYRcon.Execute(Qry)
        With rsAnyUse
            If .RecordCount > 0 Then
                .MoveFirst
                flxSPBreakDown.Rows = 1
                Do While Not .EOF
                    flxSPBreakDown.AddItem vbTab & !Branch & vbTab & !Qty
                .MoveNext
                Loop
            End If
        End With
        Timer1.Enabled = False
        Call resizeFlex
    Exit Sub
LocalErr:
    MsgBox Err.Description, vbCritical, "Error Overall Quantity"
    'Call loadFlxRecord(rsAnyUse, flxSPBreakDown)
    
End Sub

Private Sub resizeFlex()
    With flxSPBreakDown
        .FormatString = " |<MRC BRANCH |>QTY"
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.5
        .ColWidth(2) = (.Width - 300) * 0.2
        
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 1) = "" Then
                .RemoveItem (i)
            End If
        Next i
    End With
End Sub


