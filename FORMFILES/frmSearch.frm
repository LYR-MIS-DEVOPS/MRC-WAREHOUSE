VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   4320
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   3240
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   3240
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboSearch 
      Height          =   3105
      Left            =   120
      Style           =   1  'Simple Combo
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   3015
      Begin lvButton.lvButtons_H cmdOK 
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Caption         =   "&Ok"
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
         Image           =   "frmSearch.frx":1CFA
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   1560
         TabIndex        =   4
         Top             =   160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Caption         =   "&Cancel"
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
         Image           =   "frmSearch.frx":2294
         cBack           =   -2147483633
      End
   End
   Begin VB.Label lblCap 
      Caption         =   "Search for"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Public FrmSource As Form
Public QrySearch As String
Dim MySearchID As Integer
Dim varSearch As Boolean

Private Sub cboSearch_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    
    If varSearch Then
        varSearch = False
        Exit Sub
    End If
    With cboSearch
        strPartial = .Text
        i = SendMessage(.hWnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rs.MoveFirst
            rs.Move i
            MySearchID = IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0))
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                varSearch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            MySearchID = 0
        End If
    End With
End Sub

Private Sub cboSearch_Click()
    Call cboSearch_Change
End Sub

Private Sub cboSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            varSearch = True
        Case vbKeyBack
            varSearch = True
    End Select
End Sub

Private Sub cboSearch_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdOK.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If MySearchID = 0 Then
        MsgBox "Please " & lblCap.Caption & " from the list provided."
        Exit Sub
    End If
    If lblCap.Caption = "Search for MC Engine" Then
        Call FrmSource.ShowERecord(Val(MySearchID))
    Else
        Call FrmSource.ShowRecord(Val(MySearchID))
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    
    rs.Open QrySearch, LYRcon, adOpenKeyset, adLockOptimistic
    'Loading data for search
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                If Not IsNull(.Fields(1)) Then
                    cboSearch.AddItem .Fields(1)
                End If
                .MoveNext
            Loop
        End If
    End With
    'cboSearch.SetFocus

    If FrmSource.Caption = "Customer Service Record" Then
        Me.Width = Me.Width * 1.5
        Me.cboSearch.Width = Me.cboSearch.Width * 1.5
        Me.frame1.Left = (Me.Width - Me.frame1.Width) / 2
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    FrmSource.Enabled = True
    Set rs = Nothing
    Set frmSearch = Nothing
End Sub
