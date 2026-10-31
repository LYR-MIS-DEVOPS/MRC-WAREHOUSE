VERSION 5.00
Object = "{B1CCF420-672E-11D0-B2C6-444553540000}#1.1#0"; "AniGif.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form LogIn 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log-In"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
   Icon            =   "WHausman.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "WHausman.frx":058A
   ScaleHeight     =   4035
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin AniGIF.AniGIF AniGIF1 
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Width           =   3735
      MajorVer        =   1
      MinorVer        =   1
      ExtentX         =   6588
      ExtentY         =   1085
      Transparent     =   -1  'True
      GIF             =   "WHausman.frx":131E4
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   5040
      Top             =   3120
   End
   Begin VB.TextBox txtUserID 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2880
      MaxLength       =   5
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2040
      Width           =   975
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   405
      Left            =   480
      TabIndex        =   4
      Top             =   3120
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   714
      Caption         =   "OK"
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
      Image           =   "WHausman.frx":1AB5B
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   405
      Left            =   1680
      TabIndex        =   5
      Top             =   3120
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   714
      Caption         =   "Close"
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
      Image           =   "WHausman.frx":1B0F5
      cBack           =   -2147483633
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Your One Stop Motorcycle Shop!"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   3780
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Enter User ID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   2040
      Width           =   2535
   End
End
Attribute VB_Name = "LogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Str1 As String


Private Sub cmdCancel_Click()
    cmdOK.Tag = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If txtUserID.Text = "" Then
        MsgBox "Please enter your User ID!", vbCritical + vbOKOnly, "Partsman"
        Exit Sub
    Else
        Call CheckID
    End If
End Sub

Private Sub Form_Load()
    AniGIF1.Play
    Str1 = "Your One Stop Motorcycle Shop!!!! "
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If cmdOK.Tag <> "OK" Then End
End Sub

Private Sub Timer1_Timer()
    lbl.Caption = Str1
    Str1 = Right$(Str1, Len(Str1) - 1) & Left$(Str1, 1)
End Sub

Private Sub CheckID()
    Dim rsUser As New ADODB.Recordset
    Dim Qry As String
    Dim UserFound As Boolean
    
    Qry = "SELECT PCUSER.EMPID From PCUSER " & _
        "WHERE PCUSER.PCUSERID='" & txtUserID.Text & "' AND PCUSER.WAREHOUSE=True;"
    With rsUser
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            'SPWarehouse.EmpId = !EmpId
            UserFound = True
        Else
            UserFound = False
        End If
        .Close
    End With
    If Not UserFound Then
        MsgBox "Invalid User ID!", vbExclamation + vbOKOnly, "Order"
    Else
        cmdOK.Tag = "OK"
        Unload Me
        SPWarehouse.Show
    End If

End Sub


Private Sub txtUserID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call CheckID
    End If
End Sub
