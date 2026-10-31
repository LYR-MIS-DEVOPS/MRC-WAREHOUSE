VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form BrnBBOref 
   Caption         =   "BBO Reference"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid flxREF 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColorBkg    =   -2147483633
      AllowUserResizing=   1
      FormatString    =   "<P.O. REF#.   |^P.O. DATE     |<BRANCH                 |<B.O. QTY.|^ORDER    "
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Tag             =   "xyf"
      Top             =   2200
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
      Image           =   "BrnBBOref.frx":0000
      cBack           =   -2147483633
   End
   Begin VB.Label lblTotal 
      Caption         =   "TOTAL ORDER:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   2535
   End
End
Attribute VB_Name = "BrnBBOref"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal X As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Private Const HWND_NOTOPMOST = -2


Private Sub cmdOK_Click()
    BrnBackOrderPO.LoadRef = False
    Unload Me
End Sub

Private Sub Form_Activate()
    SetWindowPos Me.hwnd, HWND_TOPMOST, _
        0, 0, 0, 0, Flags

End Sub

