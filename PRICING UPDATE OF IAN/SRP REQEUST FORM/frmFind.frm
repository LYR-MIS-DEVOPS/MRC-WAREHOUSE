VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FIND"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4785
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   4575
      Begin lvButton.lvButtons_H cmdOK 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   165
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "&OK"
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
         Image           =   "frmFind.frx":57E2
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   165
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "&CANCEL"
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
         Image           =   "frmFind.frx":5D7C
         cBack           =   -2147483633
      End
   End
   Begin VB.ComboBox cboSearch 
      Height          =   3885
      Left            =   120
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label lblCaption 
      Caption         =   "Search for "
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public RecordType As String
Public OwnerForm As String

Dim arrFind() As String
Dim arrIndex As Integer
Dim RecordKey As String
Dim varSearch As Boolean

Dim BranchCode As String
Dim BRefNo As String

Sub LoadRecord()
On Error GoTo er

    Dim tmpRs As New ADODB.Recordset
    cboSearch.Clear
    
    If RecordType = "BRANCH SRP REF. NO." Then
        Call modSRP.sp_branchSRPRefNo_list(tmpRs)
        If tmpRs.RecordCount <= 0 Then
            tmpRs.Close
            Exit Sub
        End If
        
        ReDim arrFind(tmpRs.RecordCount, 3) As String
        arrIndex = 0
        tmpRs.MoveFirst
        Do While Not tmpRs.EOF
            arrFind(arrIndex, 0) = tmpRs!BRefNo
            arrFind(arrIndex, 1) = tmpRs!BranchCode
            
            Call cboSearch.AddItem(tmpRs!Description, arrIndex)
            arrIndex = arrIndex + 1
            tmpRs.MoveNext
        Loop
        tmpRs.Close
    End If
    
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->LoadRecord()")
End Sub

Private Sub cboSearch_Change()
    Dim strPartial As String
    Dim strTotal As String
    Dim i As Integer
    Dim j As Integer
    
    If varSearch Then
        varSearch = False
        Exit Sub
    End If
    
    RecordKey = ""
    
    With cboSearch
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)    'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
             
            RecordKey = arrFind(i, 0)
            
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                varSearch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
    
End Sub

Private Sub cboSearch_Click()
    BRefNo = arrFind(cboSearch.ListIndex, 0)
    BranchCode = arrFind(cboSearch.ListIndex, 1)
End Sub

Private Sub cboSearch_DblClick()
    Call cmdOK_Click
End Sub

Private Sub cboSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdOK_Click
    Else
        KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    End If
End Sub

Private Sub cboSearch_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Or KeyCode = 46 Then
        varSearch = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Set frmFind = Nothing
End Sub

Private Sub cmdOK_Click()
On Error GoTo er

    
    'If Trim(RecordKey) = "" And cboSearch.ListCount > 1 Then
    If Trim(BranchCode) = "" And Trim(BRefNo) = "" And cboSearch.ListCount > 1 Then
    
        If cboSearch.ListIndex >= 0 Then
            BRefNo = arrFind(cboSearch.ListIndex, 0)
            BranchCode = arrFind(cboSearch.ListIndex, 1)
        Else
            MsgBox "No record of this " + RecordType + "!", vbExclamation, lblCaption
            Exit Sub
        End If
    End If
    
    If OwnerForm = "BSRN" Then
        Call frmSRPREQ.ShowRecord(BranchCode, BRefNo)
    End If
    
    Unload Me
    Set frmFind = Nothing
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdOK_Click()")
End Sub

Private Sub Form_Load()
On Error GoTo er
    OwnerForm = UCase(OwnerForm)
    RecordType = UCase(RecordType)
    
    lblCaption = "SEARCH FOR " & RecordType
    DoEvents
    Call LoadRecord
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->Form_Load()")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmFind = Nothing
End Sub
