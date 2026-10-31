VERSION 5.00
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSRPREQ 
   Caption         =   "SRP REQUEST FORM"
   ClientHeight    =   9030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15285
   ControlBox      =   0   'False
   Icon            =   "frmSRPREQ.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9030
   ScaleWidth      =   15285
   WindowState     =   2  'Maximized
   Begin lvButton.lvButtons_H cmdApp 
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   661
      Caption         =   "&APPROVE"
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
      Image           =   "frmSRPREQ.frx":000C
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdFind 
      Height          =   375
      Left            =   4200
      TabIndex        =   26
      Top             =   120
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   661
      Caption         =   "&FIND"
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
      Image           =   "frmSRPREQ.frx":05A6
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   661
      Caption         =   "&PRINT"
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
      Image           =   "frmSRPREQ.frx":0B40
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdSum 
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   661
      Caption         =   "&SUMMARY"
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
      Image           =   "frmSRPREQ.frx":6332
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   120
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "&EXIT FORM"
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
      Image           =   "frmSRPREQ.frx":68CC
      cBack           =   -2147483633
   End
   Begin VB.Frame fraHoldBottom 
      Height          =   6735
      Left            =   120
      TabIndex        =   10
      Tag             =   "XYZ"
      Top             =   2160
      Width           =   15015
      Begin VB.TextBox txtNotedBy 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   11640
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Tag             =   "XYZ"
         Top             =   5640
         Width           =   3255
      End
      Begin MSFlexGridLib.MSFlexGrid flxSRP 
         Height          =   4215
         Left            =   120
         TabIndex        =   17
         Tag             =   "XYZ"
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   7435
         _Version        =   393216
         Cols            =   11
         RowHeightMin    =   315
         BackColor       =   16777215
         BackColorBkg    =   12632256
         HighLight       =   2
         GridLines       =   2
         AllowUserResizing=   1
         FormatString    =   "  |<PARTNUMBER |<DESCRIPTION |^CURRENT SRP |^NEW SRP  |^COMPETITOR'S PRICE |<COMPETITOR'S NAME |^CURRENT WSP |^NEW WSP"
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
      Begin VB.TextBox txtRecBy 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Tag             =   "XYZ"
         Top             =   6135
         Width           =   3255
      End
      Begin VB.TextBox txtPrepBy 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Tag             =   "XYZ"
         Top             =   5640
         Width           =   3255
      End
      Begin VB.TextBox txtRemarks 
         BackColor       =   &H00C0C0C0&
         Height          =   735
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Tag             =   "XYZ"
         Top             =   4560
         Width           =   6015
      End
      Begin VB.TextBox txtAppBy 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   11640
         Locked          =   -1  'True
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Tag             =   "XYZ"
         Text            =   "LORENZO C. RUBINOS JR."
         Top             =   6120
         Width           =   3255
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NOTED BY:"
         Height          =   255
         Left            =   9840
         TabIndex        =   23
         Tag             =   "XYZ"
         Top             =   5640
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "RECOMMENDED BY:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Tag             =   "XYZ"
         Top             =   6120
         Width           =   1695
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "REMARKS:"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Tag             =   "XYZ"
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PREPARED BY:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Tag             =   "XYZ"
         Top             =   5640
         Width           =   1695
      End
      Begin VB.Label lblAppBy 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "APPROVED BY:"
         Height          =   255
         Left            =   10080
         TabIndex        =   14
         Tag             =   "XYZ"
         Top             =   6120
         Width           =   1455
      End
   End
   Begin VB.Frame fraHoldTop 
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Tag             =   "XYZ"
      Top             =   480
      Width           =   15015
      Begin VB.TextBox txtSupName 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   1000
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Tag             =   "XYZ"
         Top             =   735
         Width           =   3855
      End
      Begin VB.TextBox txtBranchName 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   1000
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Tag             =   "XYZ"
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox txtStat 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Tag             =   "XYZ"
         Top             =   1110
         Width           =   2055
      End
      Begin VB.TextBox txtRefDate 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   12720
         Locked          =   -1  'True
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Tag             =   "XYZ"
         Top             =   735
         Width           =   1935
      End
      Begin VB.TextBox txtSRPRefNo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   12720
         Locked          =   -1  'True
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Tag             =   "XYZ"
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblStat 
         BackStyle       =   0  'Transparent
         Caption         =   "STATUS:"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Tag             =   "XYZ"
         Top             =   1110
         Width           =   735
      End
      Begin VB.Label lblStartBookNo 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER NAME:"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Tag             =   "XYZ"
         Top             =   735
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "REF. NO."
         Height          =   255
         Left            =   11640
         TabIndex        =   8
         Tag             =   "XYZ"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblPckngDate 
         BackStyle       =   0  'Transparent
         Caption         =   "REF. DATE:"
         Height          =   255
         Left            =   11640
         TabIndex        =   7
         Tag             =   "XYZ"
         Top             =   735
         Width           =   1095
      End
      Begin VB.Label lblTerms 
         BackStyle       =   0  'Transparent
         Caption         =   "BRANCH NAME:"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Tag             =   "XYZ"
         Top             =   360
         Width           =   1455
      End
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   9600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   15405
      m_MDI_Height    =   9510
      OriginalBorder  =   2
      OriginalCaption =   25
   End
   Begin VB.Menu mnuBSRP 
      Caption         =   "BSRP"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmSRPREQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmpRs As New ADODB.Recordset
Dim DtlsCol As Integer
Dim DtlsRow As Integer

Dim Ans As String

Dim BranchCode As String


Public Sub ShowApprovedRecord(pBREFNO As String)
On Error GoTo er
    
    Screen.MousePointer = vbHourglass
    

    Call modSRP.sp_show_approved_branchSRP(pBREFNO, tmpRs)
    If tmpRs.State = 0 Then Exit Sub

    If tmpRs.RecordCount <= 0 Then
        flxSRP.Rows = 1
        flxSRP.Rows = 2
        Screen.MousePointer = vbDefault
        tmpRs.Close
        Exit Sub
    End If

    
    flxSRP.Rows = 1
    flxSRP.Rows = 2
    
' 0        1            2             3          4               5                     6                 7          8           9          10
'"  |<PARTNUMBER |<DESCRIPTION |^CURRENT SRP |^NEW SRP  |^COMPETITOR'S PRICE |<COMPETITOR'S NAME |^CURRENT WSP |^NEW WSP  |SPAREPARTID |SUPPLIER"
   
    tmpRs.MoveFirst

            txtSupName.Text = tmpRs!Supplier
            txtSRPRefNo.Text = tmpRs!BRefNo
            txtRefDate.Text = tmpRs!DateEncode
            txtBranchName.Text = tmpRs!branchName
            txtRemarks.Text = tmpRs!Remarks
            txtStat.Text = tmpRs!Status
            txtPrepBy.Text = tmpRs!EncodedBy
            txtRecBy.Text = tmpRs!RecomBy
            txtNotedBy.Text = tmpRs!NotedBy
            txtAppBy.Text = tmpRs!ApprovedBy
            BranchCode = tmpRs!BranchCode
            
        While Not tmpRs.EOF
            If flxSRP.TextMatrix(1, 1) = "" Then
                flxSRP.TextMatrix(1, 1) = tmpRs!PARTNUMBER
                flxSRP.TextMatrix(1, 2) = tmpRs!Description
                flxSRP.TextMatrix(1, 3) = Format(tmpRs!OldSRP, "#,##0")
                flxSRP.TextMatrix(1, 4) = Format(tmpRs!NewSRP, "#,##0")
                flxSRP.TextMatrix(1, 5) = Format(tmpRs!CompePrice, "#,##0")
                flxSRP.TextMatrix(1, 6) = tmpRs!CompeName
                flxSRP.TextMatrix(1, 7) = Format(tmpRs!OldWSP, "#,##0")
                flxSRP.TextMatrix(1, 8) = Format(tmpRs!NewWSP, "#,##0")
                flxSRP.TextMatrix(1, 9) = tmpRs!SPAREPARTID

            Else
                flxSRP.AddItem _
                vbTab & tmpRs!PARTNUMBER & _
                vbTab & tmpRs!Description & _
                vbTab & Format(tmpRs!OldSRP, "#,##0") & _
                vbTab & Format(tmpRs!NewSRP, "#,##0") & _
                vbTab & Format(tmpRs!CompePrice, "#,##0") & _
                vbTab & tmpRs!CompeName & _
                vbTab & Format(tmpRs!OldWSP, "#,##0") & _
                vbTab & Format(tmpRs!NewWSP, "#,##0") & _
                vbTab & tmpRs!SPAREPARTID

            End If
        
    tmpRs.MoveNext
    Wend
    
    If flxSRP.Rows = 1 Then
        flxSRP.Rows = 2
    End If
    
    If tmpRs.State <> 0 Then tmpRs.Close
    
    
    Screen.MousePointer = vbDefault
    
    Call EnabledSelectedButton
    
Exit Sub
er:
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->ShowApprovedRecord()")
End Sub

Public Sub EnabledSelectedButton()

    cmdApp.Enabled = False 'Approve Button
    
    If txtStat.Text = "APPROVED" And txtSupName.Text <> "" Then
        cmdApp.Enabled = False   'Approve Button
        cmdPrint.Enabled = True  'PRINT Button

    ElseIf txtStat.Text = "FOR APPROVAL" And txtSupName.Text <> "" Then
        cmdApp.Enabled = True   'Approve Button
        cmdPrint.Enabled = True  'PRINT Button
        
    Else
        cmdApp.Enabled = False   'Approve Button
        cmdPrint.Enabled = False 'PRINT Button
    End If
 
End Sub


Public Sub ShowRecord(pBRANCHCODE As String, pBREFNO As String)
On Error GoTo er
    
    Screen.MousePointer = vbHourglass
    

    Call modSRP.sp_show_branchSRP(pBRANCHCODE, pBREFNO, tmpRs)
    If tmpRs.State = 0 Then Exit Sub

    If tmpRs.RecordCount <= 0 Then
        flxSRP.Rows = 1
        flxSRP.Rows = 2
        Screen.MousePointer = vbDefault
        tmpRs.Close
        Exit Sub
    End If

    
    flxSRP.Rows = 1
    flxSRP.Rows = 2
    
' 0        1            2             3          4               5                     6                 7          8           9          10
'"  |<PARTNUMBER |<DESCRIPTION |^CURRENT SRP |^NEW SRP  |^COMPETITOR'S PRICE |<COMPETITOR'S NAME |^CURRENT WSP |^NEW WSP  |SPAREPARTID |SUPPLIER"
   
    tmpRs.MoveFirst

            txtSupName.Text = tmpRs!Supplier
            txtSRPRefNo.Text = tmpRs!BRefNo
            txtRefDate.Text = tmpRs!DateEncode
            txtBranchName.Text = tmpRs!branchName
            txtRemarks.Text = tmpRs!Remarks
            txtStat.Text = tmpRs!Status
            txtPrepBy.Text = tmpRs!EncodedBy
            txtRecBy.Text = tmpRs!RecomBy
            txtNotedBy.Text = tmpRs!NotedBy
            txtAppBy.Text = tmpRs!ApprovedBy
            BranchCode = tmpRs!BranchCode
            
        While Not tmpRs.EOF
            If flxSRP.TextMatrix(1, 1) = "" Then
                flxSRP.TextMatrix(1, 1) = tmpRs!PARTNUMBER
                flxSRP.TextMatrix(1, 2) = tmpRs!Description
                flxSRP.TextMatrix(1, 3) = Format(tmpRs!OldSRP, "#,##0")
                flxSRP.TextMatrix(1, 4) = Format(tmpRs!NewSRP, "#,##0")
                flxSRP.TextMatrix(1, 5) = Format(tmpRs!CompePrice, "#,##0")
                flxSRP.TextMatrix(1, 6) = tmpRs!CompeName
                flxSRP.TextMatrix(1, 7) = Format(tmpRs!OldWSP, "#,##0")
                flxSRP.TextMatrix(1, 8) = Format(tmpRs!NewWSP, "#,##0")
                flxSRP.TextMatrix(1, 9) = tmpRs!SPAREPARTID

            Else
                flxSRP.AddItem _
                vbTab & tmpRs!PARTNUMBER & _
                vbTab & tmpRs!Description & _
                vbTab & Format(tmpRs!OldSRP, "#,##0") & _
                vbTab & Format(tmpRs!NewSRP, "#,##0") & _
                vbTab & Format(tmpRs!CompePrice, "#,##0") & _
                vbTab & tmpRs!CompeName & _
                vbTab & Format(tmpRs!OldWSP, "#,##0") & _
                vbTab & Format(tmpRs!NewWSP, "#,##0") & _
                vbTab & tmpRs!SPAREPARTID

            End If
        
    tmpRs.MoveNext
    Wend
    
    If flxSRP.Rows = 1 Then
        flxSRP.Rows = 2
    End If
    
    If tmpRs.State <> 0 Then tmpRs.Close
    
    
    Screen.MousePointer = vbDefault
    
    Call EnabledSelectedButton
    
Exit Sub
er:
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->ShowRecord()")
End Sub

Function ValidateEntry() As String
    Dim vMsg As String, Qry As String
    Dim i As Integer, Cntr As Integer
    
    vMsg = ""
    
    If txtRecBy.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field recommended by."
    If txtNotedBy.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field noted by."
    If txtAppBy.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field approved by."
    
    
' 0        1            2             3          4               5                     6                 7          8           9          10
'"  |<PARTNUMBER |<DESCRIPTION |^CURRENT SRP |^NEW SRP  |^COMPETITOR'S PRICE |<COMPETITOR'S NAME |^CURRENT WSP |^NEW WSP  |SPAREPARTID |SUPPLIER"
 
   For i = 1 To flxSRP.Rows - 1
        If flxSRP.TextMatrix(i, 1) = "" And flxSRP.TextMatrix(i, 2) = "" Then vMsg = vMsg & vbCrLf & "Please delete blank row!"
   Next
   
   ValidateEntry = vMsg
    
End Function



Sub flxSRPColsChanger()
On Error Resume Next
Dim i As Integer
    With flxSRP
        .Clear
        .Rows = 2
        .Cols = 10
        
        '               0         1            2             3          4               5                     6                 7          8
        .FormatString = "  |<PARTNUMBER |<DESCRIPTION |^CURRENT SRP |^NEW SRP  |^COMPETITOR'S PRICE |<COMPETITOR'S NAME |^CURRENT WSP |^NEW WSP"
        
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.12  'PART NUMBER
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.2  'DESCRIPTION
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.1  'CURRENT SRP
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.1  'NEW SRP
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.1 'COMPETITOR'S PRICE
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.15 'COMPETITOR'S NAME
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.1 'CURRENT WSP
        .ColWidth(8) = (.Width - .ColWidth(0)) * 0.1 'NEW WSP
        .ColWidth(9) = 0 'SPAREPARTSID

     End With
    
End Sub


Sub ClearEntryForm(Focus1stObject As Boolean)
   
    txtSupName.Text = ""
    txtSRPRefNo.Text = ""
    txtRefDate.Text = ""
    txtBranchName.Text = ""
    
    txtRemarks.Text = ""
    txtPrepBy.Text = ""
    txtNotedBy.Text = ""
    txtRecBy.Text = ""
    txtAppBy.Text = ""
    
    flxSRP.Rows = 1
    flxSRP.Rows = 2
                              
    
    cmdApp.Enabled = False
    cmdPrint.Enabled = False
End Sub

Private Sub cmdApp_Click()
On Error GoTo er
    Dim Level As String
    Dim message As String
    Dim DateEncoded As String
    
    If ValidateEntry() <> "" Then
        MsgBox ValidateEntry, vbExclamation, Me.Caption
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass

    Call CreateTempTable
        
    Ans = MsgBox("DO YOU WANT TO APPROVE?", vbQuestion + vbYesNo, Me.Caption)
    If Ans = vbNo Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    DateEncoded = Format(txtRefDate.Text, "YYYY-mm-dd")
    Call modSRP.sp_branchsrp_approve(txtSRPRefNo.Text, DateEncoded, txtRemarks.Text, txtPrepBy.Text, txtRecBy.Text, txtNotedBy.Text, txtAppBy.Text, BranchCode, Level, message)

    Screen.MousePointer = vbDefault
    
    If Trim(Level) = "" Then
        Exit Sub
    ElseIf Level = "Granted" Or Level = "Warning" Then
        If Level = "Granted" Then
            MsgBox message, vbInformation, "SAVE."
            txtStat.Text = "APPROVED"
        Else
            MsgBox message, vbExclamation, "SAVE!"
            Exit Sub
        End If
    Else
        MsgBox message, vbCritical, "Error!"
        Exit Sub
    End If
    
    Call EnabledSelectedButton

Exit Sub
er:
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "ERROR->cmdApp_Click()")
End Sub

Private Sub CreateTempTable()
On Error GoTo er
        Dim Qry As String

        Qry = "Drop Temporary Table IF Exists sp_srppricing.tmp_srppricingdtls;"
        Call LYRcon.Execute(Qry)
        DoEvents
        Qry = "Create Temporary Table IF Not Exists sp_srppricing.tmp_srppricingdtls( " & _
        "   BREFNO              VARCHAR(20),        " & _
        "   SPAREPARTID         INT(10),            " & _
        "   PARTNUMBER          VARCHAR(50),        " & _
        "   OLDSRP              DECIMAL(16,2),      " & _
        "   NEWSRP              DECIMAL(16,2),      " & _
        "   CompePrice          DECIMAL(16,2),      " & _
        "   CompeName           VARCHAR(100),       " & _
        "   OLDWSP              DECIMAL(16,2),      " & _
        "   NEWWSP              DECIMAL(16,2),      " & _
        "   SUPPLIER            LONGTEXT            " & _
        ");"

        'Text1.Text = Qry
        Call LYRcon.Execute(Qry)
        
        
        DoEvents
        Call InsertTempTable
        
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error ->CreatTempTable()")
End Sub



Private Sub InsertTempTable()
On Error GoTo er
        Dim Qry As String
        Dim i As Integer

' 0        1            2             3          4               5                     6                 7          8           9          10
'"  |<PARTNUMBER |<DESCRIPTION |^CURRENT SRP |^NEW SRP  |^COMPETITOR'S PRICE |<COMPETITOR'S NAME |^CURRENT WSP |^NEW WSP  |SPAREPARTID |SUPPLIER"

        Qry = "Insert Into sp_srppricing.tmp_srppricingdtls(BRefNo, PartNumber, OldSRP, NewSRP, CompePrice, CompeName, OldWSP, NewWSP, SparepartID, SUPPLIER) Values"

        For i = 1 To flxSRP.Rows - 1
            DoEvents
            If i > 1 Then Qry = Qry & ","
            
        Qry = Qry & "('" & txtSRPRefNo.Text & "','" & Trim(flxSRP.TextMatrix(i, 1)) & _
                        "','" & Val(Replace(Trim(flxSRP.TextMatrix(i, 3)), ",", "")) & "','" & Val(Replace(Trim(flxSRP.TextMatrix(i, 4)), ",", "")) & "','" & Val(Replace(Trim(flxSRP.TextMatrix(i, 5)), ",", "")) & _
                        "','" & Trim(flxSRP.TextMatrix(i, 6)) & "','" & Val(Replace(Trim(flxSRP.TextMatrix(i, 7)), ",", "")) & _
                        "','" & Val(Replace(Trim(flxSRP.TextMatrix(i, 8)), ",", "")) & "','" & Trim(flxSRP.TextMatrix(i, 9)) & "','" & txtSupName.Text & "')"

        Next
        'Text1.Text = Qry
        'InputBox Qry, Qry, Qry
        Call LYRcon.Execute(Qry)
        
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error ->InsertTempTable()")
End Sub
Private Sub cmdExit_Click()
    Ans = MsgBox("ARE YOU SURE YOU WANT TO EXIT?", vbQuestion + vbYesNo, Me.Caption)
    If Ans = vbNo Then Exit Sub

    Unload Me
End Sub

Private Sub cmdFind_Click()
    frmFind.RecordType = "BRANCH SRP REF. NO."
    frmFind.OwnerForm = "BSRN"
    frmFind.Show vbModal
End Sub



Private Sub cmdPrint_Click()
    Dim Qry As String
    
    Screen.MousePointer = vbHourglass
    Dim rpt As New rptBranchSRP


     Qry = "Call sp_srppricing.sp_show_branchSRP('" & BranchCode & "','" & txtSRPRefNo.Text & "');"
     
     DoEvents
     If tmpRs.State <> 0 Then tmpRs.Close
     Set tmpRs = LYRcon.Execute(Qry)


     DoEvents
     With rpt
         .DiscardSavedData
          DoEvents
         .Database.SetDataSource tmpRs
         .ReadRecords
     End With

    frmPrint.CRViewer1.ReportSource = rpt
    frmPrint.CRViewer1.ViewReport
    frmPrint.CRViewer1.Zoom 150
    Screen.MousePointer = vbDefault

    DoEvents
    frmPrint.Caption = "PRINT" & " " & "-" & " " & "[" & "BRANCH REQUEST FORM" & "]"
    frmPrint.MySource = "SRP Request form"
    frmPrint.WindowState = vbMaximized
    frmPrint.Show vbModal

End Sub

Private Sub cmdSum_Click()
    FrmSummary.mySrc = "SRP Request form"
    FrmSummary.Show
    Unload Me
End Sub

Private Sub RXResizer1_ResizeComplete()
    With flxSRP
        .Cols = 10

        '               0         1            2             3          4               5                     6                 7          8
        .FormatString = "  |<PARTNUMBER |<DESCRIPTION |^CURRENT SRP |^NEW SRP  |^COMPETITOR'S PRICE |<COMPETITOR'S NAME |^CURRENT WSP |^NEW WSP"
        
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.12  'PART NUMBER
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.2  'DESCRIPTION
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.1  'CURRENT SRP
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.1  'NEW SRP
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.1 'COMPETITOR'S PRICE
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.15 'COMPETITOR'S NAME
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.1 'CURRENT WSP
        .ColWidth(8) = (.Width - .ColWidth(0)) * 0.1 'NEW WSP
        .ColWidth(9) = 0 'SPAREPARTSID

     End With
     
End Sub


Private Sub txtAppBy_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub


