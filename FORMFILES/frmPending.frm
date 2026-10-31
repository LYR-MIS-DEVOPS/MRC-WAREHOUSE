VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPending 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PENDING PACKING NO."
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid flxprn 
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      _Version        =   393216
   End
   Begin VB.TextBox txtDRASNO 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtReason 
      Height          =   1335
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   4695
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1920
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   661
      Caption         =   "&Save"
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
      Image           =   "frmPending.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   4500
      TabIndex        =   3
      Top             =   1920
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   661
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
      Image           =   "frmPending.frx":059A
      cBack           =   -2147483633
   End
   Begin VB.Label Label2 
      Caption         =   "DRAS #:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "REASON:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "frmPending"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pNO As String, MySource As String
Dim ans As VbMsgBoxResult
Dim tmpRs As New ADODB.Recordset
Dim Qry As String, PrnVal(0 To 12) As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo er

    Dim level As String
    Dim message As String, chkExists As String
         
        ans = MsgBox("DO YOU WANT TO SAVE YOUR CHANGES?", vbQuestion + vbYesNo, Me.Caption)
        If ans = vbNo Then Exit Sub
        
        If tmpRs.State = adStateOpen Then tmpRs.Close
        
        If MySource = "frmAcctPayInvcg" Then
            If txtReason.Text = "" Then
                MsgBox "Incomplete entry on reason.", vbExclamation, "Validation"
                Exit Sub
            End If
            Label2.Visible = False
            txtDRASNO.Visible = False
            Qry = "UPDATE mrcwarehouse.supplierinvoice SET INVSTATUS = 'PENDING', PENDDATE = Now(),PENDREASON = '" + txtReason.Text + "' WHERE PACKINGNO = '" + pNO + "';"
            If tmpRs.State = adStateOpen Then tmpRs.Close
            tmpRs.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    
            frmAcctPayInvcg.cboStat.ListIndex = 0
            frmAcctPayInvcg.EnabledSelectedButton
        ElseIf MySource = "SPordering" Then
        
            If txtReason.Text = "" Then
                MsgBox "Incomplete entry on reason.", vbExclamation, "Validation"
                Exit Sub
            ElseIf txtDRASNO.Text = "" Then
                MsgBox "Incomplete entry on Discrepancy Report and Adjustment Slip(DRAS) No.", vbExclamation, "Validation"
                Exit Sub
            End If
            
            Qry = "SELECT DRASNO From rollbacktrans WHERE DRASNO='" & txtDRASNO.Text & "';"
            
            If CtrlNumExist(Qry) Then
                MsgBox "Discrepancy Report and Adjustment Slip No. " & txtDRASNO.Text & " is already exist!", vbExclamation + vbOKOnly, "Validation Error"
                Exit Sub
            End If
        
            chkExists = ""
            
            chkExists = GetFieldValue("Select RBID From rollbacktrans Where TRANSREFERENCE = '" & pNO & "' and RBSTATUS = 'PENDING' ")
            If chkExists <> "" Then
                MsgBox "This PO#: " & pNO & " is already process, Please wait for the approval!", vbInformation + vbOKOnly, "Validation"
                Exit Sub
            End If
        
            ans = MsgBox("Please Print this DRAS to finalized and for Approval!", vbQuestion + vbYesNo, "Print")
            If ans = vbNo Then
                Qry = "call sp_mrcwarehouse.sp_rollback_process('Purchase Order','" & pNO & "','" & txtReason.Text & "','PENDING'," & Val(UserID) & ",'" & txtDRASNO.Text & "')"
                
                If tmpRs.State = adStateOpen Then tmpRs.Close
                tmpRs.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                
                If tmpRs.RecordCount > 0 Then
                    tmpRs.MoveFirst
                    MsgBox tmpRs!message, vbInformation + vbOKOnly, tmpRs!level
                End If
                tmpRs.Close
                
                LYRcon.Execute "Insert Into transhistory(`DATE`, TRANSNAME, REFERENCE, `STATUS`, RESPONSIBLE ) " & _
                    "VALUES('" & Format(Now(), "yyyy-mm-dd hh:mm:ss") & "','ROLLBACK','" & pNO & "','PENDING'," & Val(UserID) & ");"
                Exit Sub
            End If
            PrnVal(0) = "TYPE: DISCR. REPORT AND ADJ. SLIP"
            PrnVal(1) = "DATE: " & Format(Now(), "MM/DD/YYYY")
            PrnVal(2) = "REF#: " & txtDRASNO.Text
            PrnVal(3) = ""
            PrnVal(4) = ""
            PrnVal(5) = ""
            PrnVal(6) = ""
            PrnVal(7) = GetFieldValue("Select CONCAT(FIRSTNAME, ' ', LASTNAME) as EMPLOYEE From employees Where EMPID = " & Val(UserID) & " ")
            PrnVal(8) = GetFieldValue("Select `VALUE` From settings where REMARKS = 'Approved By' and description = 'SPTRANSFER' ")
            PrnVal(9) = ""
            PrnVal(10) = "SUBMITTED BY"
            PrnVal(11) = "APPROVED BY"
            PrnVal(12) = IIf(txtReason.Text <> "", "REASON: " & txtReason.Text, "")
            Call PrintMyForm(PrnVal, "DISCR REPORT AND ADJ SLIP", flxprn)

            Qry = "call sp_mrcwarehouse.sp_rollback_process('Purchase Order','" & pNO & "','" & txtReason.Text & "','FOR APPROVAL'," & Val(UserID) & ",'" & txtDRASNO.Text & "')"
            
            If tmpRs.State = adStateOpen Then tmpRs.Close
            tmpRs.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            
            If tmpRs.RecordCount > 0 Then
                tmpRs.MoveFirst
                MsgBox tmpRs!message, vbInformation + vbOKOnly, tmpRs!level
            End If
            tmpRs.Close
            
            LYRcon.Execute "Insert Into transhistory(`DATE`, TRANSNAME, REFERENCE, `STATUS`, RESPONSIBLE ) " & _
                "VALUES('" & Format(Now(), "yyyy-mm-dd hh:mm:ss") & "','ROLLBACK','" & pNO & "','FOR APPROVAL'," & Val(UserID) & ");"
            
        End If
        Unload Me
Exit Sub
er:
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdSave_Click()")
End Sub

Private Sub Form_Load()
 
    If MySource = "frmAcctPayInvcg" Then
        Me.Caption = "PENDING PACKING NO."
        Label2.Visible = False
        txtDRASNO.Visible = False
    ElseIf MySource = "SPordering" Then
        Me.Caption = "PENDING PURCHASE ORDER"
        Label2.Visible = True
        txtDRASNO.Visible = True
    End If
End Sub

Private Sub txtDRASNO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtReason.SetFocus
End Sub

Private Sub txtReason_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdSave_Click
End Sub
