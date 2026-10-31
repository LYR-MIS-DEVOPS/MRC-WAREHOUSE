Attribute VB_Name = "modPMSummary"
Option Explicit
Private Qry As String
Private rs As New ADODB.Recordset

Public Sub GetMatrixID(matrixID As Integer, objRs As ADODB.Recordset)
On Error GoTo er

    DoEvents
    Qry = "Select partsmatrix.PMATRIXID from mrc2retail.partsmatrix where partsmatrix.PMATRIXID = " & matrixID & " "
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->GetMatrixID()")
End Sub

Public Sub sp_dataentry_summary(objRs As ADODB.Recordset)
On Error GoTo er

   'PROCEDURE sp_partsmatrix.sp_dataentry_summary(
   'pBRANCHCODE VarChar(4)
   ')
    DoEvents
    Qry = "Call sp_partsmatrix.sp_dataentry_summary('" & sysBranchCode & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_dataentry_summary()")

End Sub

Public Sub sp_brnchQty_summary(objRs As ADODB.Recordset)
On Error GoTo er

    'PROCEDURE sp_partsmatrix.sp_dataentry_update(
    '  IN pBRANCHCODE Varchar(4),
    '  IN pPMBRAND Varchar(50),
    '  IN pPMMODEL Varchar(50),
    '  IN pPMDESCRIPTION Varchar(50),
    '  IN pDATECREATED Date,
    '  IN pUPDATEDBY Int(11),
    '  IN pPMATRIXID Int(11)
    '  )
    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchQty_summary('" & sysBranchCode & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
  Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchQty_summary()")

End Sub

Public Sub sp_brnchIdentify_summary(pBranchCode As String, objRs As ADODB.Recordset)
On Error GoTo er

   'PROCEDURE sp_partsmatrix.sp_brnchIdentify_summary(
   'pBRANCHCODE VarChar(4)
   ')
   
    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchIdentify_summary('" & pBranchCode & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchIdentify_summary()")

End Sub

Public Sub sp_get_Branch(objRs As ADODB.Recordset)
On Error GoTo er

   'PROCEDURE sp_partsmatrix.sp_get_Branch(
   'pBRANCHCODE VarChar(4)
   ')
   
    DoEvents
    Qry = "Call sp_partsmatrix.sp_get_Branch('" & sysBranchCode & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_get_Branch()")

End Sub
