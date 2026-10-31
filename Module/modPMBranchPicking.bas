Attribute VB_Name = "modPMBranchPicking"
Option Explicit
Private Qry As String
Private rs As New ADODB.Recordset


Public Sub sp_brnchPicking_PickingListDtls(pREFNO As String, pBRANCHID As Integer, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_get_PartsmatrixID(
'  pBranchCode VarChar(4)
'  )

    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchPicking_PickingListDtls('" & sysBranchCode & "','" & pREFNO & "'," & pBRANCHID & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchPicking_PickingListDtls()")
End Sub

Public Sub sp_brnchPicking_PickingListTtl(pREFNO As String, pBRANCHID As Integer, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_get_PartsmatrixID(
'  pBranchCode VarChar(4)
'  )

    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchPicking_PickingListTtl('" & sysBranchCode & "','" & pREFNO & "'," & pBRANCHID & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchPicking_PickingListTtl()")
End Sub

Public Sub sp_get_Branch(objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_get_Branch(
'    pBranchCode VarChar(4)
'  )
   
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

Public Sub sp_brnchPicking_getRefNo(pSELBCODE As String, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_getRefNo(
'    pBranchCode VarChar(4)
'  )
   
    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchPicking_getRefNo('" & sysBranchCode & "', '" & pSELBCODE & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchPicking_getRefNo()")

End Sub


Public Sub sp_brnchPicking_PickingListSumm(objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_brnchPicking_PickingListSumm(
'    pBranchCode VarChar(4)
'  )
   
    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchPicking_PickingListSumm('" & sysBranchCode & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchPicking_PickingListSumm()")

End Sub

'******************** CODE UPDATED: 05/05/2017 ************************
Public Sub sp_get_BranchName(pBRANCHCODE As String, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_getBranchName(
'  pBranchCode VarChar(4)
'  )
    DoEvents
    Qry = "Call sp_partsmatrix.sp_get_BranchName('" & sysBranchCode & "','" & pBRANCHCODE & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_get_BranchName()")
'******************** CODE UPDATED: 05/05/2017 ************************
End Sub
