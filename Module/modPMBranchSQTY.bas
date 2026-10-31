Attribute VB_Name = "modPMBranchSQTY"
Option Explicit
Private Qry As String
Private rs As New ADODB.Recordset

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

Public Sub get_Model(pBRAND As String, pSELBCODE As String, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_get_Model(
'  pBranchCode VarChar(4)
'  ,pBRAND  Varchar(30)
'  )

    DoEvents
    Qry = "call sp_partsmatrix.sp_WHbrnchQty_getModel('" & sysBranchCode & "','" & pSELBCODE & "', '" & pBRAND & "')"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_get_Model()")
End Sub


Public Sub get_Brand(pBRANCHCODE As String, objRs As ADODB.Recordset)
On Error GoTo er

'    PROCEDURE sp_partsmatrix.sp_get_Brand(
'      pBRANCHCODE VarChar(4)
'  )

    DoEvents
   ' Qry = "Select Distinct(PMBRAND) from mrc2retail.partsmatrix INNER JOIN mrc2retail.partsmatrix_dtls ON partsmatrix.PMATRIXID = partsmatrix_dtls.PMATRIXID group by partsmatrix_dtls.PMATRIXID"
          
    Qry = "call sp_partsmatrix.sp_get_Brand('" & pBRANCHCODE & "')"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->get_Brand")
End Sub


Public Sub sp_brnchQty_update(pPMATRIXID As Integer, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_brnchQty_update(
'  pBranchCode VarChar(4)
'  ,pPMATRIXID int(11)
'  )

    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchQty_update('" & sysBranchCode & "', " & pPMATRIXID & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchQty_update()")

End Sub


Public Sub sp_brnchQty_getDetails(pSELBCODE As String, pPMATRIXID As Integer, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_brnchQty_getDetails(
'  pBranchCode VarChar(4)
'  , pPMATRIXID Int(11)
'  , pBOOLEAN Varchar(6)
'  )

    DoEvents
    Qry = "Call sp_partsmatrix.sp_WHbrnchQty_getDetails('" & sysBranchCode & "', '" & pSELBCODE & "', " & pPMATRIXID & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchQty_getDetails()")

End Sub


Public Sub sp_brnchQty_getDetails_total(pSELBCODE As String, pPMATRIXID As Integer, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_brnchQty_getDetails_total(
'    pBranchCode VarChar(4)
'  , pPMATRIXID Int(11)
'  )

    DoEvents
    Qry = "Call sp_partsmatrix.sp_WHbrnchQty_getDetails_total('" & pSELBCODE & "', " & pPMATRIXID & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchQty_getDetails_total()")

End Sub

'**************** CODE UPDATED: 04/19/17 ****************************
'Added pTmpPARAM

Public Sub sp_brnchQty_summary(pTmpPARAM As Integer, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_brnchQty_summary(
'  pBranchCode VarChar(4)
'  )

    DoEvents
    Qry = "Call sp_partsmatrix.sp_WHbrnchQty_summary('" & sysBranchCode & "'," & pTmpPARAM & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
  Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchQty_summary()")

End Sub
'************ END CODE UPDATED: 04/19/17 ****************************

'**************** CODE UPDATED: 04/19/17 ****************************
Public Sub sp_brnchQty_summary_wBRAND(pSELBCODE As String, pBRAND As String, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_brnchQty_summary(
'  pBranchCode VarChar(4)
'  )

    DoEvents
    Qry = "Call sp_partsmatrix.sp_WHbrnchQty_summary_wBRAND('" & sysBranchCode & "', '" & pSELBCODE & "', '" & pBRAND & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
  Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchQty_summary_wBRAND()")

End Sub
'************ END CODE UPDATED: 04/19/17 ****************************


'**************** CODE UPDATED: 04/21/17 ****************************
Public Sub sp_WHbrnchQty_getBranch(objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_WHbrnchQty_getBranch(
'  pBranchCode VarChar(4)
'  )

    DoEvents
    Qry = "Call sp_partsmatrix.sp_WHbrnchQty_getBranch('" & sysBranchCode & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
  Call modObjTools.ErrorMessage(Err.Description, "Error->sp_WHbrnchQty_getBranch()")

End Sub
'************ END CODE UPDATED: 04/21/17 ****************************

Public Sub sp_brnchQty_create(pPMATRIXID As Integer, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_brnchQty_update(
'  pBranchCode VarChar(4)
'  ,pPMATRIXID int(11)
'  )

    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchQty_create('" & sysBranchCode & "', " & pPMATRIXID & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchQty_create()")

End Sub
