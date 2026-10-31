Attribute VB_Name = "modPMTransfer"
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

End Sub


Public Sub sp_getRefNo(pSELBCODE As String, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_getRefNo(
'    pBranchCode VarChar(4)
'  )
   
    DoEvents
    Qry = "Call sp_partsmatrix.sp_StockTransfer_getRefNo('" & sysBranchCode & "','" & pSELBCODE & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_getRefNo()")

End Sub

Public Sub sp_StockTransfer_getDetails(pSELBCODE As String, pREFNO As String, pPMATRIXID As Integer, pSTFREFNO As String, pBOOLEAN As String, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_getRefNo(
'    pBranchCode VarChar(4)
'  )
   
    DoEvents
    Qry = "Call sp_partsmatrix.sp_StockTransfer_getDetails('" & sysBranchCode & "','" & pSELBCODE & "', '" & pREFNO & "', " & pPMATRIXID & ", '" & pSTFREFNO & "', '" & pBOOLEAN & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_StockTransfer_getDetails()")

End Sub

Public Sub sp_StockTransfer_getDetailsTtl(pSELBCODE As String, pREFNO As String, pPMATRIXID As Integer, pBOOLEAN As String, pSTFREFNO As String, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_getRefNo(
'    pBranchCode VarChar(4)
'  )
   
    DoEvents
    Qry = "Call sp_partsmatrix.sp_StockTransfer_getDetailsTtl('" & sysBranchCode & "', '" & pSELBCODE & "', '" & pREFNO & "', " & pPMATRIXID & ", '" & pBOOLEAN & "', '" & pSTFREFNO & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_StockTransfer_getDetailsTtl()")

End Sub

'********************** CODE UPDATED: 04/06/2017 *********************
Public Sub sp_StockTransfer_create(pSTFDATE As String, pTRANSTO As String, pADDRESS As String, pREFNO As String, pPREPBY As String, pAPPBY As String, pUPDATEDBY As Integer, pMODELBRAND As String, pPMATRIXID As Integer, pCHKFINAL As Integer, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_getRefNo(
'    pBranchCode VarChar(4)
'  )
   
    DoEvents
    Qry = "Call sp_partsmatrix.sp_StockTransfer_create('" & sysBranchCode & "', '" & pSTFDATE & "', '" & pTRANSTO & "', '" & pADDRESS & "', '" & pREFNO & "', '" & pPREPBY & "', '" & pAPPBY & "', " & pUPDATEDBY & ", '" & pMODELBRAND & "', " & pPMATRIXID & ", " & pCHKFINAL & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_StockTransfer_create()")

End Sub
'****************** END CODE UPDATED: 04/06/2017 *********************


Public Sub sp_StockTransfer_FinalTransfer(pDRIVER As String, pCHECKER As String, pREFNO As String, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_getRefNo(
'    pBranchCode VarChar(4)
'  )
   
    DoEvents
    Qry = "Call sp_partsmatrix.sp_StockTransfer_FinalTransfer('" & sysBranchCode & "', '" & pDRIVER & "', '" & pCHECKER & "', '" & pREFNO & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_StockTransfer_FinalTransfer()")

End Sub


Public Sub sp_StockTransfer_summary(pDATEFROM As String, pDATETO As String, pDCODE As Integer, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_getRefNo(
'    pBranchCode VarChar(4)
'  )
   
    DoEvents
    Qry = "Call sp_partsmatrix.sp_StockTransfer_summary('" & sysBranchCode & "', '" & pDATEFROM & "', '" & pDATETO & "', " & pDCODE & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_StockTransfer_summary()")

End Sub

'********************** CODE UPDATED: 04/06/2017 *********************
Public Sub sp_StockTransfer_update(pSTFDATE As String, pTRANSTO As String, pADDRESS As String, pREFNO As String, pPREPBY As String, pAPPBY As String, pUPDATEDBY As Integer, pTRANSFERID As Integer, pCHECKFINAL As Integer, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_getRefNo(
'    pBranchCode VarChar(4)
'  )
   
    DoEvents
    Qry = "Call sp_partsmatrix.sp_StockTransfer_update('" & sysBranchCode & "', '" & pSTFDATE & "', '" & pTRANSTO & "', '" & pADDRESS & "', '" & pREFNO & "', '" & pPREPBY & "', '" & pAPPBY & "', " & pUPDATEDBY & ", " & pTRANSFERID & ", " & pCHECKFINAL & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_StockTransfer_update()")

End Sub
'****************** END CODE UPDATED: 04/06/2017 *********************
