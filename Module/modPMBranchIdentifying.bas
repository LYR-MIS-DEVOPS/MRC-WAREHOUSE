Attribute VB_Name = "modPMBranchIdentifying"
Option Explicit
Private Qry As String
Private rs As New ADODB.Recordset

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



Public Sub sp_brnchIdentify_getDetails(pBRANCHCODE As String, pPMATRIXID As Integer, pBOOLEAN As String, pREFID As Integer, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_brnchIdentify_getDetails(
'  pBranchCode VarChar(4)
'  ,pPMATRIXID Int(11)
'  ,pBOOLEAN Varchar(6)
'  )
   
    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchIdentify_getDetails('" & sysBranchCode & "','" & pBRANCHCODE & "', " & pPMATRIXID & ", '" & pBOOLEAN & "', " & pREFID & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchIdentify_getDetails()")

End Sub




Public Sub sp_brnchIdentify_getDetails_total(pSELBCODE As String, pPMATRIXID As Integer, pBOOLEAN As String, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_brnchIdentify_getDetails_total(
'   pBranchCode VarChar(4)
'  ,pPMATRIXID Int(11)
'  )
   
    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchIdentify_getDetails_total('" & sysBranchCode & "','" & pSELBCODE & "', " & pPMATRIXID & ", '" & pBOOLEAN & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchIdentify_getDetails_total()")

End Sub

Public Sub sp_brnchIdentify_create(pBRANCHCODE As String, pDATEPREPARED As String, pDATEFINAL As String, pPROCESSEDBY As Integer, pUPDATEDBY As Integer, pCHKFINAL As Integer, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_brnchIdentify_create(
'   pBranchCode VarChar(4)
'  ,pREFNO Varchar(20)
'  ,pDATEPREPARED DATE
'  ,pDATEFINAL DATE
'  ,pPROCESSEDBY Varchar(30)
'  )
   
    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchIdentify_create('" & sysBranchCode & "','" & pBRANCHCODE & "', '" & pDATEPREPARED & "','" & pDATEFINAL & "','" & pPROCESSEDBY & "','" & pUPDATEDBY & "'," & pCHKFINAL & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchIdentify_create()")

End Sub



Public Sub sp_brnchIdentify_update(pSELBCODE As String, pREFID As Integer, pREFNO As String, pDATEPREPARED As String, pDATEFINAL As String, pUPDATEDBY As Integer, pCHKFINAL As Integer, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_brnchIdentify_update(
'    pBranchCode VarChar(4)
'  ,pREFID int(11)
'  ,pREFNO Varchar(20)
'  ,pDATEPREPARED DATE
'  ,pDATEFINAL DATE
'  )
   
    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchIdentify_update('" & sysBranchCode & "', '" & pSELBCODE & "', " & pREFID & ", '" & pREFNO & "', '" & pDATEPREPARED & "','" & pDATEFINAL & "'," & pUPDATEDBY & "," & pCHKFINAL & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchIdentify_update()")

End Sub

Public Sub get_Model(pBRAND As String, pSELBCODE As String, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_get_Model(
'  pBranchCode VarChar(4)
'  ,pBRAND  Varchar(30)
'  )

    DoEvents
    Qry = "call sp_partsmatrix.sp_brnchIdentfy_getModel('" & sysBranchCode & "','" & pSELBCODE & "', '" & pBRAND & "')"
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

'PROCEDURE sp_partsmatrix.sp_get_Brand(
'      pBranchCode VarChar(4)
'  )

    DoEvents
    Qry = "call sp_partsmatrix.sp_get_Brand('" & pBRANCHCODE & "')"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->get_Brand_Model()")
End Sub


Public Sub sp_brnchIdentify_summary(pDATEFROM As String, pDATETO As String, pDCODE As Integer, objRs As ADODB.Recordset)
On Error GoTo er

   'PROCEDURE sp_partsmatrix.sp_brnchIdentify_summary(
   'pBRANCHCODE VarChar(4)
   ')
   
    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchIdentify_summary('" & sysBranchCode & "','" & pDATEFROM & "','" & pDATETO & "','" & pDCODE & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchIdentify_summary()")

End Sub

Public Sub sp_get_BranchSumm(objRs As ADODB.Recordset)
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

Public Sub sp_brnchIdentify_getRefNo(pSELBCODE As String, pPMATRIXID As Integer, objRs As ADODB.Recordset)
On Error GoTo er

   'PROCEDURE sp_partsmatrix.sp_get_Branch(
   'pBRANCHCODE VarChar(4)
   ')
   
    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchIdentify_getRefNo('" & sysBranchCode & "', '" & pSELBCODE & "', " & pPMATRIXID & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchIdentify_getRefNo()")

End Sub

Public Sub sp_brnchIdentify_dsplayHeader(pREFID As Integer, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_brnchIdentify_dsplayHeader(
'  pBRANCHCODE VarChar(5)
'  ,pREFID Int(11)
'  )
   
    DoEvents
    Qry = "Call sp_partsmatrix.sp_brnchIdentify_dsplayHeader('" & sysBranchCode & "', " & pREFID & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_brnchIdentify_dsplayHeader()")

End Sub
