Attribute VB_Name = "modPMDataEntry"
Option Explicit
Private Qry As String
Private rs As New ADODB.Recordset

Public Sub sp_get_PartsmatrixID(objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_get_PartsmatrixID(
'  pBranchCode VarChar(4)
'  )

    DoEvents
    Qry = "Call sp_partsmatrix.sp_get_PartsmatrixID('" & sysBranchCode & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_get_PartsmatrixID()")
End Sub


Public Sub sp_get_Partnumber(objRs As ADODB.Recordset)
On Error GoTo er
    
'PROCEDURE sp_partsmatrix.sp_get_Partnumber(
'   pBranchCode VarChar(4)
'  )

    DoEvents
    Qry = "Call sp_partsmatrix.sp_get_Partnumber('" & sysBranchCode & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:

   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_get_Partnumber()")

End Sub

Public Sub sp_dataentry_create(pPMBRAND As String, pPMODEL As String, pPMDESCRIPTION As String, pDATECREATED As String, pUPDATEDBY As Integer, pENCODEDBY As Integer, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_dataentry_create(
'   pBranchCode VarChar(4)
'  ,pPMBRAND               Varchar(50)
'  ,pPMMODEL               Varchar(50)
'  ,pPMDESCRIPTION          Varchar(50)
'  ,pDATECREATED            Date
'  ,pUPDATEDBY              Int(11)
'  ,pENCODEDBY              Int(11)
'  )

    DoEvents
    Qry = "Call sp_partsmatrix.sp_dataentry_create('" & sysBranchCode & "','" & pPMBRAND & "','" & pPMODEL & "', '" & pPMDESCRIPTION & "', '" & pDATECREATED & "', " & pUPDATEDBY & ", " & pENCODEDBY & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_dataentry_create()")

End Sub



Public Sub sp_dataentry_update(pPMBRAND As String, pPMMODEL As String, pPMDESCRIPTION As String, pDATECREATED As String, pUPDATEDBY As Integer, pPMATRIXID As Integer, objRs As ADODB.Recordset)
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
    Qry = "Call sp_partsmatrix.sp_dataentry_update('" & sysBranchCode & "', '" & pPMBRAND & "', '" & pPMMODEL & "', '" & pPMDESCRIPTION & "', '" & pDATECREATED & "', " & pUPDATEDBY & "," & pPMATRIXID & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_dataentry_update()")

End Sub


Public Sub sp_dataentry_getDetails(pPMATRIXID As Integer, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_dataentry_getDetails(
'  IN pBRANCHCODE Varchar(4),
'  IN pPMATRIXID Int(11)
'  )
    
    DoEvents
    Qry = "Call sp_partsmatrix.sp_dataentry_getDetails('" & sysBranchCode & "', " & pPMATRIXID & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_dataentry_getDetails()")

End Sub

Public Sub sp_dataentry_getDetails_total(pPMATRIXID As Integer, objRs As ADODB.Recordset)
On Error GoTo er

'PROCEDURE sp_partsmatrix.sp_dataentry_getDetails_total(
'  IN pBRANCHCODE Varchar(4),
'  IN pPMATRIXID Int(11)
'  )
    
    DoEvents
    Qry = "Call sp_partsmatrix.sp_dataentry_getDetails_total('" & sysBranchCode & "'," & pPMATRIXID & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_dataentry_getDetails_total()")

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


