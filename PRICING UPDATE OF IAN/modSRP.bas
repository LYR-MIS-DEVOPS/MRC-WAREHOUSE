Attribute VB_Name = "modSRP"
Option Explicit
Private Qry As String
Private rs As New ADODB.Recordset

Public Sub sp_get_SMSQTY_LDI(pSPAREPARTID As Integer, objRs As ADODB.Recordset)
On Error GoTo er

    'PROCEDURE sp_srppricing.sp_get_SMSQTY_LDI(pSPAREPARTID INT)
    
    DoEvents
    Qry = "call sp_srppricing.sp_get_SMSQTY_LDI(" & pSPAREPARTID & ");"
    
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_get_SMSQTY_LDI()")
End Sub

Public Sub sp_show_approved_branchSRP(pBREFNO As String, objRs As ADODB.Recordset)
On Error GoTo er

    'PROCEDURE sp_srppricing.sp_show_approved_branchSRP(
    '   pBREFNO varchar(20))
    
    DoEvents
    Qry = "call sp_srppricing.sp_show_approved_branchSRP('" & pBREFNO & "');"
    
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_show_approved_branchSRP()")
End Sub

Public Sub sp_branchsrp_approve(pBREFNO As String, pDATEENCODE As String, pREMARKS As String, pENCODEDBY As String, pRECOMBY As String, pNOTEDBY As String, pAPPROVEDBY As String, pBRANCHCODE As String, ByRef pLevel As String, ByRef pMessage As String)
           
On Error GoTo er

    DoEvents
    'PROCEDURE sp_srppricing.sp_branchSRP_approve(
    '    pBREFNO VarChar(20)
    '  , pDATEENCODE   Varchar(60)
    '  , pDATEAPPROVE  Varchar(60)
    '  , pREMARKS      Varchar(200)
    '  , pENCODEDBY    Varchar(60)
    '  , pRECOMBY      Varchar(60)
    '  , pNOTEDBY      Varchar(60)
    '  , pAPPROVEDBY   Varchar(60)
    '  , pBranchCode   Varchar(5)
    '  , pSUPPLIER     LongText
    '  , pENCODER      Int
    '  )

    Qry = "Call sp_srppricing.sp_branchSRP_approve('" & pBREFNO & "','" & pDATEENCODE & "','" & pREMARKS & _
          "','" & pENCODEDBY & "','" & pRECOMBY & "','" & pNOTEDBY & "','" & pAPPROVEDBY & _
          "','" & pBRANCHCODE & "'," & UserID & ")"
                
    'InputBox Qry, Qry, Qry
    Set rs = New ADODB.Recordset
    If rs.State <> 0 Then rs.Close
    rs.Open Qry, LYRcon, adOpenStatic
    
    If rs.State = 0 Then
        pLevel = "Error"
        pMessage = "Missing Output parameters!"
    ElseIf rs.RecordCount > 0 Then
        pLevel = rs!Level
        pMessage = rs!message
    End If
    If rs.State <> 0 Then rs.Close
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_branchsrp_approve()")
End Sub


Public Sub sp_show_branchSRP(pBRANCHCODE As String, pBREFNO As String, objRs As ADODB.Recordset)
On Error GoTo er

    'PROCEDURE sp_srppricing.sp_show_branchSRP(
    '   pBRANCHCODE varchar(4),
    '   pBREFNO varchar(20))
    
    DoEvents
    Qry = "call sp_srppricing.sp_show_branchSRP('" & pBRANCHCODE & "','" & pBREFNO & "');"
    
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_show_branchSRP()")
End Sub


Public Sub sp_branchSRPRefNo_list(objRs As ADODB.Recordset)
On Error GoTo er

    DoEvents
    Qry = "call sp_srppricing.sp_branchSRPRefNo_list();"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic

Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_branchSRPRefNo_list()")
End Sub

Public Sub sp_partno_info(pCODE As String, objRs As ADODB.Recordset)
On Error GoTo er
    
    DoEvents
    Qry = "SELECT * FROM mrcwarehouse.supplier where suppliercode = '" & pCODE & "';"

    'InputBox Qry, Qry, Qry
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_partno_info()")
End Sub

Public Sub sp_partno_list(objRs As ADODB.Recordset)
On Error GoTo er

    DoEvents
    Qry = "SELECT SPAREPARTID,PARTNUMBER FROM mrcwarehouse.spareparts;"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic

Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_partno_list()")
End Sub


Public Sub sp_class_list(objRs As ADODB.Recordset)
On Error GoTo er

    DoEvents
    Qry = "SELECT DISTINCT(remarks) FROM mrcwarehouse.spareparts ORDER BY remarks ASC;"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic

Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_class_list()")
End Sub

Public Sub sp_supplier_list(objRs As ADODB.Recordset)
On Error GoTo er
    
    DoEvents
    Qry = "SELECT SUPPLIERCODE, SUPPLIERNAME FROM mrcwarehouse.supplier ORDER BY SUPPLIERNAME ASC;"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_supplier_list()")
End Sub

Public Sub sp_ws_history(pSPID As Integer, objRs As ADODB.Recordset)
On Error GoTo er
    
    'sp_srppricing.sp_load_ws_history(pSPID Int(11))
    DoEvents
    Qry = "call sp_srppricing.sp_load_ws_history(" & pSPID & ")"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_ws_history()")
End Sub

Public Sub sp_srp_history(pSPID As Integer, objRs As ADODB.Recordset)
On Error GoTo er
    
    'sp_srppricing.sp_load_srp_history(pSPID Int(11))
    DoEvents
    Qry = "call sp_srppricing.sp_load_srp_history(" & pSPID & ")"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_srp_history()")
End Sub

Public Sub sp_load_record(pSUPPLIERCODE As String, objRs As ADODB.Recordset)
On Error GoTo er

    'sp_srppricing.sp_load_record(pSUPPLIERCODE varchar(5))
    DoEvents
    Qry = "call sp_srppricing.sp_load_record('" & pSUPPLIERCODE & "')"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_load_record()")
End Sub

Public Sub sp_priceupdate_save(pSPAREPARTID As Integer, pREFERENCE As String, pCOST As String, pSRPPRICE As String, pSRPPROMAR As String, pSRPPRORATIO As String, pWSPRICE As String _
                             , pWSPROMAR As String, pWSPRORATIO As String, pMAXDISCOUNT As Integer, ByRef pLevel As String, ByRef pMessage As String)
On Error GoTo er
    'PROCEDURE sp_srppricing.sp_priceupdate_save(
    '  pSPAREPARTID Int,
    '  pREFERENCE   Varchar(20),
    '  pCOST        Varchar(20),
    '  pSRPPRICE    Varchar(20),
    '  pSRPPROMAR   Varchar(20),
    '  pSRPPRORATIO Varchar(20),
    '  pWSPRICE     Varchar(20),
    '  pWSPROMAR    Varchar(20),
    '  pWSPRORATIO  Varchar(20),
    '  pMAXDISCOUNT Int,
    '  pENCODER     Int
    '  )
    Qry = "Call sp_srppricing.sp_priceupdate_save(" & pSPAREPARTID & ",'" & pREFERENCE & "','" & pCOST & "','" & pSRPPRICE & "','" & pSRPPROMAR & "','" & pSRPPRORATIO & "','" & pWSPRICE & _
                                                  "','" & pWSPROMAR & "','" & pWSPRORATIO & "'," & pMAXDISCOUNT & "," & UserID & ");"
                
    'InputBox Qry, Qry, Qry
    Set rs = New ADODB.Recordset
    If rs.State <> 0 Then rs.Close
    rs.Open Qry, LYRcon, adOpenStatic
    
    If rs.State = 0 Then
        pLevel = "Error"
        pMessage = "Missing Output parameters!"
    ElseIf rs.RecordCount > 0 Then
        pLevel = rs!Level
        pMessage = rs!message
    End If
    If rs.State <> 0 Then rs.Close
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_priceupdate_save()")
End Sub

Public Sub sp_srppricing_update(pPACKINGNO As String, pDATEFINAL As String, pSTATUS As String, pREMARKS As String, pENCODEDBY As String, pAPPROVEDBY As String, ByRef pLevel As String, ByRef pMessage As String)
                                
On Error GoTo er

    DoEvents
    
    '    PROCEDURE sp_srppricing.sp_srppricing_update(
    '    pPACKINGNO    VarChar(20)
    '  , pDATEFINAL    varchar(60)
    '  , pSTATUS       varchar(10)
    '  , pREMARKS      varchar(200)
    '  , pENCODEDBY    varchar(60)
    '  , pENCODER      int
    '  , pAPPROVEDBY   varchar(60)
    '  )

    Qry = "Call sp_srppricing.sp_srppricing_update('" & pPACKINGNO & "','" & pDATEFINAL & "','" & pSTATUS & "','" & pREMARKS & "','" & pENCODEDBY & "'," & UserID & ",'" & pAPPROVEDBY & "');"
                
    'InputBox Qry, Qry, Qry
    Set rs = New ADODB.Recordset
    If rs.State <> 0 Then rs.Close
    rs.Open Qry, LYRcon, adOpenStatic
    
    If rs.State = 0 Then
        pLevel = "Error"
        pMessage = "Missing Output parameters!"
    ElseIf rs.RecordCount > 0 Then
        pLevel = rs!Level
        pMessage = rs!message
    End If
    If rs.State <> 0 Then rs.Close
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_srppricing_update()")
End Sub



Public Sub sp_srppricing_summary(pDATEFROM As String, pDATETO As String, objRs As ADODB.Recordset)
On Error GoTo er

    DoEvents
    
    'PROCEDURE sp_srppricing.sp_srppricing_summary(
    '   pDATEFROM VarChar(60)
    '  ,pDATETO        Varchar(60)
    '  )
    
    Qry = "Call sp_srppricing.sp_srppricing_summary('" & pDATEFROM & "','" & pDATETO & "');"
    'InputBox Qry, Qry, Qry
     Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
DoEvents
   Screen.MousePointer = vbDefault
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_srppricing_summary()")
End Sub

Public Sub sp_get_refno(objRs As ADODB.Recordset)
On Error GoTo er

    DoEvents
    Qry = "SELECT MAX(RefNo) + 1 as RefNo FROM mrcwarehouse.srppricing;"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_get_refno()")
End Sub

Public Sub sp_get_packing_info(pPACKINGNO As String, objRs As ADODB.Recordset, ByRef pLevel As String, ByRef pMessage As String)
On Error GoTo er
    
'    PROCEDURE sp_srppricing.sp_get_packing_info(
'    pPACKINGNO VarChar(20)
'    )
    
    DoEvents
    Qry = "Call sp_srppricing.sp_get_packing_info('" & pPACKINGNO & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
    If objRs.State = 0 Then
        pLevel = "Error"
        pMessage = "Missing Output parameters!"
    ElseIf objRs.RecordCount > 0 Then
        pLevel = objRs!Level
        pMessage = objRs!message
    End If
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_get_packing_info()")
End Sub

Public Sub sp_show_record(pPACKINGNO As String, pFILTERBY As String, objRs As ADODB.Recordset)
On Error GoTo er
    
'    PROCEDURE sp_srppricing.sp_get_show_record(
'    pPACKINGNO VarChar(20)
'    pFILTERBY  varchar(5)
'    )
    
    DoEvents
    Qry = "Call sp_srppricing.sp_show_record('" & pPACKINGNO & "','" & pFILTERBY & "');"
    'InputBox Qry, Qry, Qry
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_show_record()")
End Sub

Public Sub sp_srppricing_create(pPACKINGNO As String, pDATEFINAL As String, pSTATUS As String, pREMARKS As String, pENCODEDBY As String, pAPPROVEDBY As String, pENTRYTYPE As String, ByRef pLevel As String, ByRef pMessage As String)
           
On Error GoTo er

    DoEvents
    
    'PROCEDURE sp_srppricing.sp_srppricing_create(
    '   pPACKINGNO    VarChar(20)
    '  ,pDATEFINAL    varchar(60)
    '  ,pSTATUS       varchar(10)
    '  ,pREMARKS      varchar(200)
    '  ,pENCODEDBY    varchar(60)
    '  ,pAPPROVEDBY   varchar(60)
    '  ,pENTRYTYPE varchar(10)
    '  )

    Qry = "Call sp_srppricing.sp_srppricing_create('" & pPACKINGNO & "','" & pDATEFINAL & "','" & pSTATUS & "','" & pREMARKS & "','" & pENCODEDBY & "','" & pAPPROVEDBY & "','" & pENTRYTYPE & "');"
                
    'InputBox Qry, Qry, Qry
    Set rs = New ADODB.Recordset
    If rs.State <> 0 Then rs.Close
    rs.Open Qry, LYRcon, adOpenStatic
    
    If rs.State = 0 Then
        pLevel = "Error"
        pMessage = "Missing Output parameters!"
    ElseIf rs.RecordCount > 0 Then
        pLevel = rs!Level
        pMessage = rs!message
    End If
    If rs.State <> 0 Then rs.Close
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_srppricing_create()")
End Sub

