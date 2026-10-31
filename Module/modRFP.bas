Attribute VB_Name = "modRFP"
Option Explicit
Private Qry As String
Private rs As New ADODB.Recordset

Public Sub sp_load_acct_title(pRFPNO As String, objRs As ADODB.Recordset)
On Error GoTo er
    
    DoEvents
    Qry = "SELECT ch.TransNAme,vou.code,vou.Description,vou.Type,vou.Amount FROM voucherdtls as vou " & _
          "INNER JOIN chartacct as ch ON vou.TransID = ch.TransID WHERE vou.RFPNO = '" & pRFPNO & "' GROUP by vou.SeqNO Order by vou.SeqNo asc;"
    'InputBox Qry, Qry, Qry
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_load_acct_title()")
End Sub

Public Sub sp_load_rfp_info(pRFPID As Long, pPOTYPE As String, objRs As ADODB.Recordset)
On Error GoTo er
    
    DoEvents
    Qry = "Call sp_lyrpurchase.sp_get_rfp_gssd(" & pRFPID & ",'" & pPOTYPE & "');"
    'InputBox Qry, Qry, Qry
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_load_rfp_info()")
End Sub

Public Sub sp_load_rfpother_info(pRFPID As Long, objRs As ADODB.Recordset)
On Error GoTo er
    
    DoEvents
    Qry = "Call sp_lyrpurchase.sp_get_rfp_other(" & pRFPID & ");"
    'InputBox Qry, Qry, Qry
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_load_rfpother_info()")
End Sub
Public Sub sp_load_rfp_deduction(pRFPID As Integer, objRs As ADODB.Recordset)
On Error GoTo er
    
    DoEvents
    Qry = "SELECT DTYPE, DESCRIPTION, REBDUCT, TYPE FROM lyrpurchase.deduction where RFID = " & pRFPID & ";"
    'InputBox Qry, Qry, Qry
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_load_rfp_deduction()")
End Sub


Public Sub sp_rfp_other_update(pRFPID As Long, pRFPNO As String, pRFPDATE As String, pDUEDATE As String, pSTATUTORY As String, pSUPPLIERID As Long, pPREPBY As String, pCHCKEDBY As String _
                                , pAPPRVDBY As String, pCHECKNO As String, pCHECKDATE As String _
                                , pREMARKS As String, pSTATUS As String, pTRANSNAME As String, ByRef pLevel As String, ByRef pMessage As String)
           
On Error GoTo er
DoEvents

    'PROCEDURE sp_lyrpurchase.sp_rfp_other_update(
    '    pRFPID          INT,
    '    pRFPNO          varchar(60),
    '    pRFPDATE        VARCHAR(60),
    '    pDUEDATE        VARCHAR(60),
    '    pSTATUTORY      VARCHAR(10),
    '    pSUPPLIERID     INT,
    '    pPREPBY         VARCHAR(60),
    '    pCHCKEDBY       varchar(60),
    '    pAPPRVDBY       varchar(60),
    '    pCHECKNO        varchar(60),
    '    pCHECKDATE      varchar(60),
    '    pREMARKS        varchar(500),
    '    pSTATUS         varchar(45),
    '    pTRANSNAME      varchar(100),
    '    pUPDATEDBY      int
    '    )
  
    Qry = "Call sp_lyrpurchase.sp_rfp_other_update(" & pRFPID & ",'" & pRFPNO & "','" & Format(pRFPDATE, "YYYY-mm-dd") & "','" & Format(pDUEDATE, "YYYY-mm-dd") & "','" & pSTATUTORY & "'," & pSUPPLIERID & ",'" & pPREPBY & _
                                                        "','" & pCHCKEDBY & "','" & pAPPRVDBY & "','" & Trim(pCHECKNO) & "','" & pCHECKDATE & _
                                                        "','" & pREMARKS & "','" & pSTATUS & "',  '" & pTRANSNAME & "'," & UserID & " );"
                
    'InputBox Qry, Qry, Qry
    Set rs = New ADODB.Recordset
    If rs.State <> 0 Then rs.Close
    rs.Open Qry, LYRcon, adOpenStatic
    
    If rs.State = 0 Then
        pLevel = "Error"
        pMessage = "Missing Output parameters!"
    ElseIf rs.RecordCount > 0 Then
        pLevel = rs!level
        pMessage = rs!message
    End If
    If rs.State <> 0 Then rs.Close
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_rfp_other_update()")
End Sub



Public Sub sp_rfp_other_create(pRFPNO As String, pRFPDATE As String, pDUEDATE As String, pSTATUTORY As String, pSUPPLIERID As Long, pPREPBY As String, pCHCKEDBY As String _
                                , pAPPRVDBY As String, pCHECKNO As String, pCHECKDATE As String _
                                , pREMARKS As String, pSTATUS As String, pTRANSNAME As String, ByRef pLevel As String, ByRef pMessage As String)
           
On Error GoTo er
DoEvents

    'PROCEDURE sp_lyrpurchase.sp_rfp_other_create(
    '  pRFPNO          varchar(60),
    '  pRFPDATE        VARCHAR(60),
    '  pDUEDATE        VARCHAR(60),
    '  pSTATUTORY      VARCHAR(10),
    '  pSUPPLIERID     INT,
    '  pPREPBY         VARCHAR(60),
    '  pCHCKEDBY       varchar(60),
    '  pAPPRVDBY       varchar(60),
    '  pCHECKNO        varchar(60),
    '  pCHECKDATE      varchar(60),
    '  pREMARKS        varchar(500),
    '  pSTATUS         varchar(45),
    '  pTRANSNAME      varchar(100),
    '  pENCODER        INT
    '  )
  
    Qry = "Call sp_lyrpurchase.sp_rfp_other_create('" & pRFPNO & "','" & Format(pRFPDATE, "YYYY-mm-dd") & "','" & Format(pDUEDATE, "YYYY-mm-dd") & "','" & pSTATUTORY & "'," & pSUPPLIERID & ",'" & pPREPBY & _
                                                        "','" & pCHCKEDBY & "','" & pAPPRVDBY & "','" & Trim(pCHECKNO) & "','" & pCHECKDATE & _
                                                        "','" & pREMARKS & "','" & pSTATUS & "', '" & pTRANSNAME & "'," & UserID & "  );"
                
    'InputBox Qry, Qry, Qry
    Set rs = New ADODB.Recordset
    If rs.State <> 0 Then rs.Close
    rs.Open Qry, LYRcon, adOpenStatic
    
    If rs.State = 0 Then
        pLevel = "Error"
        pMessage = "Missing Output parameters!"
        frmRFPOther.RFPID = 0
    ElseIf rs.RecordCount > 0 Then
        pLevel = rs!level
        pMessage = rs!message
        frmRFPOther.RFPID = rs!RFPID
    End If
    If rs.State <> 0 Then rs.Close
    
Exit Sub
er:
    DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_rfp_other_create()")
End Sub


Public Sub sp_LoadReconciliation_dtls(pPURCHASEID As Integer, pSHIPTO As String, objRs As ADODB.Recordset)
On Error GoTo er
    
    'PROCEDURE sp_lyrpurchase.sp_LoadReconciliation_dtls(
    '  pPURCHASEID INT,
    '  pSHIPPEDTO VarChar(20)
    '  )
    
    DoEvents
    Qry = "Call sp_lyrpurchase.sp_LoadReconciliation_dtls(" & pPURCHASEID & ",'" & pSHIPTO & "');"
    'InputBox Qry, Qry, Qry
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_LoadReconciliation_dtls()")
End Sub

Public Sub sp_LoadReconciliation(pDATEFROM As String, pDATETO As String, pSTATUS As String, pTERMS As String, objRs As ADODB.Recordset)
On Error GoTo er
    
    'PROCEDURE sp_lyrpurchase.sp_LoadReconciliation(
    '  pDATEFROM VARCHAR(60),
    '  pDATETO   VARCHAR(60),
    '  pSTATUS VarChar(20)
    '  )
    
    DoEvents
    Qry = "Call sp_lyrpurchase.sp_LoadReconciliation('" & pDATEFROM & "','" & pDATETO & "','" & pSTATUS & "','" & pTERMS & "');"
    'InputBox Qry, Qry, Qry
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_LoadReconciliation()")
End Sub

Public Sub sp_supplier_list(objRs As ADODB.Recordset)
On Error GoTo er
    
    DoEvents
    Qry = "SELECT SUPPLIERID,SUPPLIERNAME,IFNULL(ADDRESS,''),IFNULL(STATUTORY,'') FROM supplier;"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_supplier_list()")
End Sub



Public Sub sp_business_unit_list(objRs As ADODB.Recordset)
On Error GoTo er
    
    DoEvents
    Qry = "SELECT pbusinessunit.Name From pbusinessunit"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_business_unit_list()")
End Sub

Public Sub sp_company_list(objRs As ADODB.Recordset)
On Error GoTo er
    
    DoEvents
    Qry = "SELECT co_Name FROM company ORDER BY co_Name ASC;"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_company_list()")
End Sub

Public Sub sp_branch_list(objRs As ADODB.Recordset)
On Error GoTo er
    
    DoEvents
    Qry = "CALL sp_lyrpurchase.sp_BranchBillTo_list()"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_branch_list()")
End Sub

Public Sub sp_get_spcrr(pSPCRRNO As String, pSUPID As Integer, ByRef pLevel As String, ByRef pMessage As String, objRs As ADODB.Recordset)
On Error GoTo er


    'PROCEDURE sp_lyrpurchase.sp_LoadSPCRR(
    '  pSPCRRNO VarChar(20)
    '  )
    
    DoEvents
    Qry = "Call sp_lyrpurchase.sp_LoadSPCRR('" & pSPCRRNO & "'," & pSUPID & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
    'InputBox Qry, Qry, Qry
    If objRs.State = 0 Then
        pLevel = "Error"
        pMessage = "Missing Output parameters!"
    ElseIf objRs.RecordCount > 0 Then
        pLevel = objRs!level
        pMessage = objRs!message
    ElseIf objRs.RecordCount = 0 Then
        pLevel = "Warning"
        pMessage = "No record found!"
    End If
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_get_spcrr()")
End Sub

Public Sub sp_load_RFPMOTOR(pSUPID As Integer, pDATEFROM As String, pDATETO As String, pBILLTO As String, objRs As ADODB.Recordset)
On Error GoTo er
    
    
'    PROCEDURE lyrpurchase.sp_LoadRFPMOTOR(
'      IN pSUPID INT,
'      IN pDfrom DATE,
'      IN pDto DATE,
'      IN pBILTO VARCHAR(50)
'      )
  
    DoEvents
    Qry = "CALL lyrpurchase.sp_LoadRFPMOTOR( " & pSUPID & ",'" & pDATEFROM & "','" & pDATETO & "','" & pBILLTO & "')"
    'InputBox Qry, Qry, Qry
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_load_RFPMOTOR()")
End Sub

Public Sub sp_load_spcrr_old(pSUPID As Integer, pDATEFROM As String, pDATETO As String, pBILLTO As String, objRs As ADODB.Recordset)
On Error GoTo er
    
    'PROCEDURE lyrpurchase.sp_LoadRFP(
    '  IN pSUPID INT,
    '  IN pDfrom DATE,
    '  IN pDto DATE,
    '  IN pBILLTO VARCHAR(50)
    '  )
        
    DoEvents
    Qry = "Call lyrpurchase.sp_LoadRFP( " & pSUPID & ",'" & pDATEFROM & "', '" & pDATETO & "','" & pBILLTO & "');"
    'InputBox Qry, Qry, Qry
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_load_spcrr_old()")
End Sub


Public Sub sp_load_spcrr(pBILLTO As String, pSUPID As Integer, pDATEFROM As String, pDATETO As String, objRs As ADODB.Recordset)
On Error GoTo er
    
    'PROCEDURE sp_lyrpurchase.sp_LoadSPCRR_by_date(
    '  pBILLTO       varchar(60),
    '  pSUPID        INT,
    '  pDATEFROM     varchar(60),
    '  pDATETO VarChar(60)
    '  )
        
    DoEvents
    Qry = "Call sp_lyrpurchase.sp_LoadSPCRR_by_date('" & pBILLTO & "', " & pSUPID & ",'" & pDATEFROM & "', '" & pDATETO & "');"
    'InputBox Qry, Qry, Qry
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_load_spcrr()")
End Sub

Public Sub sp_AccountTitles_List(objRs As ADODB.Recordset)
On Error GoTo er
    
    DoEvents
    Qry = "SELECT AcctCode,AcctTitles,`type` FROM chartacct " & _
          "GROUP BY AcctCode ORDER BY AcctCode ASC;"
    'InputBox Qry, Qry, Qry
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_AccountTitles_List()")
End Sub

Public Sub sp_get_po_info(pPONO As String, pSUPPLIERID As Long, pPOTYPE As String, objRs As ADODB.Recordset)
On Error GoTo er

    DoEvents
    Qry = "call sp_lyrpurchase.sp_getgssd_PO_info('" & pPONO & "', " & pSUPPLIERID & ", '" & pPOTYPE & "')"
    
    'InputBox Qry, Qry, Qry
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_get_po_info()")
End Sub

Public Sub sp_get_AccountTitles_Auto(pTransID As Long, objRs As ADODB.Recordset)
On Error GoTo er
    
    DoEvents
    Qry = "SELECT AcctCode,AcctTitles,`Type` FROM chartacct  where transid = " & pTransID & " and ViewOption = 'AUTO' ORDER BY seqno ASC;"
    'InputBox Qry, Qry, Qry
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_get_AccountTitles_Auto()")
End Sub

Public Sub sp_sync_acct_titles(ByRef pLevel As String, ByRef pMessage As String)
On Error GoTo er
    
    DoEvents
    Qry = "Call sp_lyrpurchase.sp_sync_acct_titles(" & UserID & ");"
    
    'InputBox Qry, Qry, Qrys
    Set rs = New ADODB.Recordset
    If rs.State <> 0 Then rs.Close
    rs.Open Qry, LYRcon, adOpenStatic
    
    If rs.State = 0 Then
        pLevel = "Error"
        pMessage = "Missing Output parameters!"
    ElseIf rs.RecordCount > 0 Then
        pLevel = rs!level
        pMessage = rs!message
    End If
    If rs.State <> 0 Then rs.Close
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_sync_acct_titles()")
End Sub

Public Sub sp_transaction_list(objRs As ADODB.Recordset)
On Error GoTo er
    
    DoEvents
    Qry = "SELECT DISTINCT TransID,TransName FROM chartacct ORDER BY TransName ASC;"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_transaction_list()")
End Sub
