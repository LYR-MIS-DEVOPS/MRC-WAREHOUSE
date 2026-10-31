Attribute VB_Name = "modAPI"
Option Explicit
Private Qry As String
Private rs As New ADODB.Recordset

Public Sub sp_get_sparepart_info(pSPAREPART As String, pBRANCHCODE As String, objRs As ADODB.Recordset, ByRef pLevel As String, ByRef pMessage As String)
On Error GoTo er
    
    'PROCEDURE sp_acctpayinvcng.sp_get_sparepart_info(
    '  pSPAREPART  varchar(50),
    '  pBRANCHCODE VarChar(6)
    '  )

    DoEvents
    Qry = "Call sp_acctpayinvcng.sp_get_sparepart_info('" & pSPAREPART & "','" & pBRANCHCODE & "');"
    'InputBox Qry, Qry, Qry
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_get_sparepart_info()")
End Sub

Public Sub sp_checkedBy_list(objRs As ADODB.Recordset)
On Error GoTo er
    
    DoEvents
    Qry = "SELECT `VALUE` as CheckedBy FROM mktg3s.settings where remarks = 'RFP Checked By'"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_checkedBy_list()")
End Sub

Public Sub sp_supplier_list(objRs As ADODB.Recordset)
On Error GoTo er
    
    DoEvents
    Qry = "SELECT SUPPLIERID,SUPPLIERNAME FROM mktg3s.supplier where SUPPLIERNAME NOT LIKE '%INACTIVE%'"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_supplier_list()")
End Sub

Public Sub sp_3s_branch_list(objRs As ADODB.Recordset)
On Error GoTo er
    
    DoEvents
    Qry = "SELECT BranchCode,BranchName,BranchId FROM mktg3s.branch where branchcategory = '3S' ORDER BY BRANCHNAME ASC;"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_3s_branch_list()")
End Sub

Public Sub sp_acctpayinvcng_summary(pDATEFROM As String, pDATETO As String, pINVSTAT As String, objRs As ADODB.Recordset)
On Error GoTo er
    DoEvents
    
    'PROCEDURE sp_acctpayinvcng.sp_acctpayinvcng_summary(
    '   pDATEFROM    VarChar(60)
    '  ,pDATETO        Varchar(60)
    '  ,pINVSTAT        Varchar(10)
    '
    '  )
    
    Qry = "Call sp_acctpayinvcng.sp_acctpayinvcng_summary('" & pDATEFROM & "','" & pDATETO & "','" & pINVSTAT & "');"
    'InputBox Qry, Qry, Qry
     Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
DoEvents
   Screen.MousePointer = vbDefault
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_acctpayinvcng_summary()")
End Sub


Public Sub sp_get_spcrr_info(pSPCRRNO As Integer, objRs As ADODB.Recordset)
On Error GoTo er


    DoEvents
    Qry = "SELECT BRANCHNAME,SUPINVID,DATERCVD,PACKINGNO,PONO,WAYBILLNO,INVOICETTLPAY,INVSTATUS FROM mktg3s.supplierinvoice INNER JOIN mktg3s.BRANCH ON supplierinvoice.BRANCHCODE = BRANCH.BRANCHCODE WHERE SUPINVID = " & pSPCRRNO & ";"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_get_spcrr_info()")
End Sub


Public Sub sp_get_packing_info(pPACKINGNO As String, pBRANCHCODE As String, objRs As ADODB.Recordset, ByRef pLevel As String, ByRef pMessage As String)
On Error GoTo er
    
'    PROCEDURE sp_acctpayinvcng.sp_get_packing_info(
'    pPACKINGNO VarChar(20)
'    pBRANCHCODE    Varchar(6)
'    )

    DoEvents
    Qry = "Call sp_acctpayinvcng.sp_get_packing_info('" & pPACKINGNO & "','" & pBRANCHCODE & "');"
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

Public Sub sp_refresh_packing_info(pPACKINGNO As String, pBCode As String, objRs As ADODB.Recordset)
On Error GoTo er

    'PROCEDURE sp_acctpayinvcng.sp_refresh_packing_info(
    '  pPACKINGNO VarChar(20)
    '  )
    
    DoEvents
    Qry = "Call sp_acctpayinvcng.sp_refresh_packing_info('" & pPACKINGNO & "','" & pBCode & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_refresh_packing_info()")
End Sub

Public Sub sp_load_rfp_deduction(pRFPID As Integer, objRs As ADODB.Recordset)
On Error GoTo er
    
    DoEvents
    Qry = "SELECT * FROM mktg3s.rfpdeduct where rfpid = " & pRFPID & ";"
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

Public Sub sp_load_rfp_info(pRFPNO As String, objRs As ADODB.Recordset)
On Error GoTo er
    
    DoEvents
    Qry = "Call sp_acctpayinvcng.sp_load_rfp_info('" & pRFPNO & "');"
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

Public Sub sp_load_packing_info(pPACKINGNO As String, pBCode As String, objRs As ADODB.Recordset)
On Error GoTo er

    'PROCEDURE sp_acctpayinvcng.sp_load_packing_info(
    '  pPACKINGNO VarChar(20)
    '  )
    
    DoEvents
    Qry = "Call sp_acctpayinvcng.sp_load_packing_info('" & pPACKINGNO & "','" & pBCode & "');"
    'InputBox Qry, Qry, Qry
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_load_packing_info()")
End Sub

Public Sub sp_rfp_update(pRFPID As Integer, pRFPNO As String, prfpdate As String, psupplierid As Integer, pSOLDTO As String, pchekedby As String, pNOTEDBY As String _
                                , papprvdby As String, pcheckno As String, pchecknodate As String, pTotalAmount As String, pTotalDeduction As String, pTotalNetPay As String _
                                , ByRef pLevel As String, ByRef pMessage As String)
           
On Error GoTo er

    DoEvents

            
    Qry = "Call sp_acctpayinvcng.sp_rfp_update(" & pRFPID & ",'" & pRFPNO & "','" & Format(prfpdate, "YYYY-mm-dd") & "'," & psupplierid & ",'" & pSOLDTO & "','" & pchekedby & _
                                                        "','" & pNOTEDBY & "','" & papprvdby & "','" & Trim(pcheckno) & "','" & Format(pchecknodate, "YYYY-mm-dd") & _
                                                        "'," & UserID & ",'" & Replace(pTotalAmount, ",", "") & "','" & Val(Replace(pTotalDeduction, ",", "")) & _
                                                        "','" & Val(Replace(pTotalNetPay, ",", "")) & "');"
                
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_rfp_update()")
End Sub
Public Sub sp_rfp_create(pRFPNO As String, prfpdate As String, psupplierid As Integer, pSOLDTO As String, pchekedby As String, pNOTEDBY As String _
                                , papprvdby As String, pcheckno As String, pchecknodate As String, pTotalAmount As String, pTotalDeduction As String, pTotalNetPay As String _
                                , ByRef pLevel As String, ByRef pMessage As String)
           
On Error GoTo er

    DoEvents


    Qry = "Call sp_acctpayinvcng.sp_rfp_create('" & pRFPNO & "','" & Format(prfpdate, "YYYY-mm-dd") & "'," & psupplierid & ",'" & pSOLDTO & "','" & pchekedby & _
                                                        "','" & pNOTEDBY & "','" & papprvdby & "','" & Trim(pcheckno) & "','" & Format(pchecknodate, "YYYY-mm-dd") & _
                                                        "'," & UserID & ",'" & Replace(pTotalAmount, ",", "") & "','" & Val(Replace(pTotalDeduction, ",", "")) & _
                                                        "','" & Val(Replace(pTotalNetPay, ",", "")) & "');"
                
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_rfp_create()")
End Sub

Public Sub sp_acctpayinvcng_create(pSPRRNO As Integer, pPACKINGNO As String, pPACKINGDATE As String, pTERMS As String, pDUEDATE As String, pINVDISC1 As String _
                                , pDISC1VAL As String, pDISC1TYPE As String, pINVDISC2 As String, pDISC2VAL As String, pDISC2TYPE As String _
                                , pINVDISC3 As String, pDISC3VAL As String, pDISC3TYPE As String, pINVDISC4 As String, pDISC4VAL As String _
                                , pDISC4TYPE As String, pINVDISC5 As String, pDISC5VAL As String, pDISC5TYPE As String, pVAT As String, pVATVALUE As String _
                                , pPREPAREDBY As String, pAPPROVEDBY As String, pINVTYPE As String, pINVSTATUS As String, pINVNOTES As String _
                                , pINVLESSPER As String, pINVLESSPER2 As String, pINVLESSPER3 As String, pINVOICELESS As Double, pOTHERAMNT As Double _
                                , pINVOICETTLPAY As Double, pRECVDBY As String, pDATERCVD As String, pCHCKNO As String, pVOCHERNO As String _
                                , pCHCKDATE As String, pDISNOTE As String, pBRANCHCODE As String, pSUPPLIER As String _
                                , pSHIPPER As String, pPONO As String, pRFP As String, pWAYBILLNO As String, ByRef pLevel As String, ByRef pMessage As String)
           
On Error GoTo er

    DoEvents

    Qry = "Call sp_acctpayinvcng.sp_acctpayinvcng_create('" & pSPRRNO & "','" & pPACKINGNO & "','" & pPACKINGDATE & "','" & pTERMS & "','" & pDUEDATE & _
                                                        "','" & pINVDISC1 & "','" & pDISC1VAL & "','" & pDISC1TYPE & _
                                                        "','" & pINVDISC2 & "','" & pDISC2VAL & "','" & pDISC2TYPE & _
                                                        "','" & pINVDISC3 & "','" & pDISC3VAL & "','" & pDISC3TYPE & _
                                                        "','" & pINVDISC4 & "','" & pDISC4VAL & "','" & pDISC4TYPE & _
                                                        "','" & pINVDISC5 & "','" & pDISC5VAL & "','" & pDISC5TYPE & _
                                                        "','" & pVAT & "','" & pVATVALUE & "','" & pPREPAREDBY & "','" & pAPPROVEDBY & _
                                                        "','" & pINVTYPE & "','" & pINVSTATUS & "',' " & pINVNOTES & _
                                                        "','" & pINVLESSPER & "','" & pINVLESSPER2 & "','" & pINVLESSPER3 & "','" & pINVOICELESS & "','" & pOTHERAMNT & "','" & pINVOICETTLPAY & _
                                                        "','" & pRECVDBY & "','" & pDATERCVD & "','" & pCHCKNO & "','" & pVOCHERNO & "','" & pCHCKDATE & "','" & pDISNOTE & "'," & UserID & ",'" & pBRANCHCODE & _
                                                        "','" & pSUPPLIER & "','" & pSHIPPER & "','" & pPONO & "','" & pRFP & "','" & pWAYBILLNO & "');"
                
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_acctpayinvcng_create()")
End Sub

Public Sub sp_acctpayinvcng_update(pSPRRNO As Integer, pPACKINGNO As String, pPACKINGDATE As String, pTERMS As String, pDUEDATE As String, pINVDISC1 As String _
                                , pDISC1VAL As String, pDISC1TYPE As String, pINVDISC2 As String, pDISC2VAL As String, pDISC2TYPE As String _
                                , pINVDISC3 As String, pDISC3VAL As String, pDISC3TYPE As String, pINVDISC4 As String, pDISC4VAL As String _
                                , pDISC4TYPE As String, pINVDISC5 As String, pDISC5VAL As String, pDISC5TYPE As String, pVAT As String, pVATVALUE As String _
                                , pPREPAREDBY As String, pAPPROVEDBY As String, pINVTYPE As String, pINVSTATUS As String, pINVNOTES As String _
                                , pINVLESSPER As String, pINVLESSPER2 As String, pINVLESSPER3 As String, pINVOICELESS As Double, pOTHERAMNT As Double _
                                , pINVOICETTLPAY As Double, pRECVDBY As String, pDATERCVD As String, pCHCKNO As String, pVOCHERNO As String _
                                , pCHCKDATE As String, pDISNOTE As String, pBRANCHCODE As String, pSUPPLIER As String _
                                , pSHIPPER As String, pPONO As String, pRFP As String, pWAYBILLNO As String, ByRef pLevel As String, ByRef pMessage As String)
                                
On Error GoTo er
    DoEvents
    Qry = "Call sp_acctpayinvcng.sp_acctpayinvcng_update('" & pSPRRNO & "','" & pPACKINGNO & "','" & pPACKINGDATE & "','" & pTERMS & "','" & pDUEDATE & _
                                                        "','" & pINVDISC1 & "','" & pDISC1VAL & "','" & pDISC1TYPE & _
                                                        "','" & pINVDISC2 & "','" & pDISC2VAL & "','" & pDISC2TYPE & _
                                                        "','" & pINVDISC3 & "','" & pDISC3VAL & "','" & pDISC3TYPE & _
                                                        "','" & pINVDISC4 & "','" & pDISC4VAL & "','" & pDISC4TYPE & _
                                                        "','" & pINVDISC5 & "','" & pDISC5VAL & "','" & pDISC5TYPE & _
                                                        "','" & pVAT & "','" & pVATVALUE & "','" & pPREPAREDBY & "','" & pAPPROVEDBY & _
                                                        "','" & pINVTYPE & "','" & pINVSTATUS & "',' " & pINVNOTES & _
                                                        "','" & pINVLESSPER & "','" & pINVLESSPER2 & "','" & pINVLESSPER3 & "','" & pINVOICELESS & "','" & pOTHERAMNT & "','" & pINVOICETTLPAY & _
                                                        "','" & pRECVDBY & "','" & pDATERCVD & "','" & pCHCKNO & "','" & pVOCHERNO & "','" & pCHCKDATE & "','" & pDISNOTE & "'," & UserID & ",'" & pBRANCHCODE & _
                                                        "','" & pSUPPLIER & "','" & pSHIPPER & "','" & pPONO & "','" & pRFP & "','" & pWAYBILLNO & "');"
                
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_acctpayinvcng_update()")
End Sub

