Attribute VB_Name = "modAPI"
Option Explicit
Private Qry As String
Private rs As New ADODB.Recordset

Public Sub sp_3s_branch_list(objRs As ADODB.Recordset)
On Error GoTo er
    
    DoEvents
    Qry = "SELECT BranchCode,BranchName FROM mrcwarehouse.branch where branchcategory = '3S';"
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


Public Sub sp_get_packing_info(pPACKINGNO As String, objRs As ADODB.Recordset, ByRef pLevel As String, ByRef pMessage As String)
On Error GoTo er
    
'    PROCEDURE sp_acctpayinvcng.sp_get_packing_info(
'    pPACKINGNO VarChar(20)
'    pBRANCHCODE    Varchar(6)
'    )

    DoEvents
    Qry = "Call sp_acctpayinvcng.sp_get_packing_info('" & pPACKINGNO & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
    If objRs.State = 0 Then
        pLevel = "Error"
        pMessage = "Missing Output parameters!"
    ElseIf objRs.RecordCount > 0 Then
        pLevel = objRs!level
        pMessage = objRs!message
    End If
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_get_packing_info()")
End Sub

Public Sub sp_refresh_packing_info(pPACKINGNO As String, objRs As ADODB.Recordset)
On Error GoTo er

    'PROCEDURE sp_acctpayinvcng.sp_refresh_packing_info(
    '  pPACKINGNO VarChar(20)
    '  )
    
    DoEvents
    Qry = "Call sp_acctpayinvcng.sp_refresh_packing_info('" & pPACKINGNO & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_refresh_packing_info()")
End Sub

Public Sub sp_load_packing_info(pPACKINGNO As String, objRs As ADODB.Recordset)
On Error GoTo er

    'PROCEDURE sp_acctpayinvcng.sp_load_packing_info(
    '  pPACKINGNO VarChar(20)
    '  )
    
    DoEvents
    Qry = "Call sp_acctpayinvcng.sp_load_packing_info('" & pPACKINGNO & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_load_packing_info()")
End Sub


Public Sub sp_acctpayinvcng_create(pSPRRNO As Integer, pPACKINGNO As String, pTERMS As String, pDUEDATE As String, pINVDISC1 As String _
                                , pDISC1VAL As String, pDISC1TYPE As String, pINVDISC2 As String, pDISC2VAL As String, pDISC2TYPE As String _
                                , pINVDISC3 As String, pDISC3VAL As String, pDISC3TYPE As String, pINVDISC4 As String, pDISC4VAL As String _
                                , pDISC4TYPE As String, pINVDISC5 As String, pDISC5VAL As String, pDISC5TYPE As String, pVAT As String, pVATVALUE As String _
                                , pPREPAREDBY As String, pAPPROVEDBY As String, pINVTYPE As String, pINVSTATUS As String, pINVNOTES As String _
                                , pINVLESSPER As String, pINVLESSPER2 As String, pINVLESSPER3 As String, pINVOICELESS As Double, pOTHERAMNT As Double, pINVOICETTLPAY As Double, pRECVDBY As String, pDATERCVD As String _
                                , pCHCKNO As String, pVOCHERNO As String, pCHCKDATE As String, pDISNOTE As String, ByRef pLevel As String, ByRef pMessage As String, pDELIVERYNO As String, pINVOICENO As String, pSPCRRNO As String)
           
On Error GoTo er

    DoEvents
    
    'PROCEDURE sp_acctpayinvcng.sp_acctpayinvcng_create(
    '       pSPRRNO Int(11)
    '      ,pPACKINGNO           varchar(20)
    '      ,pTERMS               int(10)
    '      ,pDUEDATE             date
    '      ,pINVDISC1            VARCHAR(15)
    '      ,pDISC1VAL            VARCHAR(15)
    '      ,pDISC1TYPE           VARCHAR(50)
    '      ,pINVDISC2            VARCHAR(15)
    '      ,pDISC2VAL            VARCHAR(15)
    '      ,pDISC2TYPE           VARCHAR(50)
    '      ,pINVDISC3            VARCHAR(15)
    '      ,pDISC3VAL            VARCHAR(15)
    '      ,pDISC3TYPE           VARCHAR(50)
    '      ,pINVDISC4            VARCHAR(15)
    '      ,pDISC4VAL            VARCHAR(15)
    '      ,pDISC4TYPE           VARCHAR(50)
    '      ,pINVDISC5            VARCHAR(15)
    '      ,pDISC5VAL            VARCHAR(15)
    '      ,pDISC5TYPE           VARCHAR(50)
    '      ,pVAT                 VARCHAR(15)
    '      ,pVATVALUE            Decimal(16,4)
    '      ,pPREPAREDBY          varchar(50)
    '      ,pAPPROVEDBY          varchar(50)
    '      ,pINVTYPE             varchar(25)
    '      ,pINVSTATUS           varchar(10)
    '      ,pINVNOTES            varchar(100)
    '      ,pINVLESSPER          varchar(5)
    '      ,pINVLESSPER2          varchar(5)
    '      ,pINVLESSPER3          varchar(5)
    '      ,pINVOICELESS         decimal(16,4)
    '      ,pOTHERAMNT           decimal(16,4)
    '      ,pNCTTLPAY            decimal(16,4)
    '      ,pINVOICETTLPAY       decimal(16,4)
    '      ,pRECVDBY             varchar(50)
    '      ,pDATERCVD            date
    '      ,pCHCKNO              varchar(45)
    '      ,pVOCHERNO            varchar(45)
    '      ,pCHCKDATE            varchar(45)
    '      ,pDISNOTE             varchar(200)
    '      ,pENCODER             int
    '      ,pBRANCHCODE          varchar(6)
'    )

    Qry = "Call sp_acctpayinvcng.sp_acctpayinvcng_create('" & pSPRRNO & "','" & pPACKINGNO & "','" & pTERMS & "','" & pDUEDATE & _
                                                        "','" & pINVDISC1 & "','" & pDISC1VAL & "','" & pDISC1TYPE & _
                                                        "','" & pINVDISC2 & "','" & pDISC2VAL & "','" & pDISC2TYPE & _
                                                        "','" & pINVDISC3 & "','" & pDISC3VAL & "','" & pDISC3TYPE & _
                                                        "','" & pINVDISC4 & "','" & pDISC4VAL & "','" & pDISC4TYPE & _
                                                        "','" & pINVDISC5 & "','" & pDISC5VAL & "','" & pDISC5TYPE & _
                                                        "','" & pVAT & "','" & pVATVALUE & "','" & pPREPAREDBY & "','" & pAPPROVEDBY & _
                                                        "','" & pINVTYPE & "','" & pINVSTATUS & "',' " & pINVNOTES & _
                                                        "','" & pINVLESSPER & "','" & pINVLESSPER2 & "','" & pINVLESSPER3 & "','" & pINVOICELESS & "','" & pOTHERAMNT & "','" & pINVOICETTLPAY & _
                                                        "','" & pRECVDBY & "','" & pDATERCVD & "','" & pCHCKNO & "','" & pVOCHERNO & "','" & pCHCKDATE & "','" & pDISNOTE & "'," & UserID & ",'" & pDELIVERYNO & "','" & pINVOICENO & "','" & pSPCRRNO & "');"
                
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_acctpayinvcng_create()")
End Sub

Public Sub sp_acctpayinvcng_update(pSPRRNO As Integer, pSUPNAME As String, pPACKINGNO As String, pTERMS As String, pDUEDATE As String, pINVDISC1 As String _
                                , pDISC1VAL As String, pDISC1TYPE As String, pINVDISC2 As String, pDISC2VAL As String, pDISC2TYPE As String _
                                , pINVDISC3 As String, pDISC3VAL As String, pDISC3TYPE As String, pINVDISC4 As String, pDISC4VAL As String _
                                , pDISC4TYPE As String, pINVDISC5 As String, pDISC5VAL As String, pDISC5TYPE As String, pVAT As String, pVATVALUE As String _
                                , pPREPAREDBY As String, pAPPROVEDBY As String, pINVTYPE As String, pINVSTATUS As String, pINVNOTES As String _
                                , pINVLESSPER As String, pINVLESSPER2 As String, pINVLESSPER3 As String, pINVOICELESS As Double, pOTHERAMNT As Double, pINVOICETTLPAY As Double, pRECVDBY As String, pDATERCVD As String _
                                , pCHCKNO As String, pVOCHERNO As String, pCHCKDATE As String, pDISNOTE As String, ByRef pLevel As String, ByRef pMessage As String, pDELIVERYNO As String, pINVOICENO As String, pSPCRRNO As String)
                                
On Error GoTo er
    DoEvents
    
    'PROCEDURE sp_acctpayinvcng.sp_acctpayinvcng_create(
    '       pSPRRNO Int(11)
    '      ,pSUPNAME             Varchar(50)
    '      ,pPACKINGNO           varchar(20)
    '      ,pTERMS               int(10)
    '      ,pDUEDATE             date
    '      ,pINVDISC1            VARCHAR(15)
    '      ,pDISC1VAL            VARCHAR(15)
    '      ,pDISC1TYPE           VARCHAR(50)
    '      ,pINVDISC2            VARCHAR(15)
    '      ,pDISC2VAL            VARCHAR(15)
    '      ,pDISC2TYPE           VARCHAR(50)
    '      ,pINVDISC3            VARCHAR(15)
    '      ,pDISC3VAL            VARCHAR(15)
    '      ,pDISC3TYPE           VARCHAR(50)
    '      ,pINVDISC4            VARCHAR(15)
    '      ,pDISC4VAL            VARCHAR(15)
    '      ,pDISC4TYPE           VARCHAR(50)
    '      ,pINVDISC5            VARCHAR(15)
    '      ,pDISC5VAL            VARCHAR(15)
    '      ,pDISC5TYPE           VARCHAR(50)
    '      ,pVAT                 VARCHAR(15)
    '      ,pVATVALUE            Decimal(16,4)
    '      ,pPREPAREDBY          varchar(50)
    '      ,pAPPROVEDBY          varchar(50)
    '      ,pINVSTATUS           varchar(10)
    '      ,pINVNOTES            varchar(100)
    '      ,pNCLESSPER           varchar(5)
    '      ,pNCLESSPER2           varchar(5)
    '      ,pNCLESSPER3           varchar(5)
    '      ,pNCLESS              decimal(16,4)
    '      ,pINVLESSPER          varchar(5)
    '      ,pINVLESSPER2          varchar(5)
    '      ,pINVLESSPER3          varchar(5)
    '      ,pINVOICELESS         decimal(16,4)
    '      ,pOTHERAMNT           decimal(16,4)
    '      ,pNCTTLPAY            decimal(16,4)
    '      ,pINVOICETTLPAY       decimal(16,4)
    '      ,pRECVDBY             varchar(50)
    '      ,pDATERCVD            date
    '      ,pCHCKNO              varchar(45)
    '      ,pVOCHERNO            varchar(45)
    '      ,pCHCKDATE            varchar(45)
    '      ,pDISNOTE             varchar(200)
    '      ,pENCODER             int
    '      ,pBRANCHCODE          varchar(6)
'    )

    Qry = "Call sp_acctpayinvcng.sp_acctpayinvcng_update('" & pSPRRNO & "','" & pSUPNAME & "','" & pPACKINGNO & "','" & pTERMS & "','" & pDUEDATE & _
                                                        "','" & pINVDISC1 & "','" & pDISC1VAL & "','" & pDISC1TYPE & _
                                                        "','" & pINVDISC2 & "','" & pDISC2VAL & "','" & pDISC2TYPE & _
                                                        "','" & pINVDISC3 & "','" & pDISC3VAL & "','" & pDISC3TYPE & _
                                                        "','" & pINVDISC4 & "','" & pDISC4VAL & "','" & pDISC4TYPE & _
                                                        "','" & pINVDISC5 & "','" & pDISC5VAL & "','" & pDISC5TYPE & _
                                                        "','" & pVAT & "','" & pVATVALUE & "','" & pPREPAREDBY & "','" & pAPPROVEDBY & _
                                                        "','" & pINVTYPE & "','" & pINVSTATUS & "','" & pINVNOTES & _
                                                        "','" & pINVLESSPER & "','" & pINVLESSPER2 & "','" & pINVLESSPER3 & "','" & pINVOICELESS & "','" & pOTHERAMNT & "','" & pINVOICETTLPAY & _
                                                        "','" & pRECVDBY & "','" & pDATERCVD & "','" & pCHCKNO & "','" & pVOCHERNO & "','" & pCHCKDATE & "','" & pDISNOTE & "'," & UserID & ",'" & pDELIVERYNO & "','" & pINVOICENO & "','" & pSPCRRNO & "');"
                
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_acctpayinvcng_update()")
End Sub

Public Sub sp_load_acct_title(pRFPID As Long, objRs As ADODB.Recordset)
On Error GoTo er
    
    DoEvents
    Qry = "SELECT b.BRANCHNAME, ch.TRANSACTION as TransName,vou.code,vou.Description,vou.Type,vou.Amount,vou.TotDebit, vou.TotCredit FROM mrcwarehouse.voucherdtls as vou " & _
          "LEFT JOIN mrcwarehouse.branch as b ON b.BRANCHCODE = vou.BCODE LEFT JOIN accounting_corp.trans_legend as ch ON vou.TransID = ch.Trans_Legend_ID WHERE vou.RFPID = " & pRFPID & " GROUP by vou.SeqNO Order by vou.SeqNo asc;"
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

Public Sub sp_sync_acct_titles(ByRef pLevel As String, ByRef pMessage As String)
On Error GoTo er
    
    DoEvents
    Qry = "Call sp_acctpayinvcng.sp_sync_acct_titles(" & UserID & ");"
    
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
    Qry = "SELECT DISTINCT tl.Trans_Legend_ID, tl.TRANSACTION FROM accounting_corp.trans_legend as tl Where tl.Status='A' and tl.BUNIT_CODE IN (4,14,21) and tl.DEPCODE IN ('PAP','CAR','CAD') and tl.AcctTransFor IN ('CHECK DISBURSEMENT','PURCHASE BOOK') ORDER BY TRANSACTION ASC;"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_transaction_list()")
End Sub

Public Sub sp_AccountTitles_List(objRs As ADODB.Recordset)
On Error GoTo er
    
    DoEvents
    Qry = "SELECT tlDtls.ACCT_CODE as AcctCode, coa.ACCT_TITTLE as AcctTitles, tlDtls.DRCR as `TYPE` FROM accounting_corp.trans_legend_details as tlDtls Join accounting_corp.trans_legend as tl ON tl.TRANS_LEGEND_ID = tlDtls.TRANS_LEGEND_ID Join accounting_corp.chart_of_accts as coa ON coa.ACCT_CODE = tlDtls.ACCT_CODE and coa.BUNIT_CODE IN (4,14, 21) Where tl.Status='A' and tl.DEPCODE IN ('PAP','CAR','CAD') and tl.AcctTransFor IN ('CHECK DISBURSEMENT','PURCHASE BOOK') " & _
          "GROUP BY tlDtls.ACCT_CODE ORDER BY tlDtls.ACCT_CODE ASC;"
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


Public Sub sp_get_AccountTitles_Auto(pTransID As Long, objRs As ADODB.Recordset)
On Error GoTo er
    
    DoEvents
    Qry = "SELECT tlDtls.ACCT_CODE as AcctCode, coa.ACCT_TITTLE as AcctTitles, tlDtls.DRCR as `TYPE` FROM accounting_corp.trans_legend as tl Join accounting_corp.trans_legend_details as tlDtls ON tlDtls.Trans_Legend_ID = tl.Trans_Legend_ID Join accounting_corp.chart_of_accts as coa ON coa.ACCT_CODE = tlDtls.ACCT_CODE and coa.BUNIT_CODE IN (4,14) where tl.Trans_Legend_ID = " & pTransID & " and tlDtls.ViewOption = 'AUTO' GROUP BY tlDtls.ACCT_CODE ORDER BY tlDtls.sqno ASC;"
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

Public Sub sp_company_list(objRs As ADODB.Recordset)
On Error GoTo er
    
    DoEvents
    Qry = "SELECT co_Name FROM mrcwarehouse.company ORDER BY co_Name ASC;"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_company_list()")
End Sub

Public Sub sp_load_rfpother_info(pRFPID As Long, objRs As ADODB.Recordset)
On Error GoTo er
    
    DoEvents
    Qry = "Call sp_mrcwarehouse.sp_get_rfp_other(" & pRFPID & ");"
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

Public Sub sp_rfp_other_create(pRFPNO As String, pRFPDATE As String, pDUEDATE As String, pSTATUTORY As String, pPAYEE As String, pPREPBY As String, pCHCKEDBY As String _
                                , pNOTEDBY As String, pAPPRVDBY As String, pCHECKNO As String, pCHECKDATE As String _
                                , pREMARKS As String, pSTATUS As String, pTRANSNAME As String, pBILLTO As String, pCOMPANY As String, pDOCDATEREC As String, ByRef pLevel As String, ByRef pMessage As String)
           
On Error GoTo er
DoEvents

    'PROCEDURE sp_mrcwarehouse.sp_rfp_other_create(
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
  
    Qry = "Call sp_mrcwarehouse.sp_rfp_other_create('" & pRFPNO & "','" & Format(pRFPDATE, "YYYY-mm-dd") & "','" & Format(pDUEDATE, "YYYY-mm-dd") & "','" & pSTATUTORY & "','" & pPAYEE & "','" & pPREPBY & _
                                                        "','" & pCHCKEDBY & "','" & pNOTEDBY & "','" & pAPPRVDBY & "','" & Trim(pCHECKNO) & "','" & pCHECKDATE & _
                                                        "','" & pREMARKS & "','" & pSTATUS & "', '" & pTRANSNAME & "'," & UserID & ",'" & pBILLTO & "','" & pCOMPANY & "','" & Format(pDOCDATEREC, "YYYY-mm-dd") & "'  );"
                
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

Public Sub sp_rfp_other_update(pRFPID As Long, pRFPNO As String, pRFPDATE As String, pDUEDATE As String, pSTATUTORY As String, pPAYEE As String, pPREPBY As String, pCHCKEDBY As String, pNOTEDBY As String _
                                , pAPPRVDBY As String, pCHECKNO As String, pCHECKDATE As String _
                                , pREMARKS As String, pSTATUS As String, pTRANSNAME As String, pBILLTO As String, pCOMPANY As String, pDOCDATEREC As String, ByRef pLevel As String, ByRef pMessage As String)
           
On Error GoTo er
DoEvents

    'PROCEDURE sp_mrcwarehouse.sp_rfp_other_update(
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
  
    Qry = "Call sp_mrcwarehouse.sp_rfp_other_update(" & pRFPID & ",'" & pRFPNO & "','" & Format(pRFPDATE, "YYYY-mm-dd") & "','" & Format(pDUEDATE, "YYYY-mm-dd") & "','" & pSTATUTORY & "','" & pPAYEE & "','" & pPREPBY & _
                                                        "','" & pCHCKEDBY & "','" & pNOTEDBY & "','" & pAPPRVDBY & "','" & Trim(pCHECKNO) & "','" & pCHECKDATE & _
                                                        "','" & pREMARKS & "','" & pSTATUS & "',  '" & pTRANSNAME & "'," & UserID & ",'" & pBILLTO & "','" & pCOMPANY & "','" & Format(pDOCDATEREC, "YYYY-mm-dd") & "' );"
                
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

Public Sub sp_load_rfp_deduction(pRFPID As Integer, objRs As ADODB.Recordset)
On Error GoTo er
    
    DoEvents
    Qry = "SELECT DESCRIPTION as DTYPE,DEDUCTTYPE as TYPE, DEDUCTION_REB as REBDUCT, PARTICULARS as DESCRIPTION FROM rfp_deduct where RFPID = " & pRFPID & ";"
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
