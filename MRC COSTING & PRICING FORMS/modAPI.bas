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
        pLevel = objRs!Level
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
                                , pCHCKNO As String, pVOCHERNO As String, pCHCKDATE As String, pDISNOTE As String, ByRef pLevel As String, ByRef pMessage As String)
           
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
                                                        "','" & pRECVDBY & "','" & pDATERCVD & "','" & pCHCKNO & "','" & pVOCHERNO & "','" & pCHCKDATE & "','" & pDISNOTE & "'," & UserID & ");"
                
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

Public Sub sp_acctpayinvcng_update(pSPRRNO As Integer, pSUPNAME As String, pPACKINGNO As String, pTERMS As String, pDUEDATE As String, pINVDISC1 As String _
                                , pDISC1VAL As String, pDISC1TYPE As String, pINVDISC2 As String, pDISC2VAL As String, pDISC2TYPE As String _
                                , pINVDISC3 As String, pDISC3VAL As String, pDISC3TYPE As String, pINVDISC4 As String, pDISC4VAL As String _
                                , pDISC4TYPE As String, pINVDISC5 As String, pDISC5VAL As String, pDISC5TYPE As String, pVAT As String, pVATVALUE As String _
                                , pPREPAREDBY As String, pAPPROVEDBY As String, pINVTYPE As String, pINVSTATUS As String, pINVNOTES As String _
                                , pINVLESSPER As String, pINVLESSPER2 As String, pINVLESSPER3 As String, pINVOICELESS As Double, pOTHERAMNT As Double, pINVOICETTLPAY As Double, pRECVDBY As String, pDATERCVD As String _
                                , pCHCKNO As String, pVOCHERNO As String, pCHCKDATE As String, pDISNOTE As String, ByRef pLevel As String, ByRef pMessage As String)
                                
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
                                                        "','" & pRECVDBY & "','" & pDATERCVD & "','" & pCHCKNO & "','" & pVOCHERNO & "','" & pCHCKDATE & "','" & pDISNOTE & "'," & UserID & ");"
                
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

