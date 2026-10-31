Attribute VB_Name = "modSRPREQ"
Option Explicit
Private Qry As String
Private rs As New ADODB.Recordset

Public Sub sp_srppricing_update(pREFNO As String, pDATEFINAL As String, pREMARKS As String, pENCODEDBY As String, pRECOMBY As String, pNOTEDBY As String, pAPPROVEDBY As String, ByRef pLevel As String, ByRef pMessage As String)
                                
On Error GoTo er

    DoEvents
    
    'PROCEDURE sp_mrcretails.sp_srppricing_update(
    '    pBranchCode   Varchar(5)
    '  , pREFNO        Integer
    '  , pDATEFINAL    Varchar(60)
    '  , pREMARKS      Varchar(200)
    '  , pENCODEDBY    Varchar(60)
    '  , pRECOMBY      Varchar(60)
    '  , pNOTEDBY      Varchar(60)
    '  , pAPPROVEDBY   Varchar(60)
    '  , pENCODER      Integer
    '  )

    Qry = "Call sp_mrcretails.sp_srppricing_update('" & sysBranchCode & "','" & pREFNO & "','" & pDATEFINAL & "','" & pREMARKS & "','" & pENCODEDBY & "','" & pRECOMBY & "','" & pNOTEDBY & "','" & pAPPROVEDBY & "'," & UserID & ");"
                
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_srppricing_update()")
End Sub



Public Sub sp_srppricing_summary(pDATEFROM As String, pDATETO As String, objRs As ADODB.Recordset)
On Error GoTo er

    DoEvents
    
    'PROCEDURE sp_mrcretails.sp_srppricing_summary(
    '   pBRANCHCODE VarChar(5)
    '  ,pDATEFROM      Varchar(60)
    '  ,pDATETO        Varchar(60)
    '  )
    
    Qry = "Call sp_mrcretails.sp_srppricing_summary('" & sysBranchCode & "','" & pDATEFROM & "','" & pDATETO & "');"
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
    
    'PROCEDURE sp_mrcretails.sp_getRefNo(pBranchCode Varchar(5))
    
    DoEvents
    Qry = "Call sp_mrcretails.sp_getRefNo('" & sysBranchCode & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
    Screen.MousePointer = vbDefault
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_get_refno()")
End Sub

Public Sub sp_partno_desc(pPARTNO As String, objRs As ADODB.Recordset)
On Error GoTo er
    
    'PROCEDURE sp_mrcretails.sp_partno_desc(
    '   pBRANCHCODE VarChar(5)
    '  ,pPartNo     Varchar(20)
    '  )
    
    DoEvents
    Qry = "Call sp_mrcretails.sp_partno_desc('" & sysBranchCode & "','" & pPARTNO & "');"
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

Public Sub sp_show_record(pBREFNO As String, objRs As ADODB.Recordset)
On Error GoTo er
    
'PROCEDURE sp_mrcretails.sp_show_record(
'  pBRANCHCODE VarChar(4)
'  ,pREFNO Integer
'  )
    
    DoEvents
    Qry = "Call sp_mrcretails.sp_show_record('" & sysBranchCode & "','" & pBREFNO & "');"
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

Public Sub sp_srppricing_create(pDATEFINAL As String, pSTATUS As String, pREMARKS As String, pENCODEDBY As String, pRECOMBY As String, pNOTEDBY As String, pAPPROVEDBY As String, ByRef pLevel As String, ByRef pMessage As String)
           
On Error GoTo er

    DoEvents
    
    'PROCEDURE sp_mrcretails.sp_srppricing_create(
    '    pBranchCode   Varchar(5)
    '  , pDATEFINAL    Varchar(60)
    '  , pSTATUS       Varchar(10)
    '  , pREMARKS      Varchar(200)
    '  , pENCODEDBY    Varchar(60)
    '  , pRECOMBY      Varchar(60)
    '  , pNOTEDBY      Varchar(60)
    '  , pAPPROVEDBY   Varchar(60)
    '  )

    Qry = "Call sp_mrcretails.sp_srppricing_create('" & sysBranchCode & "','" & pDATEFINAL & "','" & pSTATUS & "','" & pREMARKS & "','" & pENCODEDBY & "','" & pRECOMBY & "','" & pNOTEDBY & "','" & pAPPROVEDBY & "');"
                
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->sp_srppricing_create()")
End Sub

