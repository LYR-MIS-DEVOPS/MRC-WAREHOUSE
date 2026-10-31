Attribute VB_Name = "modMotoR"
Global LYRcon       As New ADODB.Connection
Global Oldcon       As New ADODB.Connection
'Global MainCon     As New ADODB.Connection      'E remarks kung mag compile para sa main
Global OnlineCon    As New ADODB.Connection
Public ObjKeyTrapping As New clsKeyTrapping
'------------------------------------------------------------------------------------------
Global BrchConMU    As New ADODB.Connection
Global BrchConYM    As New ADODB.Connection
Global BrchConMD    As New ADODB.Connection
Global BrchConVAN   As New ADODB.Connection
Global conUlas      As New ADODB.Connection
Global conPanabo    As New ADODB.Connection
Global conPol       As New ADODB.Connection
Global conToril     As New ADODB.Connection
Global conMRC2      As New ADODB.Connection
Global conVAN1      As New ADODB.Connection
Global conVAN2      As New ADODB.Connection
Global conMRC1      As New ADODB.Connection
Global conYM        As New ADODB.Connection
Global conGenSan    As New ADODB.Connection
Global conWsale     As New ADODB.Connection
Global conWhouse    As New ADODB.Connection
Global conYmarbel   As New ADODB.Connection
Global conStotomas  As New ADODB.Connection
Global conSamal     As New ADODB.Connection
Global conVan3      As New ADODB.Connection
Global conCabantian As New ADODB.Connection
Global conPartsCat As New ADODB.Connection
Global rsAnyUse     As New ADODB.Recordset
'Global whouseEnd    As New warehouseEndpoint.clsConnection

'user variables
'Global UserID       As Integer
'Global FrDate       As Date
'Global ToDate       As Date
'Global DispBtn      As Boolean
'Global UserType     As String
'
'Global SysBranch    As String
'Global SysAddress   As String
'Global SysTel       As String
'Global DateNow      As String
'Global CurrDate     As String
'Global LstDayMnth   As String

Global UserID           As Integer
Global UserName         As String
Global FrDate           As Date
Global ToDate           As Date
Global DispBtn          As Boolean
Global UserType         As String
Global UserPosition     As String
Global ServerTime       As String
Global SysBranch        As String
Global sysBranchCode    As String
Global SysAddress       As String
Global SysTel           As String
Global SysBranchID      As Integer
Global SysBranchOIC     As String
Global wCost            As Long
Global CurrDate         As String
Global LstDayMnth       As String
Public LoadOnProcess As Boolean
'----------------------------------------------------------------------------------------
'report function declaration
Declare Function CreateFieldDefFile Lib "p2smon.dll" (lpUnk As Object, ByVal filename As String, ByVal bOverWriteExistingFile As Long) As Long

'variable for saving/updating pictures
Global Chunk()      As Byte
Global lngLengh     As Long
Global intChunks    As Integer
Global intFragment  As Integer
'Global Const ChunkSize = 1000 '16384
Global Const ChunkSize = 16384

Global Const lngDataFile = 1
'searching variables

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Global Const CB_FINDSTRING = &H14C
Declare Function CharUpper Lib "user32" Alias "CharUpperA" (ByVal lpsz As String) As String

'----- START CREATE ODBC
Private Const REG_SZ = 1    'Constant for a string variable type.
Private Const HKEY_LOCAL_MACHINE = &H80000002

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias _
       "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, _
       phkResult As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
       "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
       ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal _
       cbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" _
       (ByVal hKey As Long) As Long
'------- END ODBC

'Public OldWindowProc As Long
'Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'
'Public Const GWL_WNDPROC = (-4)
'Public Const WM_USER = &H400
'' Pass along all messages except the one that
'' makes the context menu appear.
'Public Function NoPopupWindowProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Const WM_CONTEXTMENU = &H7B
'    If msg <> WM_CONTEXTMENU Then _
'        NoPopupWindowProc = CallWindowProc( _
'            OldWindowProc, hwnd, msg, wParam, _
'            lParam)
'End Function

Sub Main()
    On Error GoTo LocalErr
    LYRcon.CursorLocation = adUseClient
    LYRcon.ConnectionTimeout = 60
    
    
    LYRcon.Open "dsn=mrcwarehouse"
    'LYRcon.Open "dsn=lyrmpsobrecarywh"
    'LYRcon.Open "dsn=whouseTesting"
    'LYRcon.ConnectionString = whouseEnd.getConnnection(29)
    'LYRcon.Open
    
    Call SBranch
    
    Warehouse.Show
    LogIn.Show vbModal
    'frmSplash.Show vbModal
   Exit Sub

LocalErr:
    MsgBox Err.Description '"Error connecting to server..." & vbCrLf & "Make sure you are connected to the Server before running the program.", vbExclamation + vbOKOnly, "Connection Failed"
    
End Sub

Public Function ConnectPartsCatalog() As Boolean

    On Error GoTo LocalErr
    If conPartsCat.State = adStateOpen Then: conPartsCat.Close
    conPartsCat.CursorLocation = adUseClient
    conPartsCat.ConnectionTimeout = 60
    conPartsCat.Open "dsn=partscatalog"
    ConnectPartsCatalog = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to Parts Catalog Database...."), vbExclamation
    ConnectPartsCatalog = False
End Function

Public Sub OldServer183()
    On Error GoTo LocalErr
    If Oldcon.State = adStateOpen Then: Oldcon.Close
    Oldcon.CursorLocation = adUseClient
    Oldcon.ConnectionTimeout = 60
    
    Oldcon.Open "dsn=sp_mrc_backup_monitoring"
   Exit Sub

LocalErr:
    MsgBox Err.Description '"Error connecting to server..." & vbCrLf & "Make sure you are connected to the Server before running the program.", vbExclamation + vbOKOnly, "Connection Failed"
    
End Sub

'Public Sub ConnectMain()
'
'    On Error GoTo LocalErr
'    OnlineCon.CursorLocation = adUseClient
'    OnlineCon.ConnectionTimeout = 60
'    OnlineCon.Open "dsn=sqlmcpdepotonline"
'
'    Exit Sub
'
'LocalErr:
'    MsgBox ("CANNOT CONNECT TO MRC WAREHOUSE. YOUR P.O. SAVE LOCALLY"), vbExclamation
'
'End Sub

'Public Sub ConnectOnline(branchID As String)
'    Dim rsBrn As New ADODB.Recordset
'    Dim brchName As String
'    Dim mBrc As String
'
'    On Error GoTo LocalErr
'
'    mBrc = "SELECT branch.BRANCHNAME From Branch Where branch.BRANCHID = ' & val(branchid) & ';"
'        If branchID = 15 Then 'mcpdepot
'            OnlineCon.CursorLocation = adUseClient
'            OnlineCon.ConnectionTimeout = 60
'            OnlineCon.Open "dsn=sqlmcpdepotonline"
'
'            With rsBrn
'                .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
'                    If .RecordCount > 0 Then
'                        .MoveFirst
'                            brchName = !BRANCHNAME
'                    End If
'                .Close
'            End With
'
'        ElseIf branchID = 13 Then 'mrc ulas
'            OnlineCon.CursorLocation = adUseClient
'            OnlineCon.ConnectionTimeout = 60
'            OnlineCon.Open "dsn=sqlmtrulasonline"
'
'            With rsBrn
'                .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
'                    If .RecordCount > 0 Then
'                        .MoveFirst
'                            brchName = !BRANCHNAME
'                    End If
'                .Close
'            End With
'
'         ElseIf branchID = 12 Then 'mrc sobrecarey
'            OnlineCon.CursorLocation = adUseClient
'            OnlineCon.ConnectionTimeout = 60
'            OnlineCon.Open "dsn=sqlmtrtgmonline"
'
'            With rsBrn
'                .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
'                    If .RecordCount > 0 Then
'                        .MoveFirst
'                            brchName = !BRANCHNAME
'                    End If
'                .Close
'            End With
'
'        ElseIf branchID = 16 Then
'            OnlineCon.CursorLocation = adUseClient
'            OnlineCon.ConnectionTimeout = 60
'            OnlineCon.Open "dsn=sqlmtrvanonline"
'
'            With rsBrn
'                .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
'                    If .RecordCount > 0 Then
'                        .MoveFirst
'                            brchName = !BRANCHNAME
'                    End If
'                .Close
'            End With
'
'        ElseIf branchID = 22 Then
'            OnlineCon.CursorLocation = adUseClient
'            OnlineCon.ConnectionTimeout = 60
'            OnlineCon.Open "dsn=mrcvan2online"
'
'            With rsBrn
'                .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
'                    If .RecordCount > 0 Then
'                        .MoveFirst
'                            brchName = !BRANCHNAME
'                    End If
'                .Close
'            End With
'
'        ElseIf branchID = 27 Then
'            OnlineCon.CursorLocation = adUseClient
'            OnlineCon.ConnectionTimeout = 60
'            OnlineCon.Open "dsn=mrcvan3online"
'
'            With rsBrn
'                .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
'                    If .RecordCount > 0 Then
'                        .MoveFirst
'                            brchName = !BRANCHNAME
'                    End If
'                .Close
'            End With
'
'         ElseIf branchID = 28 Then
'            OnlineCon.CursorLocation = adUseClient
'            OnlineCon.ConnectionTimeout = 60
'            OnlineCon.Open "dsn=mrc2retailonline"
'
'            With rsBrn
'                .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
'                    If .RecordCount > 0 Then
'                        .MoveFirst
'                            brchName = !BRANCHNAME
'                    End If
'                .Close
'            End With
'
'        ElseIf branchID = 21 Then
'            OnlineCon.CursorLocation = adUseClient
'            OnlineCon.ConnectionTimeout = 60
'            OnlineCon.Open "dsn=sqlmrc2tgmonline"
'
'            With rsBrn
'                .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
'                    If .RecordCount > 0 Then
'                        .MoveFirst
'                            brchName = !BRANCHNAME
'                    End If
'                .Close
'            End With
'
'
'        End If
'    Exit Sub
'
'LocalErr:
'    MsgBox ("CANNOT CONNECT TO " & brchName & ". YOUR P.O. SAVE LOCALLY"), vbExclamation
'
'End Sub

Public Sub ConnectOnline(branchID As String)
    Dim rsBrn As New ADODB.Recordset
    Dim brchName As String
    Dim mBrc As String
    
    On Error GoTo LocalErr
    
    mBrc = "SELECT branch.BRANCHNAME From Branch Where branch.BRANCHID = ' & val(branchid) & ';"

    If branchID = 12 Then 'mrc ulas
            If OnlineCon.State = adStateOpen Then: OnlineCon.Close
            OnlineCon.CursorLocation = adUseClient
            OnlineCon.ConnectionTimeout = 60
            OnlineCon.Open "dsn=sqlmtrtgmonline"
            
            With rsBrn
                .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                    If .RecordCount > 0 Then
                        .MoveFirst
                            brchName = !branchName
                    End If
                .Close
            End With
    End If

'---------------------------------------------------------------------------------------------------------------
    If branchID = 15 Then 'mcpdepot
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=sqlmcpdepotonline"

        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
'---------------------------------------------------------------------------------------------------------------
    ElseIf branchID = 13 Then 'mrc ulas
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=sqlmtrulasonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With

'---------------------------------------------------------------------------------------------------------------
    ElseIf branchID = 16 Then 'mrc van 1
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=sqlmtrvanonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
'---------------------------------------------------------------------------------------------------------------
    ElseIf branchID = 22 Then 'mrc van 2
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrcvan2online"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
'---------------------------------------------------------------------------------------------------------------
    ElseIf branchID = 27 Then 'mrc van 3
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrcvan3online"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
'---------------------------------------------------------------------------------------------------------------
     ElseIf branchID = 28 Then 'mrc2retail
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrc2retailonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
'---------------------------------------------------------------------------------------------------------------
    ElseIf branchID = 21 Then 'mrc 2 wholesale
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=sqlmrc2tgmonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
        
    ElseIf branchID = 19 Then 'MRC PANABO
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrcpanaboonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
        
    ElseIf branchID = 20 Then 'HONDA POLOMOLOK
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrcpolomolokonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
    
     ElseIf branchID = 26 Then 'MOTORHUB MARBEL
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrcmarbelonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
        
    ElseIf branchID = 23 Then 'MRC GENSAN
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrcgensanonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
    
    ElseIf branchID = 24 Then 'HONDA STOTOMAS
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=bsstotomasonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
        
    ElseIf branchID = 25 Then 'HONDA SAMAL
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrcsamalonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
    
    ElseIf branchID = 14 Then 'SPAREPARTS MOTORHUB TAGUM
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrcMotorhubTagumonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
        
    ElseIf branchID = 30 Then 'MRC CABANTIAN
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrccabantianonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
    ElseIf branchID = 32 Then 'MC Parts Depot Rizal TAgum
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mcpdrizalonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
    ElseIf branchID = 33 Then 'Binos SpareParts Tagum
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrcbinostagumonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
'************************ CODE UPDATED: 07/14/17 *********************
    ElseIf branchID = 34 Then 'MRC MATINA
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrcmatinaonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
    ElseIf branchID = 36 Then 'MRCVAN5 - aka VAN74
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrcvan5online"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
    ElseIf branchID = 31 Then 'MRCVAN4 - VAN73
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrcvan4online"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
    ElseIf branchID = 43 Then 'MRCKAPALONG
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrckapalongonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
    ElseIf branchID = 42 Then 'MRCMANKILAM
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrcmankilamonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
    ElseIf branchID = 45 Then 'MRCAPOKON
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrcapokononline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
    ElseIf branchID = 46 Then 'MRCSTOTOMAS
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrcstotomasonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    brchName = !branchName
                End If
            .Close
        End With
    ElseIf branchID = 37 Then 'MRC MOTORHUB ULAS
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrcmotorhubulasonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
    ElseIf branchID = 39 Then 'MRC PIAGGIO TAGUM
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrcpiaggiotgmonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
    ElseIf branchID = 47 Then 'MRC MONKAYO
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrcmonkayoonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
    ElseIf branchID = 48 Then 'MRC MATI RIZAL
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrcmatirizalonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
    ElseIf branchID = 49 Then 'MRC TORIL
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrctorilonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
    ElseIf branchID = 50 Then 'MRC APLAYA
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrcaplayaonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
    ElseIf branchID = 53 Then 'MC PARTS DEPOT PUBLIC MARKET
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrctagumpmonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
    ElseIf branchID = 54 Then 'MRC CALINAN
        If OnlineCon.State = adStateOpen Then: OnlineCon.Close
        OnlineCon.CursorLocation = adUseClient
        OnlineCon.ConnectionTimeout = 60
        OnlineCon.Open "dsn=mrccalinanonline"
        
        With rsBrn
            .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                        brchName = !branchName
                End If
            .Close
        End With
    End If
'******************** END CODE UPDATED: 07/14/17 *********************
    Exit Sub
    
LocalErr:
    MsgBox ("CANNOT CONNECT TO " & mBrc & ". YOUR P.O. SAVE LOCALLY"), vbExclamation

End Sub


Public Sub CenterFrm(ByVal Parentfrm As MDIForm, ByVal Childfrm As Form)
    Childfrm.Left = (Parentfrm.Width \ 2) - (Childfrm.Width \ 2)
    Childfrm.Top = (Parentfrm.Height \ 2) - (Childfrm.Height \ 2) - 600
End Sub

Public Function MyPaperSize() As Integer
    Dim MyPrn As Printer
    Dim i As Integer
    'Dim doctoprint As New System.Drawing.Printing.PrintDocument
    
    'MyPaperSize = MyPrn.PaperSize
    
   '         doctoprint.PrinterSettings.PrinterName = "YourPrinterName" '(ex. "Epson SQ-1170 ESC/P 2")
            For i = 0 To doctoprint.PrinterSettings.PaperSizes.Count - 1
                Dim rawKind As Integer
                If doctoprint.PrinterSettings.PaperSizes(i).PaperName = "ADEL 8.5X5.5" Then
                    rawKind = CInt(doctoprint.PrinterSettings.PaperSizes(i).GetType().GetField("ADEL 8.5X5.5", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic).GetValue(doctoprint.PrinterSettings.PaperSizes(i)))
                    MyPaperSize = rawKind
                    Exit Function
               End If
            Next
     MyPaperSize = 0
End Function

Public Sub PrintMyForm(ByRef MyVal() As String, FormName As String, ByVal MyFlx As MSFlexGrid)
    Dim rsPref As New ADODB.Recordset
    Dim mcLyr As CRAXDRT.Application
    Dim mcReport As CRAXDRT.Report

    Dim Qry As String, z As Integer
    Dim i As Integer, j As Integer, NxtRow As Integer, InsertHere As Integer
    Dim MyDesc() As String, myTop() As String, myLeft() As String, myLen() As String
    Dim myHeight() As String, myFsize() As String, myFbold() As String, myAlign() As String
    Dim mySubLeft() As String, mySubLen() As String, mySubBold() As String, mySubAlign() As String
    Dim txtObj As TextObject
    
    Set mcLyr = New CRAXDRT.Application
    Set mcReport = mcLyr.NewReport
    
    'Access
    'Qry = "SELECT PRINTPREF.* From PRINTPREF " & _
          "WHERE (((PRINTPREF.FORMNAME)='" & FormName & "'));"
    'mySQL
    Qry = "SELECT PRINTPREF.* From PRINTPREF " & _
          "WHERE PRINTPREF.FORMNAME='" & FormName & "';"
    
    rsPref.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsPref.RecordCount > 0 Then
        rsPref.MoveFirst
        'Assign data to variables
        MyDesc = Split(rsPref!Description, ",")
        myTop = Split(rsPref!Top, ",")
        myLeft = Split(rsPref!Left, ",")
        myLen = Split(rsPref!Length, ",")
        myHeight = Split(rsPref!Height, ",")
        myFsize = Split(rsPref!FontSize, ",")
        myFbold = Split(rsPref!FontBold, ",")
        myAlign = Split(rsPref!Alignment, ",")
        InsertHere = IIf(rsPref!INSERTINDEX <> 0, rsPref!INSERTINDEX, -1)
        'Start Printing here
        If FormName = "INVOICE" Or FormName = "REQUISITION" Then
            mcReport.PaperSize = crPaperEnvelopeB6
        ElseIf FormName = "JOB ORDER" Then
            mcReport.PaperSize = crDefaultPaperSize
        Else
            mcReport.PaperSize = crDefaultPaperSize
        End If
        mcReport.TopMargin = 0
        mcReport.BottomMargin = 0
        mcReport.Sections(1).Height = 0
        mcReport.Sections(2).Height = 0
        With mcReport.Sections(3)
            .Height = rsPref!FORMHEIGHT
            For i = LBound(MyDesc) To UBound(MyDesc)
                If i = InsertHere Then
                    NxtRow = rsPref!LOOPSTARTTOP
                    mySubLeft = Split(rsPref!SUBLEFT, ",")
                    mySubLen = Split(rsPref!SUBLEN, ",")
                    mySubBold = Split(rsPref!SUBFBOLD, ",")
                    mySubAlign = Split(rsPref!SUBALIGN, ",")
                    If FormName = "Promissory Note" Then z = 10
                    For j = 1 To MyFlx.Rows - 1
                        For k = LBound(mySubLeft) To UBound(mySubLeft)
                            Set txtObj = .AddTextObject(MyFlx.TextMatrix(j, k + 1), mySubLeft(k), NxtRow)
                            txtObj.Width = Val(mySubLen(k)): txtObj.HorAlignment = Val(mySubAlign(k)): txtObj.Font.Name = "Arial": txtObj.Font.Size = Val(rsPref!LOOPFSIZE): txtObj.Height = Val(rsPref!LOOPHEIGHT)
                        Next k
                        If FormName = "Promissory Note" Then    'Control rows in Promi
                            NxtRow = NxtRow + rsPref!LOOPINTERVAL - z
                            z = IIf(z = 70, 10, z + 10)
                        Else
                            NxtRow = NxtRow + rsPref!LOOPINTERVAL
                        End If
                    Next j
                End If
                Set txtObj = .AddTextObject(MyVal(i), Val(myLeft(i)), Val(myTop(i)))
                txtObj.Width = Val(myLen(i)): txtObj.HorAlignment = Val(myAlign(i)): txtObj.Font.Name = "Arial": txtObj.Font.Size = Val(myFsize(i)): txtObj.Height = Val(myHeight(i)): txtObj.Font.Bold = Val(myFbold(i))
            Next i
        End With
        mcReport.Sections(4).Height = 0
        mcReport.Sections(5).Height = 0
        
        mcReport.PrintOut False, 1, , 1, 1
    Else
        MsgBox "No Printing Preference found!", vbCritical + vbOKOnly, "Unknown Form Name"
    End If
    rsPref.Close
End Sub


Function DTCpy(MyQry As String)
    Dim rsPrn As New ADODB.Recordset
    Dim myCnt As Integer, Tmp As String
    
    With rsPrn
        .Open MyQry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            myCnt = IIf(IsNull(.Fields(0)), 0, .Fields(0)) + 1
            .Fields(0) = myCnt
            .Update
        End If
        .Close
    End With
    If myCnt = 1 Then
        DTCpy = Format(Now, "mm/dd/yy hh:mm ampm")
    Else
        If myCnt = 2 Then
            Tmp = "nd"
        ElseIf myCnt = 3 Then
            Tmp = "rd"
        Else
            Tmp = "th"
        End If
        DTCpy = Format(Now, "mm/dd/yy hh:mm ampm") & " " & myCnt & Tmp & " copy"
    End If
End Function

Public Function RepAlign(aStr As String) As Integer
    Select Case aStr
        Case "crLeftAlign"
            RepAlign = 3
        Case "crHorCenterAlign"
            RepAlign = 1
        Case "crRightAlign"
            RepAlign = 4
    End Select
End Function

Public Function PureNum(nStr As String) As Single
    Dim Tmp As String
    Dim j As Integer
    
    Tmp = nStr
    If InStr(Tmp, "P") <> 0 Then
        Tmp = Right$(Tmp, Len(Tmp) - 1)
    End If
    If Tmp = "" Then
        PureNum = 0
        Exit Function
    End If
    Do While Not InStr(Tmp, ",") = 0
        j = InStr(Tmp, ",")
        Tmp = Left$(Tmp, j - 1) & Mid$(Tmp, j + 1, Len(Tmp) - j)
    Loop
    
    PureNum = Tmp
End Function

Public Function CtrlNumExist(Qry As String) As Boolean
    Dim rsExist As New ADODB.Recordset
    
    With rsExist
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            CtrlNumExist = True
        Else
            CtrlNumExist = False
        End If
        .Close
    End With
End Function

Public Function GetUserName() As String
    Dim rsEmp As New ADODB.Recordset
    Dim Qry As String
    
    '****Access
    'Qry = "SELECT [LASTNAME] & ', ' & [FIRSTNAME] & ' ' & Left$([MIDDLENAME],1) & '.' AS USERNAME, EMPLOYEES.EMPID " & _
          "From EMPLOYEES WHERE (((EMPLOYEES.EMPID)=" & UserID & "));"
    'mySQL
    Qry = "SELECT concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.'))) AS USERNAME " & _
          "From EMPLOYEES WHERE EMPLOYEES.EMPID=" & UserID & ";"
    
    With rsEmp
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetUserName = IIf(IsNull(.Fields(0)), "", .Fields(0))
        Else
            GetUserName = ""
        End If
        .Close
    End With
End Function

Public Function GetSettings(MyRemarks As String) As String
    Dim rsSettings As New ADODB.Recordset
    Dim Qry As String
    '****Access
    'Qry = "SELECT SETTINGS.VALUE, SETTINGS.DESCRIPTION, SETTINGS.REMARKS From SETTINGS " & _
          "WHERE (((SETTINGS.DESCRIPTION)='EMP DEFAULT VALUE') AND ((SETTINGS.REMARKS)='" & MyRemarks & "'));"
    
    'mySQL
    Qry = "SELECT SETTINGS.VALUE, SETTINGS.DESCRIPTION, SETTINGS.REMARKS From SETTINGS " & _
          "WHERE SETTINGS.DESCRIPTION='EMP DEFAULT VALUE' AND SETTINGS.REMARKS='" & MyRemarks & "';"
    
    With rsSettings
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            GetSettings = IIf(IsNull(!Value), "", !Value)
        Else
            GetSettings = ""
        End If
        .Close
    End With
End Function

Public Function GetFieldValue(MyQ As String) As Variant
    Dim rsAll As New ADODB.Recordset
    
    With rsAll
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        .Open MyQ, LYRcon, adOpenKeyset, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveFirst
            GetFieldValue = IIf(IsNull(.Fields(0)), "", .Fields(0))
        Else
            GetFieldValue = ""
        End If
        .Close
    End With
        
End Function

Public Function GetFieldValueMain(MyQ As String) As Variant
    Dim rsAll As New ADODB.Recordset
    
    With rsAll
        .Open MyQ, MainCon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetFieldValueMain = IIf(IsNull(.Fields(0)), "", .Fields(0))
        Else
            GetFieldValueMain = ""
        End If
        .Close
    End With
        
End Function

Public Sub LoadUserAccess()
    Dim rsAcc As New ADODB.Recordset
    Dim Qry As String
    
    Qry = "SELECT useraccess.EMPID, menu.MENUID, menu.MENUNAME, UserAccess.ACCESSTYPE From UserAccess " & _
          "INNER JOIN menu ON (useraccess.MENUID = menu.MENUID) Where useraccess.EMPID = " & UserID & " AND useraccess.ACCESSTYPE <> 3 Order By UserAccess.MENUID"
    rsAcc.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsAcc.RecordCount > 0 Then
        rsAcc.MoveFirst
        Do While Not rsAcc.EOF
            With Warehouse
              Select Case rsAcc!MENUNAME
                Case "Employees"
                    .mnuEmp.Enabled = True
                Case "User Profile"
                    .mnuUser.Enabled = True
                Case "Customer"
                    .mnuCustomer.Enabled = True
                Case "Supplier"
                    .mnuSupplier.Enabled = True
                Case "Branch"
                    .mnuBranch.Enabled = True
                Case "P.O. to Supplier"
                    .mnuPOsupplier.Enabled = True
                Case "P.O. from Branch"
                    .mnuPOBranch.Enabled = True
                Case "Receiving Receipts"
                    .mnuReceiving.Enabled = True
                Case "Sales Return"
                    .mnuSalesReturn.Enabled = True
                Case "Stock Transfer"
                    .mnuStockTransfer.Enabled = True
                Case "Purchase Return"
                    .mnuPurchaseReturn.Enabled = True
                Case "Merchandise Inventory"
                    .mnuMerchandise.Enabled = True
                Case "Cardex"
                    .mnuCardex.Enabled = True
                Case "Cardex (RED)"
                    .mnuCardexRed.Enabled = True
                Case "Back Order"
                    .mnuBackOrder.Enabled = True
                Case "Price Adjustment"
                    .mnuPrice.Enabled = True
                Case "Parts Properties"
                    .mnuPartsProp.Enabled = True
                Case "Receipts"
                    .mnuReceipts.Enabled = True
                Case "Provisional"
                    .mnuPr.Enabled = True
                Case "Checks"
                    .mnuChecks.Enabled = True
                Case "Cashflow"
                    .mnuCashflow.Enabled = True
                Case "Adv. for Liquidation"
                    .mnuAFL.Enabled = True
                Case "Statement of Accounts"
                    .mnuStatement.Enabled = True
                Case "Aging of Accounts"
                    .mnuAging.Enabled = True
                Case "DMCM"
                    .mnuDMCM.Enabled = True
                Case "Supplier Invoice"
                    .mnuInvoice.Enabled = True
                Case "Petty Cash Vouchers"
                    .mnuPCV.Enabled = True
                Case "Cash Voucher"
                    .mnuCV.Enabled = True
                Case "Daily Sales"
                    .mnuDailySales.Enabled = True
                Case "Daily Cash"
                    .mnuDailyCash.Enabled = True
                Case "Daily Expenses"
                    .mnuDailyExpense.Enabled = True
                Case "Monthly Sales"
                    .mnuMonSales.Enabled = True
                Case "Monthly Collection"
                    .mnuMonColl.Enabled = True
                Case "Monthly Expenses"
                    .mnuMonExpense.Enabled = True
                Case "Receiving Receipt Summary"
                    .mnuReceivingSum.Enabled = True
                Case "Detailed Outgoing Summary "
                    .mnuTrasferSum.Enabled = True
                Case "Sales Report at Selling Price"
                    .mnuSaleSellPrice.Enabled = True
                Case "Sales Report at Cost Price"
                    .mnuSalesCostPrice.Enabled = True
                Case "Sales Return Summary"
                    .mnuSalRetSum.Enabled = True
                Case "Pruchase Return Summary"
                    .mnuPurchRetSum.Enabled = True
                Case "Inventory Adjustment Summary"
                    .mnuInvAdjSum.Enabled = True
                Case "Sales Distribution Report"
                    .mnuSalesDistrib.Enabled = True
                Case "Sales DB for Extract"
                    .mnuSalesDB.Enabled = True
                Case "Backup"
                    .mnuBackup.Enabled = True
                Case "Scripting"
                    .mnuScripting.Enabled = True
                Case "Printing Preferences"
                    .mnuPrintPref.Enabled = True
                Case "Options"
                    .mnuOptions.Enabled = True
                Case "Settings"
                    .mnuSettings.Enabled = True
'                Case "Branch Back Order P.O."
'                    .mnuBBOPO.Enabled = True
                Case "Branch Back Order Summary"
                    .mnuBBOsumm.Enabled = True
                Case "Supplier Back Order Summary"
                    .mnuSBOsumm.Enabled = True
                Case "Allocation Branch Back Order"
                    .mnuAllocBBO.Enabled = True
                Case "Inventory Adjustment"
                    .mnuInvAdj.Enabled = True
                Case "Branch Receiving Discrepancy"
                    .mnuBranchRecDis.Enabled = True
                Case "Monthly Merchandise Inventory"
                    .mnuMonthlyInv.Enabled = True
                Case "Branch Replenishment"
                    .mnuBReplenish.Enabled = True
                Case "Branch Replenishment"
                    .mnuBReplenish.Enabled = True
                Case "Branch Discrepancy Process"
                    .mnuSDisc.Enabled = True
                Case "Settlement of Discrepancy Adj."
                    .mnuSDiscAdj.Enabled = True
                Case "SRP Updater"
                    .mnuSRPupdater.Enabled = True
'                Case "Replenishment Stock Transfer"
'                    .mnuReplSTrans.Enabled = True
                Case "Parts Catalog"
                    .mnuPartsCatalog.Enabled = True
                Case "Wholesale P.O to Supplier"
                    .mnuWSPOsupplier.Enabled = True
                Case "Wholesale Receiving Receipt"
                    .mnuWSReceiving.Enabled = True
                Case "Wholesale Merchandise Inventory"
                    .mnuWSMerchandise.Enabled = True
                Case "Wholesale Cardex"
                    .mnuWSCardex.Enabled = True
                Case "Wholesale Cardex(RED)"
                    .mnuWSCardexRed.Enabled = True
                Case "Wholesale Inventory Adjustment"
                    .mnuWSInvAdj.Enabled = True
                Case "Retail Stock Transfer"
                    .mnuRStockTransfer.Enabled = True
                Case "Detailed Supplier Service Ratio"
                    .mnuDtldServRatio.Enabled = True
                Case "Inventory Audit"
                    .mnuInvAudit.Enabled = True
                Case "Intransit Summary"
                    .mnuIntrnstSumm.Enabled = True
                Case "Request For Payment"
                    .mnuRFP.Enabled = True
                Case "RFP Reconciliation"
                    .mnuRFPRecon.Enabled = True
                Case "Retail Inventory Inquiry"
                    .mnuRInvInq.Enabled = True
                Case "Adjustment Request"
                    .mnuAdjReq.Enabled = True
                Case "Cancellation of Transaction"
                    .mnuCancelTrans.Enabled = True
            End Select
            rsAcc.MoveNext
          End With
        Loop
        rsAcc.Close
    End If
    
End Sub

Public Sub UserAccess()
    Dim rsUA As New ADODB.Recordset
    Dim Qry As String, Ndx As Integer, IDx As Integer
    Dim mnuX As Menu
    
    'Disable all submenu
    '****Access
    'Qry = "SELECT USERACCESS.* From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & "));"
    
    'mySQL
    Qry = "SELECT USERACCESS.* From UserAccess WHERE USERACCESS.EMPID=" & UserID & ";"
    
    rsUA.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsUA.RecordCount > 0 Then
        rsUA.MoveFirst
        Do While Not rsUA.EOF
            Ndx = rsUA!MenuIndx
            IDx = rsUA!MENUID
            Select Case Ndx
                Case 8 To 12
                    For Each mnuX In MotoRep.mSubPI
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 13 To 17
                    For Each mnuX In MotoRep.mSubSP
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 18 To 20, 97, 102
                    If Ndx = 18 Then MotoRep.BtnMain(1).Enabled = True
                    For Each mnuX In MotoRep.subOrder
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 21 To 24
                    For Each mnuX In MotoRep.subReceive
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 25 To 29
                    If Ndx = 25 Then MotoRep.BtnMain(2).Enabled = True
                    For Each mnuX In MotoRep.subSales
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 30 To 34
                    If Ndx = 30 Then MotoRep.BtnMain(0).Enabled = True
                    For Each mnuX In MotoRep.SPsubStocks
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 35 To 37
                    For Each mnuX In MotoRep.SPsubADJ
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 38 To 43
                    For Each mnuX In MotoRep.SubSRVC
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 44 To 50
                    For Each mnuX In MotoRep.mSubTR
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 51 To 56
                    For Each mnuX In MotoRep.subAdj
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 57 To 60
                    For Each mnuX In MotoRep.subDisb
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 61 To 63, 98, 99
                    If Ndx = 62 Then MotoRep.BtnMain(5).Enabled = True
                    For Each mnuX In MotoRep.subCashFlow
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 64 To 68, 104
                    If Ndx = 67 Then MotoRep.BtnMain(3).Enabled = True
                    For Each mnuX In MotoRep.subAR
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 69 To 72
                    For Each mnuX In MotoRep.mSubRep
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 73 To 75
                    If Ndx = 73 Then MotoRep.BtnMain(4).Enabled = True
                    For Each mnuX In MotoRep.subRptDailySales
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 76 To 81, 100
                    For Each mnuX In MotoRep.subRptSP
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 82 To 84
                    For Each mnuX In MotoRep.subRptSrv
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 85 To 89, 101
                    For Each mnuX In MotoRep.SubRptMonthly
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 90 To 93
                    For Each mnuX In MotoRep.mSubTools
                        If mnuX.Index = IDx Then mnuX.Enabled = True
                    Next
                Case 94
                    MotoRep.AudTools(1).Enabled = True
                Case 95
                    MotoRep.mSetPrn.Enabled = True
                Case 96
                    MotoRep.mSetOpt.Enabled = True
            End Select
            rsUA.MoveNext
        Loop
    End If
End Sub

Public Function ComputerName() As String
    Dim Cn As String

    Cn = Space(25)
    
    GetComputerName Cn, 25
    
    ComputerName = Left$(RTrim(Cn), Len(RTrim(Cn)) - 1)
End Function

Public Sub SBranch()
    Dim rsBrn As New ADODB.Recordset
    Dim mBrc As String
    '
    mBrc = "SELECT date_format(curdate(), '%M %d, %Y') as serTime , Branch.BRANCHNAME, Branch.branchid, Branch.Address, Branch.CONTACTNOS, Branch.manager, Branch.BRANCHCODE From Branch Where Branch.SYSTEMBRANCH = 1;"
    With rsBrn
        .Open mBrc, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            ServerTime = !sertime
            SysBranchID = Val(!branchID)
            SysBranch = !branchName
            SysAddress = !Address
            SysBranchOIC = !MANAGER
            sysBranchCode = IIf(IsNull(!BranchCode), "", !BranchCode)
            SysTel = IIf(IsNull(!CONTACTNOS), "", !CONTACTNOS)
        End If
        .Close
    End With
    
End Sub
Public Function GetFieldValueMU(MyQ As String) As Variant
    Dim rsBrch As New ADODB.Recordset
    
    With rsBrch
        .Open MyQ, conUlas, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetFieldValueMU = IIf(IsNull(.Fields(0)), "", .Fields(0))
        Else
            GetFieldValueMU = ""
        End If
        .Close
    End With
        
End Function

Public Function GetFieldValueVAN(MyQ As String) As Variant
    Dim rsBrch As New ADODB.Recordset

    With rsBrch
        .Open MyQ, conVAN1, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetFieldValueVAN = IIf(IsNull(.Fields(0)), "", .Fields(0))
        Else
            GetFieldValueVAN = ""
        End If
        .Close
    End With
        
End Function

Public Function GetFieldValueMRC2(MyQ As String) As Variant
    Dim rsBrch As New ADODB.Recordset

    With rsBrch
        .Open MyQ, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetFieldValueMRC2 = IIf(IsNull(.Fields(0)), "", .Fields(0))
        Else
            GetFieldValueMRC2 = ""
        End If
        .Close
    End With
        
End Function

Public Function GetFieldValueYM(MyQ As String) As Variant
    Dim rsBrch As New ADODB.Recordset
    
    With rsBrch
        .Open MyQ, conYM, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetFieldValueYM = IIf(IsNull(.Fields(0)), "", .Fields(0))
        Else
            GetFieldValueYM = ""
        End If
        .Close
    End With
        
End Function

Public Function GetFieldValueMD(MyQ As String) As Variant
    Dim rsBrch As New ADODB.Recordset
    
    With rsBrch
        .Open MyQ, conToril, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetFieldValueMD = IIf(IsNull(.Fields(0)), "", .Fields(0))
        Else
            GetFieldValueMD = ""
        End If
        .Close
    End With
        
End Function

Public Function GetSalesQtyBranch(MyID As Long, myCON As String)
    Dim rsQty As New ADODB.Recordset
    Dim MyQ As String
    Dim MySales As Long, MyRet As Long
    
    
    'If myCON = "MRC2" Then
        'MyQ = "SELECT SUM(`6msq`.OUT) AS 6mSalesQty From 6msq Where 6msq.SPAREPARTID=" & MyID & " AND 6msq.TRANSDATE >='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' AND 6msq.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
    'Else
        MyQ = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
              "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
    'End If
    With rsQty
        If myCON = "VAN1" Then
            .Open MyQ, conVAN1, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "VAN2" Then
            .Open MyQ, conVAN2, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "MRC2" Then
            .Open MyQ, conMRC2, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "MRC1" Then
            .Open MyQ, conMRC1, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "ULAS" Then
            .Open MyQ, conUlas, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "TORIL" Then
            .Open MyQ, conToril, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "YAMAHA" Then
            .Open MyQ, conYM, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "PANABO" Then
            .Open MyQ, conPanabo, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "POLOMOLOK" Then
            .Open MyQ, conPol, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "GENSAN" Then
            .Open MyQ, conGenSan, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "GANTE" Then
            .Open MyQ, conMRC2, adOpenKeyset, adLockOptimistic
         ElseIf myCON = "MARBEL" Then
            .Open MyQ, conYmarbel, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "STOTOMAS" Then
            .Open MyQ, conStotomas, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "SAMAL" Then
            .Open MyQ, conSamal, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "VAN3" Then
            .Open MyQ, conVan3, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "WSALE" Then
            .Open MyQ, conWsale, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "CABANTIAN" Then
            .Open MyQ, conCabantian, adOpenKeyset, adLockOptimistic
        End If
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    MyQ = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueBranch(MyQ, myCON))
    GetSalesQtyBranch = MySales - MyRet
End Function



Public Function GetSalesQtyULAS(MyID As Long)
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, conUlas, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueBranch(Qry, "ULAS"))
    GetSalesQtyULAS = MySales - MyRet
End Function

Public Function GetSalesQtyBrchYM(MyID As Long)
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, BrchConYM, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueYM(Qry))
    GetSalesQtyBrchYM = MySales - MyRet
End Function

Public Function GetSalesQtyBrchMD(MyID As Long)
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, BrchConMD, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueMD(Qry))
    GetSalesQtyBrchMD = MySales - MyRet
End Function


Public Function GetSalesQtyBrchVAN(MyID As Long)
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, BrchConVAN, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueVAN(Qry))
    GetSalesQtyBrchVAN = MySales - MyRet
End Function


Public Function GetSalesQtyBrchMRC2(MyID As Long)
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, BrnConMRC2, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueMRC2(Qry))
    GetSalesQtyBrchMRC2 = MySales - MyRet
End Function


Public Function GetDueDate(InvDate As Date, Term As Integer) As Date
    Dim mDay As Integer, TmpDue As Date, D30nxtM As Date, D45nxtM As Date
    
    
    If Term < 30 Then
        GetDueDate = DateAdd("d", Term, InvDate)
        Exit Function
    End If
    mDay = Day(InvDate)
    If mDay <= 15 Then
        TmpDue = DateSerial(Year(DateAdd("M", 2, InvDate)), Month(DateAdd("M", 2, InvDate)), 1) '1st Day of InvDate+2 months
        D30nxtM = DateAdd("D", -1, TmpDue)
        Select Case Term
            Case 30 To 35
                GetDueDate = D30nxtM
            Case 36 To 45
                GetDueDate = DateSerial(Year(DateAdd("M", 2, InvDate)), Month(DateAdd("M", 2, InvDate)), 15)
            Case 46 To 60
                TmpDue = DateSerial(Year(DateAdd("M", 2, D30nxtM)), Month(DateAdd("M", 2, D30nxtM)), 1) '1st Day of InvDate+2 months
                GetDueDate = DateAdd("D", -1, TmpDue)
            Case 61 To 75
                GetDueDate = DateSerial(Year(DateAdd("M", 3, InvDate)), Month(DateAdd("M", 3, InvDate)), 15)
            Case 76 To 90
                TmpDue = DateSerial(Year(DateAdd("M", 3, D30nxtM)), Month(DateAdd("M", 3, D30nxtM)), 1) '1st Day of InvDate+3 months
                GetDueDate = DateAdd("D", -1, TmpDue)
            Case 91 To 105
                GetDueDate = DateSerial(Year(DateAdd("M", 4, InvDate)), Month(DateAdd("M", 4, InvDate)), 15)
            Case 106 To 120
                TmpDue = DateSerial(Year(DateAdd("M", 4, D30nxtM)), Month(DateAdd("M", 4, D30nxtM)), 1) '1st Day of InvDate+4 months
                GetDueDate = DateAdd("D", -1, TmpDue)
            Case 121 To 135
                GetDueDate = DateSerial(Year(DateAdd("M", 5, InvDate)), Month(DateAdd("M", 5, InvDate)), 15)
            Case 136 To 150
                TmpDue = DateSerial(Year(DateAdd("M", 5, D30nxtM)), Month(DateAdd("M", 5, D30nxtM)), 1) '1st Day of InvDate+5 months
                GetDueDate = DateAdd("D", -1, TmpDue)
            Case 151 To 165
                GetDueDate = DateSerial(Year(DateAdd("M", 6, InvDate)), Month(DateAdd("M", 6, InvDate)), 15)
            Case 166 To 180
                TmpDue = DateSerial(Year(DateAdd("M", 6, D30nxtM)), Month(DateAdd("M", 6, D30nxtM)), 1) '1st Day of InvDate+6 months
                GetDueDate = DateAdd("D", -1, TmpDue)
            Case Else
                GetDueDate = DateAdd("D", Term, InvDate)
        End Select
    Else
        D30nxtM = DateSerial(Year(DateAdd("M", 2, InvDate)), Month(DateAdd("M", 2, InvDate)), 15) '15th Day of InvDate+2 months
        TmpDue = DateSerial(Year(DateAdd("M", 3, InvDate)), Month(DateAdd("M", 3, InvDate)), 1) '1st Day of InvDate+3 months
        D45nxtM = DateAdd("D", -1, TmpDue)  '30th Day of InvDate+2 months
        Select Case Term
            Case 30
                GetDueDate = D30nxtM
            Case 36 To 45
                GetDueDate = D45nxtM
            Case 46 To 60
                GetDueDate = DateAdd("M", 1, D30nxtM)
            Case 61 To 75
                TmpDue = DateSerial(Year(DateAdd("M", 2, D45nxtM)), Month(DateAdd("M", 2, D45nxtM)), 1) '1st Day of D45nxtM+2 months
                GetDueDate = DateAdd("D", -1, TmpDue)
            Case 76 To 90
                GetDueDate = DateAdd("M", 2, D30nxtM)
            Case 91 To 105
                TmpDue = DateSerial(Year(DateAdd("M", 3, D45nxtM)), Month(DateAdd("M", 3, D45nxtM)), 1) '1st Day of D45nxtM+3 months
                GetDueDate = DateAdd("D", -1, TmpDue)
            Case 106 To 120
                GetDueDate = DateAdd("M", 3, D30nxtM)
            Case 121 To 135
                TmpDue = DateSerial(Year(DateAdd("M", 4, D45nxtM)), Month(DateAdd("M", 4, D45nxtM)), 1) '1st Day of D45nxtM+4 months
                GetDueDate = DateAdd("D", -1, TmpDue)
            Case 136 To 150
                GetDueDate = DateAdd("M", 4, D30nxtM)
            Case 151 To 165
                TmpDue = DateSerial(Year(DateAdd("M", 5, D45nxtM)), Month(DateAdd("M", 5, D45nxtM)), 1) '1st Day of D45nxtM+5 months
                GetDueDate = DateAdd("D", -1, TmpDue)
            Case 166 To 180
                GetDueDate = DateAdd("M", 5, D30nxtM)
            Case Else
                GetDueDate = DateAdd("D", Term, InvDate)
        End Select
    End If
End Function


Public Function GetFieldValueBranch(MyQ As String, myCON As String) As Variant
    Dim rsAll As New ADODB.Recordset
    
    With rsAll
        If myCON = "VAN1" Then
            .Open MyQ, conVAN1, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "VAN2" Then
            .Open MyQ, conVAN2, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "MRC2" Then
            .Open MyQ, conMRC2, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "MRC1" Then
            .Open MyQ, conMRC1, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "ULAS" Then
            .Open MyQ, conUlas, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "TORIL" Then
            .Open MyQ, conToril, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "YAMAHA" Then
            .Open MyQ, conYM, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "PANABO" Then
            .Open MyQ, conPanabo, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "POLOMOLOK" Then
            .Open MyQ, conPol, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "GENSAN" Then
            .Open MyQ, conGenSan, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "GANTE" Then
            .Open MyQ, conMRC2, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "MARBEL" Then
            .Open MyQ, conYmarbel, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "STOTOMAS" Then
            .Open MyQ, conStotomas, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "SAMAL" Then
            .Open MyQ, conSamal, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "VAN3" Then
            .Open MyQ, conVan3, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "WSALE" Then
            .Open MyQ, conWsale, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "CABANTIAN" Then
            .Open MyQ, conCabantian, adOpenKeyset, adLockOptimistic
        End If
        
        
        If .RecordCount > 0 Then
            .MoveFirst
            GetFieldValueBranch = IIf(IsNull(.Fields(0)), "", .Fields(0))
        Else
            GetFieldValueBranch = ""
        End If
        .Close
    End With
        
End Function

Public Function ConnectMRC2() As Boolean 'mrc 2 retail

    On Error GoTo LocalErr
    If conMRC2.State = adStateOpen Then: conMRC2.Close
    conMRC2.CursorLocation = adUseClient
    conMRC2.ConnectionTimeout = 60
    conMRC2.Open "dsn=mrc2retail"
    ConnectMRC2 = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to MRC2 Database...."), vbExclamation
    ConnectMRC2 = False
End Function

Public Function ConnectVAN2() As Boolean

    On Error GoTo LocalErr
    If conVAN2.State = adStateOpen Then: conVAN2.Close
    conVAN2.CursorLocation = adUseClient
    conVAN2.ConnectionTimeout = 60
    conVAN2.Open "dsn=MRCVAN2"
    ConnectVAN2 = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to VAN2 Database...."), vbExclamation
    ConnectVAN2 = False
End Function

Public Function ConnectVAN1() As Boolean

    On Error GoTo LocalErr
    If conVAN1.State = adStateOpen Then: conVAN1.Close
    conVAN1.CursorLocation = adUseClient
    conVAN1.ConnectionTimeout = 60
    conVAN1.Open "dsn=SQLMTRVAN"
    ConnectVAN1 = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to VAN1 Database...."), vbExclamation
    ConnectVAN1 = False
End Function

Public Function ConnectMRC1() As Boolean

    On Error GoTo LocalErr
    If conMRC1.State = adStateOpen Then: conMRC1.Close
    conMRC1.CursorLocation = adUseClient
    conMRC1.ConnectionTimeout = 60
    conMRC1.Open "dsn=SQLMTRTGM"
    ConnectMRC1 = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to MRC1 Database...."), vbExclamation
    ConnectMRC1 = False
End Function

Public Function ConnectMRCGensan() As Boolean

    On Error GoTo LocalErr
    If conGenSan.State = adStateOpen Then: conGenSan.Close
    conGenSan.CursorLocation = adUseClient
    conGenSan.ConnectionTimeout = 60
    conGenSan.Open "dsn=MRCGENSAN"
    ConnectMRCGensan = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to MRC GENSAN Database...."), vbExclamation
    ConnectMRCGensan = False
End Function


Public Function ConnectUlas() As Boolean

    On Error GoTo LocalErr
    If conUlas.State = adStateOpen Then: conUlas.Close
    conUlas.CursorLocation = adUseClient
    conUlas.ConnectionTimeout = 60
    conUlas.Open "dsn=SQLMTRULAS"
    ConnectUlas = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to MOTO-R ULAS Database...."), vbExclamation
    ConnectUlas = False
End Function

Public Function ConnectToril() As Boolean

    On Error GoTo LocalErr
    If conToril.State = adStateOpen Then: conToril.Close
    conToril.CursorLocation = adUseClient
    conToril.ConnectionTimeout = 60
    conToril.Open "dsn=sqlmcpdepot"
    ConnectToril = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to MCPARTS DEPOT Database...."), vbExclamation
    ConnectToril = False
End Function

Public Function ConnectPanabo() As Boolean

    On Error GoTo LocalErr
    If conPanabo.State = adStateOpen Then: conPanabo.Close
    conPanabo.CursorLocation = adUseClient
    conPanabo.ConnectionTimeout = 60
    conPanabo.Open "dsn=MRCPANABO"
    ConnectPanabo = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to MRC PANABO Database...."), vbExclamation
    ConnectPanabo = False
End Function

Public Function ConnectHondaPol() As Boolean

    On Error GoTo LocalErr
    If conPol.State = adStateOpen Then: conPol.Close
    conPol.CursorLocation = adUseClient
    conPol.ConnectionTimeout = 60
    conPol.Open "dsn=mrcpolomolok"
    ConnectHondaPol = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to BINOS POLOMOLOK Database...."), vbExclamation
    ConnectHondaPol = False
End Function

Public Function ConnectYamaha() As Boolean

    On Error GoTo LocalErr
    If conYM.State = adStateOpen Then: conYM.Close
    conYM.CursorLocation = adUseClient
    conYM.ConnectionTimeout = 60
    conYM.Open "dsn=YAMAHA"
    ConnectYamaha = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to YAMAHA MOTORHUB Database...."), vbExclamation
    ConnectYamaha = False
End Function

Public Function ConnectYmarbel() As Boolean

    On Error GoTo LocalErr
    If conYmarbel.State = adStateOpen Then: conYmarbel.Close
    conYmarbel.CursorLocation = adUseClient
    conYmarbel.ConnectionTimeout = 60
    conYmarbel.Open "dsn=mrcmarbel"
    ConnectYmarbel = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to YAMAHA MOTORHUB MARBEL Database...."), vbExclamation
    ConnectYmarbel = False
End Function

Public Function ConnectStotomas() As Boolean

    On Error GoTo LocalErr
    If conStotomas.State = adStateOpen Then: conStotomas.Close
    conStotomas.CursorLocation = adUseClient
    conStotomas.ConnectionTimeout = 60
    conStotomas.Open "dsn=mrcstotomas"
    ConnectStotomas = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to BINOS STO TOMAS SPAREPARTS Database...."), vbExclamation
    ConnectStotomas = False
End Function

Public Function ConnectSamal() As Boolean

    On Error GoTo LocalErr
    If conSamal.State = adStateOpen Then: conSamal.Close
    conSamal.CursorLocation = adUseClient
    conSamal.ConnectionTimeout = 60
    conSamal.Open "dsn=mrcsamal"
    ConnectSamal = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to BINOS SAMAL SPAREPARTS Database...."), vbExclamation
    ConnectSamal = False
End Function

Public Function ConnectVan3() As Boolean

    On Error GoTo LocalErr
    If conVan3.State = adStateOpen Then: conVan3.Close
    conVan3.CursorLocation = adUseClient
    conVan3.ConnectionTimeout = 60
    conVan3.Open "dsn=mrcvan3"
    ConnectVan3 = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to MRC VAN3 Database...."), vbExclamation
    ConnectVan3 = False
End Function

Public Function ConnectWsale() As Boolean 'wholesale

    On Error GoTo LocalErr
    If conWsale.State = adStateOpen Then: conWsale.Close
    conWsale.CursorLocation = adUseClient
    conWsale.ConnectionTimeout = 60
    conWsale.Open "dsn=sqlmrc2tgm"
    ConnectWsale = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to MRC WHOLESALE Database...."), vbExclamation
    ConnectWsale = False
End Function

Public Function ConnectCabantian() As Boolean 'mrc cabantian

    On Error GoTo LocalErr
    If conCabantian.State = adStateOpen Then: conCabantian.Close
    conCabantian.CursorLocation = adUseClient
    conCabantian.ConnectionTimeout = 60
    conCabantian.Open "dsn=mrccabantian"
    ConnectCabantian = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to MRC CABANTIAN Database...."), vbExclamation
    ConnectCabantian = False
End Function

Public Function ConnectWhouse() As Boolean 'warehouse

    On Error GoTo LocalErr
    If conWhouse.State = adStateOpen Then: conWhouse.Close
    conWhouse.CursorLocation = adUseClient
    conWhouse.ConnectionTimeout = 60
    conWhouse.Open "dsn=mrcwarehouse"
    
    'conWhouse.Open "dsn=lyrmpsobrecarywh"
    ConnectWhouse = True
    Exit Function

LocalErr:
    MsgBox ("Error connecting to MRC WAREHOUSE Database...."), vbExclamation
    ConnectWhouse = False
End Function

Public Sub sp_get_Branch(objRs As ADODB.Recordset)
On Error GoTo er

    DoEvents
    Qry = "Select br.BRANCHID, br.BRANCHNAME, br.BRANCHCODE , ifnull(br.ADDRESS,''), ifnull(br.CONTACTNOS, '') From mrcwarehouse.branch as br Where branchcategory IN ('MRC','WHOLESALE','DEPOT','WAREHOUSE')"
    'Qry = "Select br.BRANCHID, br.BRANCHNAME, br.BRANCHCODE , ifnull(br.ADDRESS,''), ifnull(br.CONTACTNOS, '') From lyrmpsobrecarywh.branch as br Where branchcategory IN ('MRC','WHOLESALE','DEPOT','WAREHOUSE')"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
DoEvents
   Call modObjTools.ErrorMessage(Err.Description, "Error->sp_get_Branch()")

End Sub

Public Function GetEmployeeName() As String
    Dim rsEmp As New ADODB.Recordset
    Dim Qry As String
    
    '****Access
    'Qry = "SELECT [LASTNAME] & ', ' & [FIRSTNAME] & ' ' & Left$([MIDDLENAME],1) & '.' AS USERNAME, EMPLOYEES.EMPID " & _
          "From EMPLOYEES WHERE (((EMPLOYEES.EMPID)=" & UserID & "));"
    'mySQL
    Qry = "SELECT concat(EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.')),' ',EMPLOYEES.LASTNAME) AS EMPLOYEENAME " & _
          "From EMPLOYEES WHERE EMPLOYEES.EMPID=" & UserID & ";"
    
    With rsEmp
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetEmployeeName = IIf(IsNull(.Fields(0)), "", .Fields(0))
        Else
            GetEmployeeName = ""
        End If
        .Close
    End With
End Function

