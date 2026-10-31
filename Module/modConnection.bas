Attribute VB_Name = "modConnection"
Option Explicit
Dim ConString As String

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Private Function getActiveXConString() As String
On Error GoTo er
    Dim objConString As Object

    getActiveXConString = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Port=3306;Database=sp_acctpayinvcng;User=root;Password=laborday;Option=3;"
    ConString = getActiveXConString

Exit Function
er:
    MsgBox "ActiveX LYRconnection String Binding Error!", vbCritical, "Binding Error!"
    End
End Function

Public Sub DBConnection()
On Error GoTo er
    Set LYRcon = New ADODB.Connection
    LYRcon.CursorLocation = adUseClient
    LYRcon.LYRconnectionTimeout = 1200
    LYRcon.CommandTimeout = 1200
    LYRcon.LYRconnectionString = getActiveXConString
    LYRcon.Open
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->DBConnection()"
    End
End Sub

Public Function TestConnection() As Boolean
On Error GoTo er
    TestConnection = False
    If LYRcon.State = 0 Then
        Exit Function
    End If
    TestConnection = True
Exit Function
er:
    TestConnection = False
    MsgBox Err.Description, vbCritical, "LYRconNECTION ERROR!"

End Function

Public Function InnerDBTest() As Boolean
On Error GoTo er
    Dim tmpRs As New ADODB.Recordset
    Dim TestVal As Integer
    Dim Qry As String
    
    InnerDBTest = False
    If tmpRs.State <> 0 Then tmpRs.Close
    
    Qry = "Select 1 as TestVal"
    If LYRcon.State = 1 Then
        tmpRs.Open Qry, LYRcon, adOpenStatic
        If tmpRs.State = 1 Then
            If tmpRs.RecordCount > 0 Then
                InnerDBTest = True
            End If
        Else
            InnerDBTest = False
        End If
    End If
    If tmpRs.State <> 0 Then tmpRs.Close
Exit Function
er:
    MsgBox Err.Description, vbCritical, "Error->InnerDBTest()"
    InnerDBTest = False
End Function

