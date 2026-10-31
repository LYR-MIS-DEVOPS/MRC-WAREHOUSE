Attribute VB_Name = "modTools"
Public Sub loadFlxRecord(rs As ADODB.Recordset, flx As Object, Optional lbl As Object = Nothing)
    
    With rs
        '.Open
        If .RecordCount = 0 Then
            flx.Clear
            flx.Rows = 2
            Exit Sub
        End If
    Set flx.DataSource = rs
    If Not (lbl Is Nothing) Then lbl.Caption = lbl.Caption & ": " & .RecordCount
    '.Close
    End With

End Sub
