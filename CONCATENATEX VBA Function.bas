
Function CONCATENATEX(rng As Range, Optional delimiter As String = ", ") As String
    Dim cell As Range
    Dim result As String
    
    For Each cell In rng
        If Trim(cell.Value) <> "" Then
            result = result & cell.Value & delimiter
        End If
    Next cell
    
    '????? delimiter ???? ????
    If Len(result) > 0 Then result = Left(result, Len(result) - Len(delimiter))
    
    CONCATENATEX = result
End Function


