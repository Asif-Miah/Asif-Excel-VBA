Attribute VB_Name = "Module1"
Function RepeatValues(Value As Variant, Count As Long) As Variant
    Dim result() As Variant
    Dim i As Long
    
    If Count <= 0 Then
        RepeatValues = Array()
        Exit Function
    End If
    
    ReDim result(1 To Count, 1 To 1)
    
    For i = 1 To Count
        result(i, 1) = Value
    Next i
    
    RepeatValues = result
End Function

