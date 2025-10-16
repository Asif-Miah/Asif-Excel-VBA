Attribute VB_Name = "Module2"
'========================================
' Custom Excel Function: WordCount
'========================================
Function WordCount(Text As String) As Long
    Dim arr() As String
    
    ' Handle blank cells
    If Trim(Text) = "" Then
        WordCount = 0
        Exit Function
    End If
    
    ' Split by space
    arr = Split(Trim(Text), " ")
    
    ' Return word count
    WordCount = UBound(arr) - LBound(arr) + 1
End Function

