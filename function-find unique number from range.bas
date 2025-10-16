Attribute VB_Name = "Module2"
Function SmartMode(rng As Range) As String
    Dim dict As Object
    Dim cell As Range
    Dim val As Variant
    Dim maxCount As Long
    Dim key As Variant
    Dim modeList As String
    
    On Error Resume Next
    Set dict = CreateObject("Scripting.Dictionary")
    On Error GoTo 0
    
    ' Loop through each cell
    For Each cell In rng
        If Not IsError(cell.Value) And Not IsEmpty(cell.Value) Then
            val = Trim(CStr(cell.Value))
            If val <> "" Then
                If dict.Exists(val) Then
                    dict(val) = dict(val) + 1
                Else
                    dict.Add val, 1
                End If
            End If
        End If
    Next cell
    
    ' If only one cell was passed
    If dict.Count = 0 Then
        SmartMode = "No data"
        Exit Function
    End If
    
    ' Find highest frequency
    maxCount = Application.Max(dict.Items)
    
    ' Collect all items with that frequency
    For Each key In dict.Keys
        If dict(key) = maxCount Then
            If modeList = "" Then
                modeList = key
            Else
                modeList = modeList & ", " & key
            End If
        End If
    Next key
    
    SmartMode = modeList
End Function

