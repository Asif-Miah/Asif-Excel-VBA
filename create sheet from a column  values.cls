
Sub CreateSheetsFromColumnA()
    Dim ws As Worksheet
    Dim cell As Range
    Dim newSheet As Worksheet
    Dim uniqueValues As Object
    Dim val As Variant
    Dim lastRow As Long
    Dim sheetName As String
    
    ' Use active sheet
    Set ws = ActiveSheet
    
    ' Create dictionary for unique values
    Set uniqueValues = CreateObject("Scripting.Dictionary")
    
    ' Find last row in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through column A and collect unique values
    For Each cell In ws.Range("A2:A" & lastRow)
        If Trim(cell.Value) <> "" Then
            val = Trim(cell.Value)
            If Not uniqueValues.exists(val) Then
                uniqueValues.Add val, 1
            End If
        End If
    Next cell
    
    ' Create sheets for each unique value
    For Each val In uniqueValues.keys
        sheetName = Left(val, 31) ' Sheet names max 31 chars
        
        ' Check if sheet already exists
        On Error Resume Next
        Set newSheet = ThisWorkbook.Sheets(sheetName)
        On Error GoTo 0
        
        ' If not exist, create new sheet
        If newSheet Is Nothing Then
            Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            newSheet.Name = sheetName
        End If
        
        ' Reset object for next loop
        Set newSheet = Nothing
    Next val
    
    MsgBox "Sheets created successfully!", vbInformation
End Sub

