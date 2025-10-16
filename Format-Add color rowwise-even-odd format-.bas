Attribute VB_Name = "Module1"
Sub ColorEvenOddRows_WithBorderOptions()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim rng As Range
    
    Set ws = ActiveSheet
    
    ' Use UsedRange for more reliable detection
    With ws.UsedRange
        lastRow = .Rows.Count + .Row - 1
        lastCol = .Columns.Count + .Column - 1
    End With
    
    ' Set the entire range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    ' Clear existing colors and borders
    rng.Interior.ColorIndex = xlNone
    rng.Borders.LineStyle = xlNone
    
    ' ===== Apply pure black borders =====
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0)
        .Weight = xlThin
    End With
    
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0)
        .Weight = xlThin
    End With
    
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0)
        .Weight = xlThin
    End With
    
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0)
        .Weight = xlThin
    End With
    
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0)
        .Weight = xlThin
    End With
    
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0)
        .Weight = xlThin
    End With
    
    ' ===== Apply alternating row colors =====
    For i = 1 To lastRow
        Dim rowRange As Range
        Set rowRange = ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol))
        
        If i Mod 2 = 0 Then
            rowRange.Interior.Color = RGB(173, 216, 230) ' Light blue for even rows
        Else
            rowRange.Interior.Color = RGB(255, 255, 255) ' White for odd rows
        End If
    Next i
    
End Sub

