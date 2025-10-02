VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cell As Range
    Dim v As String
    
    On Error GoTo SafeExit
    
    ' Only monitor column A
    If Intersect(Target, Me.Range("A2:A100000")) Is Nothing Then Exit Sub
    
    Application.EnableEvents = False
    
    For Each cell In Intersect(Target, Me.Range("A2:A100000"))
        If Not IsEmpty(cell.Value) Then
            v = Trim(CStr(cell.Value))
            v = Replace(v, " ", "")
            v = Replace(v, "-", "")
            
            ' Only numeric
            If IsNumeric(v) Then
                If Len(v) >= 10 Then
                    v = Right(v, 10)           ' Keep last 10 digits
                    cell.NumberFormat = "@"     ' Force Text
                    cell.Value = "+880" & v
                    cell.Errors.Item(xlNumberAsText).Ignore = True  ' Remove green mark
                    
                    ' Highlight valid cell with border and color
                    With cell
                        .Interior.Color = RGB(255, 255, 153) ' Light yellow
                        .Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Borders(xlEdgeTop).LineStyle = xlContinuous
                        .Borders(xlEdgeBottom).LineStyle = xlContinuous
                        .Borders(xlEdgeRight).LineStyle = xlContinuous
                        .Borders(xlInsideVertical).LineStyle = xlContinuous
                        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                    End With
                Else
                    ' Show error but keep value for editing
                    MsgBox "Number in " & cell.Address & " must be at least 10 digits.", vbExclamation
                End If
            Else
                ' Show error but keep value for editing
                MsgBox "Only numeric digits are allowed in " & cell.Address, vbExclamation
            End If
        End If
    Next cell
    
SafeExit:
    Application.EnableEvents = True
End Sub


