Attribute VB_Name = "Module1"
Function EnglishToBanglaNumber_New(ByVal EngNum As String) As String
    Dim i As Integer
    Dim ch As String
    Dim BanglaNum As String
    
    For i = 1 To Len(EngNum)
        ch = Mid(EngNum, i, 1)
        Select Case ch
            Case "0": BanglaNum = BanglaNum & ChrW(&H9E6)
            Case "1": BanglaNum = BanglaNum & ChrW(&H9E7)
            Case "2": BanglaNum = BanglaNum & ChrW(&H9E8)
            Case "3": BanglaNum = BanglaNum & ChrW(&H9E9)
            Case "4": BanglaNum = BanglaNum & ChrW(&H9EA)
            Case "5": BanglaNum = BanglaNum & ChrW(&H9EB)
            Case "6": BanglaNum = BanglaNum & ChrW(&H9EC)
            Case "7": BanglaNum = BanglaNum & ChrW(&H9ED)
            Case "8": BanglaNum = BanglaNum & ChrW(&H9EE)
            Case "9": BanglaNum = BanglaNum & ChrW(&H9EF)
            Case Else: BanglaNum = BanglaNum & ch
        End Select
    Next i

    EnglishToBanglaNumber_New = BanglaNum
End Function

