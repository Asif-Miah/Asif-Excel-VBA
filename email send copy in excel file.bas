Attribute VB_Name = "Module1"
Sub MergeMacros()
    ' Step 1: List Sent Emails to Separate Sheets By Recipient
    Call ListSentEmailsToSeparateSheetsByRecipient
    
    ' Step 2: Split Column B Into Separate Sheets
    Call SplitColumnBIntoSheets
End Sub

Sub ListSentEmailsToSeparateSheetsByRecipient()
    Dim OutlookApp As Object
    Dim OutlookNamespace As Object
    Dim SentFolder As Object
    Dim MailItem As Object
    Dim Recip As Object
    Dim i As Long, j As Long
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sheetName As String
    Dim ToName As String
    Dim BodyText As String
    Dim EmailAddress As String
    Dim ToList As String
    Dim row As Long
    Dim wsTarget As Worksheet

    Set wb = ThisWorkbook

    ' Setup Outlook
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")
    Set SentFolder = OutlookNamespace.GetDefaultFolder(5) ' olFolderSentMail
    SentFolder.Items.Sort "[SentOn]", True

    ' Loop through sent emails
    For i = 1 To SentFolder.Items.Count
        Set MailItem = SentFolder.Items(i)

        If MailItem.Class = 43 Then ' olMail
            ' Loop through "To" recipients
            For j = 1 To MailItem.Recipients.Count
                Set Recip = MailItem.Recipients(j)

                If Recip.Type = 1 Then ' 1 = To
                    On Error Resume Next
                    If Not Recip.AddressEntry Is Nothing Then
                        EmailAddress = Recip.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                        If EmailAddress = "" Then EmailAddress = Recip.Address
                    Else
                        EmailAddress = Recip.Address
                    End If
                    On Error GoTo 0

                    If EmailAddress <> "" Then
                        ' Clean sheet name
                        sheetName = Left(CleanSheetName(EmailAddress), 31)

                        ' Check if sheet exists, if not, create it
                        On Error Resume Next
                        Set wsTarget = wb.Sheets(sheetName)
                        On Error GoTo 0

                        If wsTarget Is Nothing Then
                            Set wsTarget = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
                            wsTarget.Name = sheetName
                            wsTarget.Range("A1:E1").Value = Array("Subject", "To", "CC", "Sent On", "Body")
                        End If

                        ' Add data
                        row = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).row + 1
                        wsTarget.Cells(row, 1).Value = MailItem.Subject
                        wsTarget.Cells(row, 2).Value = MailItem.To
                        wsTarget.Cells(row, 3).Value = MailItem.CC
                        wsTarget.Cells(row, 4).Value = MailItem.SentOn
                        wsTarget.Cells(row, 5).Value = Left(MailItem.Body, 1000)
                    End If
                End If
            Next j
        End If
    Next i

    MsgBox "Emails categorized by recipient into separate sheets!"
End Sub

Sub SplitColumnBIntoSheets()
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim rngData As Range
    Dim dict As Object
    Dim key As Variant
    Dim i As Long
    Dim lastRow As Long
    Dim rowNum As Long
    
    ' Set the source worksheet (currently active sheet)
    Set wsSource = ActiveSheet
    
    ' Error handling if no worksheet is selected
    If wsSource Is Nothing Then
        MsgBox "No worksheet selected!", vbExclamation
        Exit Sub
    End If
    
    ' Find the last row in Column B
    lastRow = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).row
    If lastRow < 2 Then
        MsgBox "No data found in Column B!", vbExclamation
        Exit Sub
    End If
    
    ' Set the data range (assuming headers are in Row 1)
    Set rngData = wsSource.Range("A1").CurrentRegion
    
    ' Create a dictionary to store unique values and their corresponding rows
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Collect unique values and store row numbers
    For i = 2 To lastRow
        key = wsSource.Cells(i, 2).Value ' Column B values
        If key <> "" Then ' Skip empty cells
            If Not dict.Exists(key) Then
                dict.Add key, New Collection
            End If
            dict(key).Add i ' Store row number
        End If
    Next i
    
    ' Check if we found any unique values
    If dict.Count = 0 Then
        MsgBox "No unique values found in Column B!", vbExclamation
        Exit Sub
    End If
    
    ' Disable screen updates and alerts for efficiency
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Process each unique value
    For Each key In dict.Keys
        ' Clean sheet name (remove invalid chars)
        Dim sheetName As String
        sheetName = CleanSheetName(CStr(key))
        
        ' Check if the sheet exists, create if not
        On Error Resume Next
        Set wsDest = ThisWorkbook.Worksheets(sheetName)
        On Error GoTo 0
        
        If wsDest Is Nothing Then
            ' Create a new sheet if it doesn't exist
            Set wsDest = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            wsDest.Name = sheetName
        Else
            ' Clear existing data in the sheet
            wsDest.Cells.Clear
        End If
        
        ' Copy headers from source to destination
        rngData.Rows(1).Copy wsDest.Range("A1")
        
        ' Copy matching rows to the new sheet using For loop
        rowNum = 2 ' Start copying at row 2 (after headers)
        For i = 1 To dict(key).Count
            wsSource.Rows(dict(key)(i)).Copy wsDest.Rows(rowNum)
            rowNum = rowNum + 1
        Next i
        
        Set wsDest = Nothing ' Clear the destination worksheet variable
    Next key
    
    ' Restore screen updates and alerts
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Split completed! Created " & dict.Count & " new sheets.", vbInformation

Call ListAllUnreadEmails_AllFolders
End Sub

' Helper function to clean sheet names (Excel has a 31-character limit & invalid chars)
Function CleanSheetName(str As String) As String
    Dim invalidChars As String
    invalidChars = "\/*?[]:" ' Invalid characters for Excel sheet names
    Dim i As Integer
    
    ' Remove invalid characters
    For i = 1 To Len(invalidChars)
        str = Replace(str, Mid(invalidChars, i, 1), "")
    Next i
    
    ' Trim to 31 characters if necessary
    If Len(str) > 31 Then
        str = Left(str, 31)
    End If
    
    CleanSheetName = str
End Function




Sub ListAllUnreadEmails_AllFolders()
    Dim OutlookApp As Object
    Dim OutlookNamespace As Object
    Dim Folder As Object
    Dim MailItem As Object
    Dim wb As Workbook, wsReport As Worksheet
    Dim row As Long
    Dim senderInfo As String
    Dim ReportSheetName As String

    Set wb = ThisWorkbook
    ReportSheetName = "Unread Report"

    ' Create or clear the report sheet
    On Error Resume Next
    Set wsReport = wb.Sheets(ReportSheetName)
    If Not wsReport Is Nothing Then
        wsReport.Cells.Clear
    Else
        Set wsReport = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsReport.Name = ReportSheetName
    End If
    On Error GoTo 0

    ' Set up headers
    With wsReport
        .Range("A1:E1").Value = Array("Folder", "Subject", "Sender (Name <Email>)", "Received Time", "Body Preview")
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(255, 255, 204)
        .Rows(1).HorizontalAlignment = xlCenter
    End With

    row = 2
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")

    ' Start scanning all folders
    Call ScanFoldersForUnread(OutlookNamespace.Folders, wsReport, row)

    ' Final formatting
    With wsReport
        .Columns("A:E").AutoFit
        .Columns("E:E").WrapText = True
        .Rows(1).Borders.LineStyle = xlContinuous
        .Activate
        .Range("A2").Select
        ActiveWindow.FreezePanes = True
    End With

    MsgBox "Unread email report generated successfully!"
End Sub

Sub ScanFoldersForUnread(Folders As Object, ws As Worksheet, ByRef row As Long)
    Dim Folder As Object
    Dim MailItem As Object
    Dim i As Long
    Dim senderInfo As String

    For Each Folder In Folders
        ' Check if folder contains mail items
        If Folder.DefaultItemType = 0 Then ' 0 = Mail item
            On Error Resume Next
            For i = 1 To Folder.Items.Count
                If Folder.Items(i).Class = 43 Then ' olMail
                    Set MailItem = Folder.Items(i)
                    If MailItem.UnRead = True Then
                        senderInfo = MailItem.SenderName & " <" & MailItem.SenderEmailAddress & ">"
                        ws.Cells(row, 1).Value = Folder.Name
                        ws.Cells(row, 2).Value = MailItem.Subject
                        ws.Cells(row, 3).Value = senderInfo
                        ws.Cells(row, 4).Value = MailItem.ReceivedTime
                        ws.Cells(row, 5).Value = Left(MailItem.Body, 1000)
                        row = row + 1
                    End If
                End If
            Next i
            On Error GoTo 0
        End If

        ' Recursively scan subfolders
        If Folder.Folders.Count > 0 Then
            Call ScanFoldersForUnread(Folder.Folders, ws, row)
        End If
    Next Folder
End Sub



