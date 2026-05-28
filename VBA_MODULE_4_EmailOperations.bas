Attribute VB_Name = "Module4_EmailOperations"
' ============================================================================
' BHFL CHANNEL OPERATIONS - DSA AUTOMATION
' Email Operations Module
' ============================================================================
' Purpose: Handle Outlook email creation and sending with attachments
' ============================================================================

Option Explicit

' ============================================================================
' Send email with attachment using Outlook
' ============================================================================
Sub SendEmailWithAttachment(dsaName As String, emailAddress As String, attachmentPath As String)
    On Error GoTo ErrorHandler
    
    Dim outlookApp As Object
    Dim mailItem As Object
    Dim fso As Object
    
    ' Create Outlook application object
    Set outlookApp = CreateObject("Outlook.Application")
    
    ' Validate file exists
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(attachmentPath) Then
        gErrorCount = gErrorCount + 1
        Call LogEmailError(dsaName, emailAddress, "Attachment file not found: " & attachmentPath)
        Exit Sub
    End If
    
    ' Create mail item
    Set mailItem = outlookApp.CreateItem(0) ' 0 = olMailItem
    
    ' Set email properties
    With mailItem
        .To = emailAddress
        .Subject = "DSA Update – " & dsaName
        .Body = GetEmailBody()
        
        ' Add attachment
        .Attachments.Add attachmentPath
        
        ' Display email (not send for demo)
        .Display
        ' For production use: .Send
    End With
    
    ' Log email generation
    Call LogEmailGeneration(dsaName, emailAddress, attachmentPath)
    
    ' Clean up
    Set mailItem = Nothing
    Set outlookApp = Nothing
    Set fso = Nothing
    
    Exit Sub
    
ErrorHandler:
    gErrorCount = gErrorCount + 1
    Call LogEmailError(dsaName, emailAddress, Err.Description)
    On Error Resume Next
    Set mailItem = Nothing
    Set outlookApp = Nothing
    Set fso = Nothing
End Sub

' ============================================================================
' Get email body template
' ============================================================================
Function GetEmailBody() As String
    Dim body As String
    
    body = "Hi Team," & vbCrLf & vbCrLf & _
           "Please find attached DSA update." & vbCrLf & vbCrLf & _
           "Regards," & vbCrLf & _
           "Operations Team"
    
    GetEmailBody = body
End Function

' ============================================================================
' Prepare and send all emails for DSA files
' ============================================================================
Sub PrepareAndSendEmails(outputPath As String)
    On Error GoTo ErrorHandler
    
    Dim mailMapData As Variant
    Dim fileList As Collection
    Dim i As Integer
    Dim fileName As String
    Dim dsaName As String
    Dim emailAddress As String
    Dim filePath As String
    
    ' Read MailMap
    mailMapData = ReadMailMapSheet()
    
    ' Get list of files in output folder
    Set fileList = GetFilesInFolder(outputPath)
    
    ' Process each file
    For i = 1 To fileList.Count
        fileName = fileList(i)
        dsaName = ExtractDSANameFromFileName(fileName)
        filePath = outputPath & fileName
        
        ' Get email from MailMap
        emailAddress = GetEmailForDSA(mailMapData, dsaName)
        
        ' Validate email
        If emailAddress = "" Or InStr(emailAddress, "@") = 0 Then
            gErrorCount = gErrorCount + 1
            Call LogEmailError(dsaName, emailAddress, "Missing or Invalid Email")
            GoTo NextEmail
        End If
        
        ' Send email
        Call SendEmailWithAttachment(dsaName, emailAddress, filePath)
        gMailsGenerated = gMailsGenerated + 1
        
NextEmail:
    Next i
    
    Exit Sub
    
ErrorHandler:
    gErrorCount = gErrorCount + 1
    Call LogError("PrepareAndSendEmails", Err.Description)
End Sub

' ============================================================================
' Log email generation
' ============================================================================
Sub LogEmailGeneration(dsaName As String, emailAddress As String, attachmentFile As String)
    On Error Resume Next
    
    Dim logsSheet As Worksheet
    Dim lastRow As Integer
    Dim fileName As String
    
    ' Extract file name from path
    fileName = Right(attachmentFile, InStr(StrReverse(attachmentFile), "\") - 1)
    
    ' Get logs sheet
    On Error Resume Next
    Set logsSheet = ThisWorkbook.Sheets("Logs")
    On Error GoTo 0
    
    If logsSheet Is Nothing Then
        Call CreateLogsSheet
        Set logsSheet = ThisWorkbook.Sheets("Logs")
    End If
    
    ' Add log entry
    lastRow = logsSheet.Cells(Rows.Count, 1).End(xlUp).Row + 1
    logsSheet.Cells(lastRow, 1).Value = Now()
    logsSheet.Cells(lastRow, 2).Value = dsaName
    logsSheet.Cells(lastRow, 3).Value = fileName
    logsSheet.Cells(lastRow, 4).Value = emailAddress
    logsSheet.Cells(lastRow, 5).Value = "SUCCESS"
    logsSheet.Cells(lastRow, 6).Value = "Email prepared and displayed"
    
End Sub

' ============================================================================
' Check if Outlook is installed
' ============================================================================
Function IsOutlookInstalled() As Boolean
    On Error GoTo ErrorHandler
    
    Dim outlookApp As Object
    Set outlookApp = CreateObject("Outlook.Application")
    Set outlookApp = Nothing
    
    IsOutlookInstalled = True
    Exit Function
    
ErrorHandler:
    IsOutlookInstalled = False
End Function

' ============================================================================
' Send email with multiple recipients
' ============================================================================
Sub SendEmailMultipleRecipients(dsaName As String, emailAddresses As String, attachmentPath As String)
    On Error GoTo ErrorHandler
    
    Dim outlookApp As Object
    Dim mailItem As Object
    Dim fso As Object
    
    ' Create Outlook application object
    Set outlookApp = CreateObject("Outlook.Application")
    
    ' Validate file exists
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(attachmentPath) Then
        gErrorCount = gErrorCount + 1
        Call LogError("SendEmailMultipleRecipients", "Attachment file not found")
        Exit Sub
    End If
    
    ' Create mail item
    Set mailItem = outlookApp.CreateItem(0)
    
    ' Set email properties
    With mailItem
        .To = emailAddresses
        .Subject = "DSA Update – " & dsaName
        .Body = GetEmailBody()
        .Attachments.Add attachmentPath
        .Display
    End With
    
    Set mailItem = Nothing
    Set outlookApp = Nothing
    Set fso = Nothing
    
    Exit Sub
    
ErrorHandler:
    gErrorCount = gErrorCount + 1
    Call LogError("SendEmailMultipleRecipients", Err.Description)
    On Error Resume Next
    Set mailItem = Nothing
    Set outlookApp = Nothing
    Set fso = Nothing
End Sub

' ============================================================================
' Send email with CC
' ============================================================================
Sub SendEmailWithCC(dsaName As String, toAddress As String, ccAddress As String, attachmentPath As String)
    On Error GoTo ErrorHandler
    
    Dim outlookApp As Object
    Dim mailItem As Object
    Dim fso As Object
    
    Set outlookApp = CreateObject("Outlook.Application")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(attachmentPath) Then
        gErrorCount = gErrorCount + 1
        Call LogError("SendEmailWithCC", "Attachment file not found")
        Exit Sub
    End If
    
    Set mailItem = outlookApp.CreateItem(0)
    
    With mailItem
        .To = toAddress
        .CC = ccAddress
        .Subject = "DSA Update – " & dsaName
        .Body = GetEmailBody()
        .Attachments.Add attachmentPath
        .Display
    End With
    
    Set mailItem = Nothing
    Set outlookApp = Nothing
    Set fso = Nothing
    
    Exit Sub
    
ErrorHandler:
    gErrorCount = gErrorCount + 1
    Call LogError("SendEmailWithCC", Err.Description)
End Sub

' ============================================================================
' Send email with BCC
' ============================================================================
Sub SendEmailWithBCC(dsaName As String, toAddress As String, bccAddress As String, attachmentPath As String)
    On Error GoTo ErrorHandler
    
    Dim outlookApp As Object
    Dim mailItem As Object
    Dim fso As Object
    
    Set outlookApp = CreateObject("Outlook.Application")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(attachmentPath) Then
        gErrorCount = gErrorCount + 1
        Call LogError("SendEmailWithBCC", "Attachment file not found")
        Exit Sub
    End If
    
    Set mailItem = outlookApp.CreateItem(0)
    
    With mailItem
        .To = toAddress
        .BCC = bccAddress
        .Subject = "DSA Update – " & dsaName
        .Body = GetEmailBody()
        .Attachments.Add attachmentPath
        .Display
    End With
    
    Set mailItem = Nothing
    Set outlookApp = Nothing
    Set fso = Nothing
    
    Exit Sub
    
ErrorHandler:
    gErrorCount = gErrorCount + 1
    Call LogError("SendEmailWithBCC", Err.Description)
End Sub

' ============================================================================
' Validate email address format
' ============================================================================
Function IsValidEmail(emailAddress As String) As Boolean
    Dim emailPattern As String
    
    ' Simple validation - check for @ and .
    If Len(emailAddress) > 0 And InStr(emailAddress, "@") > 0 And InStr(emailAddress, ".") > InStr(emailAddress, "@") Then
        IsValidEmail = True
    Else
        IsValidEmail = False
    End If
End Function

' ============================================================================
' Log email error
' ============================================================================
Sub LogEmailError(dsaName As String, emailAddress As String, reason As String)
    On Error Resume Next
    
    Dim logsSheet As Worksheet
    Dim lastRow As Integer
    
    On Error Resume Next
    Set logsSheet = ThisWorkbook.Sheets("Logs")
    On Error GoTo 0
    
    If logsSheet Is Nothing Then
        Call CreateLogsSheet
        Set logsSheet = ThisWorkbook.Sheets("Logs")
    End If
    
    lastRow = logsSheet.Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    logsSheet.Cells(lastRow, 1).Value = Now()
    logsSheet.Cells(lastRow, 2).Value = dsaName
    logsSheet.Cells(lastRow, 4).Value = emailAddress
    logsSheet.Cells(lastRow, 5).Value = "FAILED"
    logsSheet.Cells(lastRow, 6).Value = reason
End Sub
