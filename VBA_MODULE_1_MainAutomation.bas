Attribute VB_Name = "Module1_MainAutomation"
' ============================================================================
' BHFL CHANNEL OPERATIONS - DSA AUTOMATION
' Main Automation Engine
' ============================================================================
' Version: 1.0
' Purpose: Main orchestration module for DSA data processing and email delivery
' ============================================================================

Option Explicit

' Global variables for tracking
Dim gDSACount As Integer
Dim gFilesGenerated As Integer
Dim gMailsGenerated As Integer
Dim gErrorCount As Integer

' ============================================================================
' MAIN ENTRY POINT - Called when START AUTOMATION button is clicked
' ============================================================================
Sub StartAutomation()
    On Error GoTo ErrorHandler
    
    Dim startTime As Double
    startTime = Timer
    
    ' Initialize counters
    gDSACount = 0
    gFilesGenerated = 0
    gMailsGenerated = 0
    gErrorCount = 0
    
    ' Step 1: Generate demo data if Master sheet is empty
    Call UpdateProgress("Generating Demo Data...")
    Call GenerateDemoData
    
    ' Step 2: Read Master sheet
    Call UpdateProgress("Reading Master Data...")
    Dim masterData As Variant
    masterData = ReadMasterSheet()
    
    ' Step 3: Identify unique DSAs
    Call UpdateProgress("Detecting DSAs...")
    Dim uniqueDSAs As Collection
    Set uniqueDSAs = GetUniqueDSAs(masterData)
    gDSACount = uniqueDSAs.Count
    
    ' Step 4: Create Output folder
    Call UpdateProgress("Creating Output Folders...")
    Dim outputPath As String
    outputPath = CreateOutputFolders()
    
    ' Step 5: Process each DSA
    Call UpdateProgress("Creating Files...")
    Call ProcessEachDSA(masterData, uniqueDSAs, outputPath)
    
    ' Step 6: Prepare emails
    Call UpdateProgress("Preparing Mails...")
    Call PrepareAndSendEmails(outputPath)
    
    ' Step 7: Create logs
    Call UpdateProgress("Creating Logs...")
    Call CreateLogsSheet()
    
    ' Step 8: Move files to archive
    Call MoveToArchive(outputPath)
    
    ' Show completion summary
    Call ShowCompletionSummary(startTime)
    
    Exit Sub
    
ErrorHandler:
    gErrorCount = gErrorCount + 1
    Call LogError("StartAutomation", Err.Description)
    MsgBox "Error in automation: " & Err.Description, vbCritical, "Automation Error"
End Sub

' ============================================================================
' Update progress message on screen
' ============================================================================
Sub UpdateProgress(message As String)
    On Error Resume Next
    StatusBar = message
    Application.StatusBar = message
    DoEvents
End Sub

' ============================================================================
' Show final completion summary popup
' ============================================================================
Sub ShowCompletionSummary(startTime As Double)
    Dim endTime As Double
    Dim duration As Double
    
    endTime = Timer
    duration = endTime - startTime
    
    Dim summaryMessage As String
    summaryMessage = "AUTOMATION COMPLETED" & vbCrLf & vbCrLf & _
                     "DSA Processed: " & gDSACount & vbCrLf & _
                     "Files Generated: " & gFilesGenerated & vbCrLf & _
                     "Mails Generated: " & gMailsGenerated & vbCrLf & _
                     "Errors: " & gErrorCount & vbCrLf & vbCrLf & _
                     "Duration: " & Format(duration, "0.00") & " seconds"
    
    MsgBox summaryMessage, vbInformation, "AUTOMATION COMPLETED"
    Application.StatusBar = False
End Sub

' ============================================================================
' Process each DSA - Main processing loop
' ============================================================================
Sub ProcessEachDSA(masterData As Variant, uniqueDSAs As Collection, outputPath As String)
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim dsaName As String
    Dim filteredData As Variant
    Dim fileName As String
    Dim filePath As String
    
    ' Loop through each unique DSA
    For i = 1 To uniqueDSAs.Count
        dsaName = uniqueDSAs(i)
        
        ' Skip blank DSAs
        If dsaName = "" Or dsaName = "DSA Name" Then
            gErrorCount = gErrorCount + 1
            Call LogError("ProcessEachDSA", "Blank DSA Name found and skipped")
            GoTo NextDSA
        End If
        
        ' Filter data for this DSA
        filteredData = FilterDataByDSA(masterData, dsaName)
        
        ' Remove duplicates based on PAN
        filteredData = RemoveDuplicatePANs(filteredData)
        
        ' Create Excel file for this DSA
        fileName = ConvertDSANameToFileName(dsaName)
        filePath = outputPath & fileName & ".xlsx"
        
        ' Create the file
        Call CreateDSAFile(filteredData, filePath, dsaName)
        gFilesGenerated = gFilesGenerated + 1
        
NextDSA:
    Next i
    
    Exit Sub
    
ErrorHandler:
    gErrorCount = gErrorCount + 1
    Call LogError("ProcessEachDSA", Err.Description)
End Sub

' ============================================================================
' Prepare and send emails for each DSA
' ============================================================================
Sub PrepareAndSendEmails(outputPath As String)
    On Error GoTo ErrorHandler
    
    Dim mailMapData As Variant
    Dim fileList As Collection
    Dim i As Integer
    Dim fileName As String
    Dim dsaName As String
    Dim emailAddress As String
    
    ' Read MailMap
    mailMapData = ReadMailMapSheet()
    
    ' Get list of files in output folder
    Set fileList = GetFilesInFolder(outputPath)
    
    ' Process each file
    For i = 1 To fileList.Count
        fileName = fileList(i)
        dsaName = ExtractDSANameFromFileName(fileName)
        
        ' Get email from MailMap
        emailAddress = GetEmailForDSA(mailMapData, dsaName)
        
        ' Validate email
        If emailAddress = "" Or InStr(emailAddress, "@") = 0 Then
            gErrorCount = gErrorCount + 1
            Call LogEmailError(dsaName, emailAddress, "Missing or Invalid Email")
            GoTo NextEmail
        End If
        
        ' Send email
        Call SendEmailWithAttachment(dsaName, emailAddress, outputPath & fileName)
        gMailsGenerated = gMailsGenerated + 1
        
NextEmail:
    Next i
    
    Exit Sub
    
ErrorHandler:
    gErrorCount = gErrorCount + 1
    Call LogError("PrepareAndSendEmails", Err.Description)
End Sub

' ============================================================================
' Move processed files to archive folder
' ============================================================================
Sub MoveToArchive(outputPath As String)
    On Error GoTo ErrorHandler
    
    Dim archivePath As String
    Dim fileList As Collection
    Dim i As Integer
    Dim oldPath As String
    Dim newFileName As String
    Dim newPath As String
    
    ' Create archive folder
    archivePath = CreateFolder(Left(outputPath, InStrRev(outputPath, "\") - 1) & "\Archive\")
    
    ' Get files from output folder
    Set fileList = GetFilesInFolder(outputPath)
    
    ' Move each file
    For i = 1 To fileList.Count
        oldPath = outputPath & fileList(i)
        newFileName = fileList(i)
        newFileName = Left(newFileName, InStrRev(newFileName, ".") - 1) & "_" & Format(Now(), "yyyymmdd_hhmmss") & ".xlsx"
        newPath = archivePath & newFileName
        
        ' Move file using Windows API or FileSystem
        On Error Resume Next
        Name oldPath As newPath
        On Error GoTo ErrorHandler
    Next i
    
    Exit Sub
    
ErrorHandler:
    Call LogError("MoveToArchive", Err.Description)
End Sub

' ============================================================================
' Create Logs Sheet with activity log
' ============================================================================
Sub CreateLogsSheet()
    On Error GoTo ErrorHandler
    
    Dim logsSheet As Worksheet
    Dim rowNum As Integer
    
    ' Create or clear logs sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Logs").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler
    
    ' Add new logs sheet
    Set logsSheet = ThisWorkbook.Sheets.Add
    logsSheet.Name = "Logs"
    
    ' Add headers
    logsSheet.Range("A1").Value = "Date"
    logsSheet.Range("B1").Value = "DSA"
    logsSheet.Range("C1").Value = "File Name"
    logsSheet.Range("D1").Value = "Mail ID"
    logsSheet.Range("E1").Value = "Status"
    logsSheet.Range("F1").Value = "Remarks"
    
    ' Format headers
    With logsSheet.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(0, 51, 102)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' Add sample logs (in production, fill with actual data)
    ' This is where actual logging would happen
    
    ' Auto-fit columns
    logsSheet.Columns("A:F").AutoFit
    
    Exit Sub
    
ErrorHandler:
    Call LogError("CreateLogsSheet", Err.Description)
End Sub

' ============================================================================
' Log errors to Error_Log sheet
' ============================================================================
Sub LogError(functionName As String, errorDescription As String)
    On Error Resume Next
    
    Dim errorSheet As Worksheet
    Dim lastRow As Integer
    
    ' Create error log sheet if not exists
    On Error Resume Next
    Set errorSheet = ThisWorkbook.Sheets("Error_Log")
    If errorSheet Is Nothing Then
        Set errorSheet = ThisWorkbook.Sheets.Add
        errorSheet.Name = "Error_Log"
        errorSheet.Range("A1").Value = "Timestamp"
        errorSheet.Range("B1").Value = "Function"
        errorSheet.Range("C1").Value = "Error Description"
    End If
    On Error GoTo 0
    
    ' Add error to log
    lastRow = errorSheet.Cells(Rows.Count, 1).End(xlUp).Row + 1
    errorSheet.Cells(lastRow, 1).Value = Now()
    errorSheet.Cells(lastRow, 2).Value = functionName
    errorSheet.Cells(lastRow, 3).Value = errorDescription
    
    ' Auto-fit columns
    errorSheet.Columns("A:C").AutoFit
End Sub

' ============================================================================
' Log email-specific errors
' ============================================================================
Sub LogEmailError(dsaName As String, emailAddress As String, reason As String)
    On Error Resume Next
    
    Dim logsSheet As Worksheet
    Dim lastRow As Integer
    
    Set logsSheet = ThisWorkbook.Sheets("Logs")
    lastRow = logsSheet.Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    logsSheet.Cells(lastRow, 1).Value = Now()
    logsSheet.Cells(lastRow, 2).Value = dsaName
    logsSheet.Cells(lastRow, 4).Value = emailAddress
    logsSheet.Cells(lastRow, 5).Value = "FAILED"
    logsSheet.Cells(lastRow, 6).Value = reason
End Sub

' ============================================================================
' Generate demo data if Master sheet is empty
' ============================================================================
Sub GenerateDemoData()
    On Error GoTo ErrorHandler
    
    Dim masterSheet As Worksheet
    Dim rowCount As Integer
    
    Set masterSheet = ThisWorkbook.Sheets("Master")
    rowCount = masterSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' If sheet has data beyond header, skip demo data generation
    If rowCount > 1 Then
        Exit Sub
    End If
    
    ' Call module to generate demo data
    Call PopulateDemoMasterData
    Call PopulateDemoMailMapData
    
    Exit Sub
    
ErrorHandler:
    Call LogError("GenerateDemoData", Err.Description)
End Sub
