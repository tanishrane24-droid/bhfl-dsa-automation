Attribute VB_Name = "Module3_FileOperations"
' ============================================================================
' BHFL CHANNEL OPERATIONS - DSA AUTOMATION
' File Operations Module
' ============================================================================
' Purpose: Handle folder creation, file creation, and file management
' ============================================================================

Option Explicit

' Windows API for file operations
Private Declare Function CreateDirectoryA Lib "kernel32" (ByVal lpPathName As String, lpSecurityAttributes As Any) As Long
Private Declare Function DeleteFileA Lib "kernel32" (ByVal lpFileName As String) As Long
Private Declare Function GetFileAttributesA Lib "kernel32" (ByVal lpFileName As String) As Long

Const FILE_ATTRIBUTE_DIRECTORY = 16

' ============================================================================
' Create output folder structure
' ============================================================================
Function CreateOutputFolders() As String
    On Error GoTo ErrorHandler
    
    Dim basePath As String
    Dim outputPath As String
    Dim archivePath As String
    
    ' Base path in Documents folder
    basePath = Environ("USERPROFILE") & "\Documents\BHFL_DSA_Automation\"
    
    ' Create base folder
    Call CreateFolder(basePath)
    
    ' Create Output subfolder
    outputPath = basePath & "Output\"
    Call CreateFolder(outputPath)
    
    ' Create Archive subfolder
    archivePath = basePath & "Archive\"
    Call CreateFolder(archivePath)
    
    ' Create Logs subfolder
    Call CreateFolder(basePath & "Logs\")
    
    CreateOutputFolders = outputPath
    
    Exit Function
    
ErrorHandler:
    Call LogError("CreateOutputFolders", Err.Description)
    CreateOutputFolders = ""
End Function

' ============================================================================
' Create folder if it doesn't exist
' ============================================================================
Function CreateFolder(folderPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if folder exists
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    
    CreateFolder = True
    
    Set fso = Nothing
    
    Exit Function
    
ErrorHandler:
    Call LogError("CreateFolder", Err.Description)
    CreateFolder = False
End Function

' ============================================================================
' Create DSA-specific Excel file
' ============================================================================
Sub CreateDSAFile(filteredData As Variant, filePath As String, dsaName As String)
    On Error GoTo ErrorHandler
    
    Dim newWorkbook As Workbook
    Dim newSheet As Worksheet
    Dim i As Long
    Dim j As Integer
    
    ' Create new workbook
    Set newWorkbook = Workbooks.Add
    Set newSheet = newWorkbook.Sheets(1)
    
    ' Write data to sheet
    Dim rowNum As Long
    Dim colNum As Integer
    
    ' Write all data from filtered array
    If UBound(filteredData) > 0 Then
        For i = LBound(filteredData, 1) To UBound(filteredData, 1)
            For j = LBound(filteredData, 2) To UBound(filteredData, 2)
                newSheet.Cells(i, j).Value = filteredData(i, j)
            Next j
        Next i
    End If
    
    ' Format headers
    With newSheet.Range("1:" & 1)
        .Font.Bold = True
        .Interior.Color = RGB(0, 51, 102)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' Auto-fit columns
    newSheet.Columns("A:H").AutoFit
    
    ' Add title row
    newSheet.Rows.Insert Shift:=xlDown
    newSheet.Range("A1").Value = "DSA: " & dsaName
    newSheet.Range("A1").Font.Bold = True
    newSheet.Range("A1").Font.Size = 12
    
    ' Save workbook
    newWorkbook.SaveAs fileName:=filePath, FileFormat:=xlOpenXMLWorkbook
    newWorkbook.Close
    
    ' Log file creation
    Call LogFileCreation(dsaName, filePath)
    
    Exit Sub
    
ErrorHandler:
    gErrorCount = gErrorCount + 1
    Call LogError("CreateDSAFile", Err.Description & " - " & filePath)
    On Error Resume Next
    newWorkbook.Close False
End Sub

' ============================================================================
' Get list of files in a folder
' ============================================================================
Function GetFilesInFolder(folderPath As String) As Collection
    On Error GoTo ErrorHandler
    
    Dim fileCollection As New Collection
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    ' Loop through files in folder
    For Each file In folder.Files
        ' Only add .xlsx files
        If Right(file.Name, 5) = ".xlsx" Then
            fileCollection.Add file.Name
        End If
    Next file
    
    Set GetFilesInFolder = fileCollection
    
    Set file = Nothing
    Set folder = Nothing
    Set fso = Nothing
    
    Exit Function
    
ErrorHandler:
    Call LogError("GetFilesInFolder", Err.Description)
    Set GetFilesInFolder = New Collection
End Function

' ============================================================================
' Log file creation to logs sheet
' ============================================================================
Sub LogFileCreation(dsaName As String, filePath As String)
    On Error Resume Next
    
    Dim logsSheet As Worksheet
    Dim lastRow As Integer
    
    ' Get or create logs sheet
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
    logsSheet.Cells(lastRow, 3).Value = Right(filePath, InStr(StrReverse(filePath), "\") - 1)
    logsSheet.Cells(lastRow, 5).Value = "SUCCESS"
    logsSheet.Cells(lastRow, 6).Value = "File created successfully"
End Sub

' ============================================================================
' Reverse a string (helper function)
' ============================================================================
Function StrReverse(inputStr As String) As String
    Dim i As Long
    Dim result As String
    
    For i = Len(inputStr) To 1 Step -1
        result = result & Mid(inputStr, i, 1)
    Next i
    
    StrReverse = result
End Function

' ============================================================================
' Delete folder and its contents
' ============================================================================
Function DeleteFolder(folderPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(folderPath) Then
        fso.DeleteFolder folderPath, True
    End If
    
    DeleteFolder = True
    
    Set fso = Nothing
    
    Exit Function
    
ErrorHandler:
    Call LogError("DeleteFolder", Err.Description)
    DeleteFolder = False
End Function

' ============================================================================
' Copy file from source to destination
' ============================================================================
Function CopyFile(sourceFile As String, destFile As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(sourceFile) Then
        fso.CopyFile sourceFile, destFile
        CopyFile = True
    Else
        CopyFile = False
    End If
    
    Set fso = Nothing
    
    Exit Function
    
ErrorHandler:
    Call LogError("CopyFile", Err.Description)
    CopyFile = False
End Function

' ============================================================================
' Move file from source to destination
' ============================================================================
Function MoveFile(sourceFile As String, destFile As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(sourceFile) Then
        fso.MoveFile sourceFile, destFile
        MoveFile = True
    Else
        MoveFile = False
    End If
    
    Set fso = Nothing
    
    Exit Function
    
ErrorHandler:
    Call LogError("MoveFile", Err.Description)
    MoveFile = False
End Function

' ============================================================================
' Check if file exists
' ============================================================================
Function FileExists(filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(filePath)
    
    Set fso = Nothing
    
    Exit Function
    
ErrorHandler:
    FileExists = False
End Function

' ============================================================================
' Get file size in bytes
' ============================================================================
Function GetFileSize(filePath As String) As Long
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(filePath) Then
        Set file = fso.GetFile(filePath)
        GetFileSize = file.Size
    Else
        GetFileSize = 0
    End If
    
    Set file = Nothing
    Set fso = Nothing
    
    Exit Function
    
ErrorHandler:
    GetFileSize = 0
End Function

' ============================================================================
' Get list of folders in a directory
' ============================================================================
Function GetFoldersInPath(folderPath As String) As Collection
    On Error GoTo ErrorHandler
    
    Dim folderCollection As New Collection
    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(folderPath) Then
        Set folder = fso.GetFolder(folderPath)
        
        For Each subFolder In folder.SubFolders
            folderCollection.Add subFolder.Name
        Next subFolder
    End If
    
    Set GetFoldersInPath = folderCollection
    
    Set subFolder = Nothing
    Set folder = Nothing
    Set fso = Nothing
    
    Exit Function
    
ErrorHandler:
    Set GetFoldersInPath = New Collection
End Function

' ============================================================================
' Clear all files in a folder
' ============================================================================
Sub ClearFolderContents(folderPath As String)
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(folderPath) Then
        Set folder = fso.GetFolder(folderPath)
        
        ' Delete all files
        For Each file In folder.Files
            file.Delete
        Next file
    End If
    
    Set file = Nothing
    Set folder = Nothing
    Set fso = Nothing
    
    Exit Sub
    
ErrorHandler:
    Call LogError("ClearFolderContents", Err.Description)
End Sub
