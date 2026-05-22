Attribute VB_Name = "Module2_DataProcessing"
' ============================================================================
' BHFL CHANNEL OPERATIONS - DSA AUTOMATION
' Data Processing Module
' ============================================================================
' Purpose: Handle all data reading, filtering, and transformation logic
' ============================================================================

Option Explicit

' ============================================================================
' Read Master sheet and return as array
' ============================================================================
Function ReadMasterSheet() As Variant
    On Error GoTo ErrorHandler
    
    Dim masterSheet As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataArray As Variant
    
    Set masterSheet = ThisWorkbook.Sheets("Master")
    
    ' Find last row and column with data
    lastRow = masterSheet.Cells(Rows.Count, 1).End(xlUp).Row
    lastCol = masterSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    
    ' Read all data including headers
    If lastRow > 1 Then
        ReadMasterSheet = masterSheet.Range(Cells(1, 1), Cells(lastRow, lastCol)).Value
    Else
        ReadMasterSheet = Array()
    End If
    
    Exit Function
    
ErrorHandler:
    Call LogError("ReadMasterSheet", Err.Description)
    ReadMasterSheet = Array()
End Function

' ============================================================================
' Read MailMap sheet and return as array
' ============================================================================
Function ReadMailMapSheet() As Variant
    On Error GoTo ErrorHandler
    
    Dim mailMapSheet As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataArray As Variant
    
    Set mailMapSheet = ThisWorkbook.Sheets("MailMap")
    
    ' Find last row and column
    lastRow = mailMapSheet.Cells(Rows.Count, 1).End(xlUp).Row
    lastCol = mailMapSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    
    ' Read all data
    If lastRow > 1 Then
        ReadMailMapSheet = mailMapSheet.Range(Cells(1, 1), Cells(lastRow, lastCol)).Value
    Else
        ReadMailMapSheet = Array()
    End If
    
    Exit Function
    
ErrorHandler:
    Call LogError("ReadMailMapSheet", Err.Description)
    ReadMailMapSheet = Array()
End Function

' ============================================================================
' Get unique DSA names from master data
' ============================================================================
Function GetUniqueDSAs(masterData As Variant) As Collection
    On Error GoTo ErrorHandler
    
    Dim uniqueDSAs As New Collection
    Dim i As Long
    Dim dsaName As String
    Dim found As Boolean
    Dim j As Integer
    
    ' DSA Name is in column 4 (D)
    Const DSA_COLUMN As Integer = 4
    
    ' Loop through all data rows (skip header)
    For i = 2 To UBound(masterData, 1)
        dsaName = Trim(masterData(i, DSA_COLUMN))
        
        ' Skip blank DSAs
        If dsaName <> "" Then
            ' Check if already in collection
            found = False
            For j = 1 To uniqueDSAs.Count
                If uniqueDSAs(j) = dsaName Then
                    found = True
                    Exit For
                End If
            Next j
            
            ' Add if not found
            If Not found Then
                uniqueDSAs.Add dsaName
            End If
        End If
    Next i
    
    Set GetUniqueDSAs = uniqueDSAs
    
    Exit Function
    
ErrorHandler:
    Call LogError("GetUniqueDSAs", Err.Description)
    Set GetUniqueDSAs = New Collection
End Function

' ============================================================================
' Filter data by DSA name
' ============================================================================
Function FilterDataByDSA(masterData As Variant, dsaName As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim filteredArray() As Variant
    Dim i As Long
    Dim j As Integer
    Dim filteredRowCount As Long
    Dim headerAdded As Boolean
    
    ' DSA Name is in column 4
    Const DSA_COLUMN As Integer = 4
    
    ' First pass - count matching rows
    Dim matchCount As Long
    matchCount = 0
    
    For i = 2 To UBound(masterData, 1)
        If Trim(masterData(i, DSA_COLUMN)) = dsaName Then
            matchCount = matchCount + 1
        End If
    Next i
    
    ' Initialize filtered array with headers + matching rows
    ReDim filteredArray(1 To matchCount + 1, 1 To UBound(masterData, 2))
    
    ' Copy header row
    For j = 1 To UBound(masterData, 2)
        filteredArray(1, j) = masterData(1, j)
    Next j
    
    ' Copy matching rows
    filteredRowCount = 1
    For i = 2 To UBound(masterData, 1)
        If Trim(masterData(i, DSA_COLUMN)) = dsaName Then
            filteredRowCount = filteredRowCount + 1
            For j = 1 To UBound(masterData, 2)
                filteredArray(filteredRowCount, j) = masterData(i, j)
            Next j
        End If
    Next i
    
    FilterDataByDSA = filteredArray
    
    Exit Function
    
ErrorHandler:
    Call LogError("FilterDataByDSA", Err.Description)
    FilterDataByDSA = Array()
End Function

' ============================================================================
' Remove duplicate PAN numbers, keeping first occurrence
' ============================================================================
Function RemoveDuplicatePANs(filteredData As Variant) As Variant
    On Error GoTo ErrorHandler
    
    Dim resultArray() As Variant
    Dim i As Long
    Dim j As Integer
    Dim resultRow As Long
    Dim panNumber As String
    Dim panFound As Boolean
    Dim k As Long
    
    ' PAN is in column 3
    Const PAN_COLUMN As Integer = 3
    
    ' Initialize result array
    ReDim resultArray(1 To UBound(filteredData, 1), 1 To UBound(filteredData, 2))
    
    ' Copy header
    For j = 1 To UBound(filteredData, 2)
        resultArray(1, j) = filteredData(1, j)
    Next j
    
    resultRow = 1
    
    ' Loop through data and skip duplicates
    For i = 2 To UBound(filteredData, 1)
        panNumber = Trim(filteredData(i, PAN_COLUMN))
        
        ' Check if PAN is blank
        If panNumber = "" Then
            gErrorCount = gErrorCount + 1
            Call LogError("RemoveDuplicatePANs", "Blank PAN found for DSA row")
            GoTo NextPAN
        End If
        
        ' Check if PAN already exists in result
        panFound = False
        For k = 2 To resultRow
            If Trim(resultArray(k, PAN_COLUMN)) = panNumber Then
                panFound = True
                Exit For
            End If
        Next k
        
        ' Add if not duplicate
        If Not panFound Then
            resultRow = resultRow + 1
            For j = 1 To UBound(filteredData, 2)
                resultArray(resultRow, j) = filteredData(i, j)
            Next j
        End If
        
NextPAN:
    Next i
    
    ' Resize array to actual row count
    ReDim Preserve resultArray(1 To resultRow, 1 To UBound(filteredData, 2))
    
    RemoveDuplicatePANs = resultArray
    
    Exit Function
    
ErrorHandler:
    Call LogError("RemoveDuplicatePANs", Err.Description)
    RemoveDuplicatePANs = Array()
End Function

' ============================================================================
' Get email address for a DSA from MailMap data
' ============================================================================
Function GetEmailForDSA(mailMapData As Variant, dsaName As String) As String
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim emailAddress As String
    
    ' DSA Name in column 1, Email in column 2
    For i = 2 To UBound(mailMapData, 1)
        If Trim(mailMapData(i, 1)) = dsaName Then
            emailAddress = Trim(mailMapData(i, 2))
            GetEmailForDSA = emailAddress
            Exit Function
        End If
    Next i
    
    ' Not found
    GetEmailForDSA = ""
    
    Exit Function
    
ErrorHandler:
    Call LogError("GetEmailForDSA", Err.Description)
    GetEmailForDSA = ""
End Function

' ============================================================================
' Extract DSA name from file name
' ============================================================================
Function ExtractDSANameFromFileName(fileName As String) As String
    On Error GoTo ErrorHandler
    
    Dim dsaName As String
    Dim extensionPos As Integer
    
    ' Remove .xlsx extension
    extensionPos = InStr(fileName, ".xlsx")
    If extensionPos > 0 Then
        dsaName = Left(fileName, extensionPos - 1)
    Else
        dsaName = fileName
    End If
    
    ' Replace underscores with spaces
    dsaName = Replace(dsaName, "_", " ")
    
    ExtractDSANameFromFileName = dsaName
    
    Exit Function
    
ErrorHandler:
    Call LogError("ExtractDSANameFromFileName", Err.Description)
    ExtractDSANameFromFileName = ""
End Function

' ============================================================================
' Convert DSA name to file name format
' Example: ABC Finance → ABC_Finance
' ============================================================================
Function ConvertDSANameToFileName(dsaName As String) As String
    On Error GoTo ErrorHandler
    
    Dim fileName As String
    
    ' Replace spaces with underscores
    fileName = Replace(dsaName, " ", "_")
    
    ' Remove special characters
    fileName = Replace(fileName, "/", "_")
    fileName = Replace(fileName, "\", "_")
    fileName = Replace(fileName, ":", "_")
    fileName = Replace(fileName, "*", "_")
    fileName = Replace(fileName, "?", "_")
    fileName = Replace(fileName, """", "_")
    fileName = Replace(fileName, "<", "_")
    fileName = Replace(fileName, ">", "_")
    fileName = Replace(fileName, "|", "_")
    
    ConvertDSANameToFileName = fileName
    
    Exit Function
    
ErrorHandler:
    Call LogError("ConvertDSANameToFileName", Err.Description)
    ConvertDSANameToFileName = "DSA_File"
End Function
