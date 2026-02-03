Attribute VB_Name = "ListObject_Patterns"
Option Explicit

' =============================================================================
' ADVANCED LISTOBJECT (TABLE) PATTERNS
' =============================================================================
' This module demonstrates enterprise-grade patterns for working with Excel
' ListObjects (Tables). Includes data validation, bulk operations, dynamic
' column management, and performance optimization techniques.
' =============================================================================

' =============================================================================
' TABLE MANAGEMENT CLASS
' =============================================================================

Private Type TableConfig
    TableName As String
    WorksheetName As String
    HasHeaders As Boolean
    AutoFilter As Boolean
End Type

' =============================================================================
' PUBLIC INTERFACE
' =============================================================================

Public Sub CreateTableFromRange(ByVal targetRange As Range, _
                                ByVal tableName As String, _
                                Optional ByVal hasHeaders As Boolean = True)
    'Creates a new ListObject from an existing range with validation

    On Error GoTo ErrorHandler

    ' Validate inputs
    If targetRange Is Nothing Then
        Err.Raise vbObjectError + 1, "CreateTableFromRange", "Target range cannot be Nothing"
    End If

    If Trim(tableName) = "" Then
        Err.Raise vbObjectError + 2, "CreateTableFromRange", "Table name cannot be empty"
    End If

    ' Check if table already exists
    If TableExists(targetRange.Worksheet.Parent, tableName) Then
        Err.Raise vbObjectError + 3, "CreateTableFromRange", _
            "Table '" & tableName & "' already exists"
    End If

    ' Create table
    Dim newTable As ListObject
    Set newTable = targetRange.Worksheet.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=targetRange, _
        XlListObjectHasHeaders:=IIf(hasHeaders, xlYes, xlNo))

    newTable.Name = tableName
    newTable.TableStyle = "TableStyleMedium2"

    Exit Sub

ErrorHandler:
    LogError Err.Number, Err.Description, "CreateTableFromRange"
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub BulkInsertData(ByVal tableName As String, _
                         ByVal dataArray() As Variant, _
                         Optional ByVal clearExisting As Boolean = False)
    'High-performance bulk insert using arrays
    ' dataArray must be 2-dimensional: (rows, columns)

    Dim tbl As ListObject
    Dim ws As Worksheet

    On Error GoTo ErrorHandler

    ' Get table reference
    Set tbl = GetTable(tableName)
    If tbl Is Nothing Then
        Err.Raise vbObjectError + 4, "BulkInsertData", "Table not found: " & tableName
    End If

    Set ws = tbl.Parent

    ' Performance optimization
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Clear existing data if requested
    If clearExisting Then
        ClearTableData tbl
    End If

    ' Validate array dimensions
    If LBound(dataArray, 2) <> 1 Then
        Err.Raise vbObjectError + 5, "BulkInsertData", _
            "Data array must start at column index 1"
    End If

    ' Ensure table has enough columns
    Dim requiredCols As Long
    requiredCols = UBound(dataArray, 2)

    Do While tbl.ListColumns.Count < requiredCols
        tbl.ListColumns.Add
    Loop

    ' Add rows and populate
    Dim startRow As Long
    startRow = tbl.ListRows.Count + 1

    Dim i As Long, j As Long
    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        Dim newRow As ListRow
        Set newRow = tbl.ListRows.Add

        For j = LBound(dataArray, 2) To UBound(dataArray, 2)
            newRow.Range.Cells(1, j).Value = dataArray(i, j)
        Next j
    Next i

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    LogError Err.Number, Err.Description, "BulkInsertData"
    GoTo CleanUp
End Sub

Public Function GetTableDataAsArray(ByVal tableName As String, _
                                   Optional ByVal includeHeaders As Boolean = False) As Variant
    'Exports table data to 2D array for processing

    Dim tbl As ListObject
    Dim dataRange As Range

    Set tbl = GetTable(tableName)
    If tbl Is Nothing Then
        GetTableDataAsArray = Array()
        Exit Function
    End If

    ' Determine range to export
    If includeHeaders Then
        Set dataRange = tbl.Range
    Else
        If tbl.ListRows.Count = 0 Then
            GetTableDataAsArray = Array()
            Exit Function
        End If
        Set dataRange = tbl.DataBodyRange
    End If

    GetTableDataAsArray = dataRange.Value
End Function

Public Sub FilterTable(ByVal tableName As String, _
                      ByVal fieldIndex As Long, _
                      ByVal criteria As Variant, _
                      Optional ByVal clearExisting As Boolean = True)
    'Applies AutoFilter to a table field

    Dim tbl As ListObject
    Set tbl = GetTable(tableName)

    If tbl Is Nothing Then Exit Sub

    ' Clear existing filters
    If clearExisting Then
        tbl.Range.AutoFilter
    End If

    ' Apply filter
    On Error Resume Next
    tbl.Range.AutoFilter Field:=fieldIndex, Criteria1:=criteria
    On Error GoTo 0
End Sub

Public Sub SortTable(ByVal tableName As String, _
                    ByVal sortColumn As String, _
                    Optional ByVal sortOrder As XlSortOrder = xlAscending, _
                    Optional ByVal hasHeaders As Boolean = True)
    'Sorts table by specified column

    Dim tbl As ListObject
    Dim ws As Worksheet

    Set tbl = GetTable(tableName)
    If tbl Is Nothing Then Exit Sub

    Set ws = tbl.Parent

    ' Clear any existing sort
    ws.Sort.SortFields.Clear

    ' Add sort field
    ws.Sort.SortFields.Add _
        Key:=tbl.ListColumns(sortColumn).Range, _
        SortOn:=xlSortOnValues, _
        Order:=sortOrder, _
        DataOption:=xlSortNormal

    ' Apply sort
    With ws.Sort
        .SetRange tbl.Range
        .Header = IIf(hasHeaders, xlYes, xlNo)
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Public Sub AddCalculatedColumn(ByVal tableName As String, _
                              ByVal columnName As String, _
                              ByVal formula As String)
    'Adds a new column with a formula that auto-fills

    Dim tbl As ListObject
    Dim newCol As ListColumn

    Set tbl = GetTable(tableName)
    If tbl Is Nothing Then Exit Sub

    ' Add column
    Set newCol = tbl.ListColumns.Add
    newCol.Name = columnName

    ' Add formula to first data cell
    ' Formula will auto-fill down in tables
    If tbl.ListRows.Count > 0 Then
        newCol.DataBodyRange.Cells(1, 1).Formula = formula
    End If
End Sub

Public Function FindInTable(ByVal tableName As String, _
                           ByVal searchColumn As String, _
                           ByVal searchValue As Variant) As ListRow
    'Finds first row matching criteria
    ' Returns Nothing if not found

    Dim tbl As ListObject
    Dim lr As ListRow
    Dim colIndex As Long

    Set tbl = GetTable(tableName)
    If tbl Is Nothing Then
        Set FindInTable = Nothing
        Exit Function
    End If

    On Error Resume Next
    colIndex = tbl.ListColumns(searchColumn).Index
    On Error GoTo 0

    If colIndex = 0 Then
        Set FindInTable = Nothing
        Exit Function
    End If

    For Each lr In tbl.ListRows
        If lr.Range.Cells(1, colIndex).Value = searchValue Then
            Set FindInTable = lr
            Exit Function
        End If
    Next lr

    Set FindInTable = Nothing
End Function

Public Sub DeleteMatchingRows(ByVal tableName As String, _
                             ByVal filterColumn As String, _
                             ByVal filterValue As Variant)
    'Deletes all rows matching criteria (efficient batch deletion)

    Dim tbl As ListObject
    Dim rowsToDelete As Collection
    Dim lr As ListRow
    Dim colIndex As Long
    Dim i As Long

    Set tbl = GetTable(tableName)
    If tbl Is Nothing Then Exit Sub

    colIndex = tbl.ListColumns(filterColumn).Index

    ' Collect rows to delete (iterate backwards-safe)
    Set rowsToDelete = New Collection
    For Each lr In tbl.ListRows
        If lr.Range.Cells(1, colIndex).Value = filterValue Then
            rowsToDelete.Add lr.Index
        End If
    Next lr

    ' Delete in reverse order to maintain index integrity
    Application.ScreenUpdating = False
    For i = rowsToDelete.Count To 1 Step -1
        tbl.ListRows(rowsToDelete(i)).Delete
    Next i
    Application.ScreenUpdating = True
End Sub

Public Sub ExportTableToCSV(ByVal tableName As String, _
                           ByVal filePath As String, _
                           Optional ByVal includeHeaders As Boolean = True, _
                           Optional ByVal delimiter As String = ",")
    'Exports table to CSV file

    Dim tbl As ListObject
    Dim dataArray() As Variant
    Dim i As Long, j As Long
    Dim fileNum As Integer
    Dim line As String

    Set tbl = GetTable(tableName)
    If tbl Is Nothing Then Exit Sub

    dataArray = GetTableDataAsArray(tableName, includeHeaders)

    fileNum = FreeFile
    Open filePath For Output As #fileNum

    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        line = ""
        For j = LBound(dataArray, 2) To UBound(dataArray, 2)
            If j > 1 Then line = line & delimiter
            line = line & EscapeCSVField(CStr(dataArray(i, j)))
        Next j
        Print #fileNum, line
    Next i

    Close #fileNum
End Sub

Public Function FindRowByField(ByRef tbl As ListObject, _
                              ByVal fieldName As String, _
                              ByVal searchValue As Variant) As ListRow
    '-------------------------------------------------------------------------------
    ' FindRowByField
    ' Description : Returns the first row where field equals search value
    ' Parameters  : tbl - ListObject - Table to search
    '             : fieldName - String - Column name to search
    '             : searchValue - Variant - Value to find
    ' Return      : ListRow - Found row or Nothing
    ' Author      : tbAffaires Team
    ' Date        : 2026-01-31
    '-------------------------------------------------------------------------------

    Dim colData As Variant
    Dim i As Long
    Dim maxIndex As Long
    Dim colDoesNotExist As Boolean

    ' Validate table
    If tbl Is Nothing Then
        Set FindRowByField = Nothing
        Exit Function
    End If

    ' Check if column exists
    On Error Resume Next
    colData = tbl.ListColumns(fieldName).DataBodyRange.Value
    colDoesNotExist = (Err.Number <> 0)
    On Error GoTo 0

    If colDoesNotExist Then
        Err.Raise vbObjectError + 1001, "FindRowByField", _
            "Field '" & fieldName & "' does not exist in table '" & tbl.Name & "'"
    End If

    If IsEmpty(colData) Then
        Set FindRowByField = Nothing
        Exit Function
    End If

    ' Initialize bounds
    i = 1
    maxIndex = UBound(colData, 1)

    ' Search with While loop
    Do While i <= maxIndex And colData(i, 1) <> searchValue
        i = i + 1
    Loop

    ' Return result
    If i <= maxIndex Then
        Set FindRowByField = tbl.ListRows(i)
    Else
        Set FindRowByField = Nothing
    End If
End Function

' =============================================================================
' HELPER FUNCTIONS
' =============================================================================

Private Function GetTable(ByVal tableName As String) As ListObject
    'Returns ListObject reference or Nothing if not found

    Dim ws As Worksheet
    Dim tbl As ListObject

    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set tbl = ws.ListObjects(tableName)
        On Error GoTo 0

        If Not tbl Is Nothing Then
            Set GetTable = tbl
            Exit Function
        End If
    Next ws

    Set GetTable = Nothing
End Function

Private Function TableExists(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    'Checks if table exists in workbook

    Dim ws As Worksheet
    Dim tbl As ListObject

    For Each ws In wb.Worksheets
        For Each tbl In ws.ListObjects
            If tbl.Name = tableName Then
                TableExists = True
                Exit Function
            End If
        Next tbl
    Next ws

    TableExists = False
End Function

Private Sub ClearTableData(ByVal tbl As ListObject)
    'Clears all data rows from table

    Do While tbl.ListRows.Count > 0
        tbl.ListRows(1).Delete
    Loop
End Sub

Private Function EscapeCSVField(ByVal fieldValue As String) As String
    'Escapes special characters for CSV output

    If InStr(fieldValue, ",") > 0 Or _
       InStr(fieldValue, "") > 0 Or _
       InStr(fieldValue, vbCr) > 0 Or _
       InStr(fieldValue, vbLf) > 0 Then

        ' Double up quotes
        fieldValue = Replace(fieldValue, """", """""")
        ' Wrap in quotes
        EscapeCSVField = """" & fieldValue & """"
    Else
        EscapeCSVField = fieldValue
    End If
End Function

Private Sub LogError(ByVal errorNumber As Long, _
                    ByVal errorDescription As String, _
                    ByVal procedureName As String)
    'Simple error logging - replace with your logging system
    Debug.Print Now & " | Error " & errorNumber & " in " & procedureName & ": " & errorDescription
End Sub

' =============================================================================
' USAGE EXAMPLES
' =============================================================================

Public Sub Example_Usage()
    'Example: Create and populate a table

    Dim testData(1 To 3, 1 To 3) As Variant
    testData(1, 1) = "John": testData(1, 2) = 30: testData(1, 3) = "Sales"
    testData(2, 1) = "Jane": testData(2, 2) = 25: testData(2, 3) = "Marketing"
    testData(3, 1) = "Bob": testData(3, 2) = 35: testData(3, 3) = "IT"

    ' Create table
    CreateTableFromRange Sheet1.Range("A1:C1"), "Employees"

    ' Insert data
    BulkInsertData "Employees", testData

    ' Add calculated column
    AddCalculatedColumn "Employees", "Bonus", "=[@Age]*100"

    ' Sort by age
    SortTable "Employees", "Age", xlDescending

    ' Filter to show only Sales
    FilterTable "Employees", 3, "Sales"
End Sub
