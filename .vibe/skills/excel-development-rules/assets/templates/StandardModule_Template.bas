Attribute VB_Name = "modTemplate"
'-------------------------------------------------------------------------------
' modTemplate
' Description : [Module description]
' Author      : [Name]
' Date        : [YYYY-MM-DD]
'-------------------------------------------------------------------------------

Option Explicit

'===============================================================================
' SECTION: Public constants
'===============================================================================

Public Const MODULE_NAME As String = "modTemplate"
Public Const MODULE_VERSION As String = "1.0.0"

'===============================================================================
' SECTION: Private constants
'===============================================================================

Private Const DEFAULT_VALUE As String = ""

'===============================================================================
' SECTION: Public procedures
'===============================================================================

'-------------------------------------------------------------------------------
' MainProcedure
' Description : Main entry point for this module
' Parameters  : None
' Return      : None
' Author      : [Name]
' Date        : [YYYY-MM-DD]
'-------------------------------------------------------------------------------
Public Sub MainProcedure()
    On Error GoTo ErrorHandler

    ' Use RAII pattern for optimization
    Dim optimizer As New ExcelOptimizer
    optimizer.Initialize

    ' Your code here

CleanUp:
    Exit Sub

ErrorHandler:
    HandleError Err.Number, Err.Description, "MainProcedure", Erl
    Resume CleanUp
End Sub

Public Function ProcessData(ByVal inputData As Variant) As Variant
    'Process data and return result
    ' Customize based on your needs

    On Error GoTo ErrorHandler

    ' Validate input
    If IsEmpty(inputData) Then
        Err.Raise vbObjectError + 1, "ProcessData", "Input data cannot be empty"
    End If

    ' Process data
    Dim result As Variant
    ' Add processing logic here

    ProcessData = result
    Exit Function

ErrorHandler:
    HandleError Err.Number, Err.Description, "ProcessData", Erl
    ProcessData = Empty
End Function

' =============================================================================
' PRIVATE HELPERS
' =============================================================================

Private Sub LogProcedureEntry(ByVal procName As String)
    'Log procedure entry (optional)
    Debug.Print Now & " | ENTER | " & MODULE_NAME & "." & procName
End Sub

Private Sub LogProcedureExit(ByVal procName As String)
    'Log procedure exit (optional)
    Debug.Print Now & " | EXIT | " & MODULE_NAME & "." & procName
End Sub

Private Sub HandleError(ByVal errNumber As Long, _
                       ByVal errDescription As String, _
                       ByVal procName As String, _
                       ByVal lineNum As Long)
    'Centralized error handling

    Dim msg As String
    msg = "Error " & errNumber & " in " & MODULE_NAME & "." & procName
    If lineNum > 0 Then
        msg = msg & " (Line " & lineNum & ")"
    End If
    msg = msg & ": " & errDescription

    ' Log to Immediate window
    Debug.Print msg

    ' Optionally display to user
    ' MsgBox msg, vbCritical
End Sub

Private Function IsValidRange(ByRef rng As Range) As Boolean
    'Validate range reference
    IsValidRange = Not rng Is Nothing
End Function

Private Function SafeTrim(ByVal value As Variant) As String
    'Safe trim that handles various data types
    If IsNull(value) Or IsEmpty(value) Then
        SafeTrim = vbNullString
    Else
        SafeTrim = Trim$(CStr(value))
    End If
End Function

' =============================================================================
' UTILITY FUNCTIONS
' =============================================================================

Public Function GetLastRow(ByRef ws As Worksheet, _
                          Optional ByVal columnLetter As String = "A") As Long
    'Get last row with data in specified column

    On Error Resume Next
    GetLastRow = ws.Cells(ws.Rows.Count, columnLetter).End(xlUp).Row
    On Error GoTo 0
End Function

Public Function WorksheetExists(ByVal sheetName As String) As Boolean
    'Check if worksheet exists in this workbook

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    WorksheetExists = Not ws Is Nothing
End Function

Public Sub SelectAllData(ByRef rng As Range)
    'Select all data in region (with error handling)

    On Error Resume Next
    If Not rng Is Nothing Then
        rng.CurrentRegion.Select
    End If
    On Error GoTo 0
End Sub
