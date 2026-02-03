Attribute VB_Name = "Error_Handling"
Option Explicit

' =============================================================================
' COMPREHENSIVE ERROR HANDLING FRAMEWORK
' =============================================================================
' This module provides enterprise-grade error handling including:
' - Centralized error logging
' - Structured exception handling
' - Error recovery strategies
' - Custom error types
' - Performance monitoring
' =============================================================================

' =============================================================================
' ERROR TYPES ENUMERATION
' =============================================================================

Public Enum ErrorType
    ERR_GENERAL = vbObjectError + 1000
    ERR_VALIDATION = vbObjectError + 1001
    ERR_DATA_ACCESS = vbObjectError + 1002
    ERR_FILE_IO = vbObjectError + 1003
    ERR_NETWORK = vbObjectError + 1004
    ERR_AUTHENTICATION = vbObjectError + 1005
    ERR_PERMISSION = vbObjectError + 1006
    ERR_TIMEOUT = vbObjectError + 1007
End Enum

Public Enum ErrorSeverity
    SEVERITY_INFO = 1
    SEVERITY_WARNING = 2
    SEVERITY_ERROR = 3
    SEVERITY_CRITICAL = 4
End Enum

' =============================================================================
' ERROR INFORMATION TYPE
' =============================================================================

Public Type ErrorInfo
    Number As Long
    Source As String
    Description As String
    Procedure As String
    Module As String
    LineNumber As Long
    Timestamp As Date
    Severity As ErrorSeverity
    UserContext As String
    StackTrace As String
    RecoveryAction As String
End Type

' =============================================================================
' ERROR LOGGER CLASS
' =============================================================================
' Create a class module named "CErrorLogger" with this code:

' CErrorLogger.cls
' Option Explicit
'
' Private mLogSheet As Worksheet
' Private mMaxLogRows As Long
'
' Public Sub Initialize()
'     Set mLogSheet = GetOrCreateLogSheet
'     mMaxLogRows = 10000 ' Archive after this
' End Sub
'
' Private Function GetOrCreateLogSheet() As Worksheet
'     Dim ws As Worksheet
'
'     On Error Resume Next
'     Set ws = ThisWorkbook.Worksheets("ErrorLog")
'     On Error GoTo 0
'
'     If ws Is Nothing Then
'         Set ws = ThisWorkbook.Worksheets.Add
'         ws.Name = "ErrorLog"
'         SetupLogHeaders ws
'     End If
'
'     Set GetOrCreateLogSheet = ws
' End Function
'
' Private Sub SetupLogHeaders(ws As Worksheet)
'     With ws.Range("A1:J1")
'         .Value = Array("Timestamp", "Severity", "Error Number", _
'                       "Description", "Procedure", "Module", _
'                       "Line", "User", "Stack Trace", "Recovery")
'         .Font.Bold = True
'     End With
'     ws.Columns.AutoFit
' End Sub
'
' Public Sub LogError(errInfo As ErrorInfo)
'     Dim nextRow As Long
'
'     nextRow = mLogSheet.Cells(mLogSheet.Rows.Count, 1).End(xlUp).Row + 1
'
'     ' Archive if too many rows
'     If nextRow > mMaxLogRows Then
'         ArchiveLog
'         nextRow = 2
'     End If
'
'     With mLogSheet
'         .Cells(nextRow, 1).Value = errInfo.Timestamp
'         .Cells(nextRow, 2).Value = GetSeverityText(errInfo.Severity)
'         .Cells(nextRow, 3).Value = errInfo.Number
'         .Cells(nextRow, 4).Value = errInfo.Description
'         .Cells(nextRow, 5).Value = errInfo.Procedure
'         .Cells(nextRow, 6).Value = errInfo.Module
'         .Cells(nextRow, 7).Value = errInfo.LineNumber
'         .Cells(nextRow, 8).Value = errInfo.UserContext
'         .Cells(nextRow, 9).Value = errInfo.StackTrace
'         .Cells(nextRow, 10).Value = errInfo.RecoveryAction
'     End With
' End Sub
'
' Private Function GetSeverityText(severity As ErrorSeverity) As String
'     Select Case severity
'         Case SEVERITY_INFO: GetSeverityText = "INFO"
'         Case SEVERITY_WARNING: GetSeverityText = "WARNING"
'         Case SEVERITY_ERROR: GetSeverityText = "ERROR"
'         Case SEVERITY_CRITICAL: GetSeverityText = "CRITICAL"
'     End Select
' End Function
'
' Private Sub ArchiveLog()
'     ' Archive old logs to separate sheet or file
'     ' Implementation depends on your needs
' End Sub

' =============================================================================
' GLOBAL ERROR HANDLER
' =============================================================================

Private mErrorLogger As CErrorLogger
Private mErrorStack As Collection
Private mCurrentProcedure As String

Public Sub InitializeErrorSystem()
    'Call once at application startup
    Set mErrorLogger = New CErrorLogger
    mErrorLogger.Initialize

    Set mErrorStack = New Collection
End Sub

Public Sub PushProcedure(ByVal procName As String)
    'Call at the start of each procedure
    mErrorStack.Add procName
    mCurrentProcedure = procName
End Sub

Public Sub PopProcedure()
    'Call at the end of each procedure
    If mErrorStack.Count > 0 Then
        mErrorStack.Remove mErrorStack.Count
    End If

    If mErrorStack.Count > 0 Then
        mCurrentProcedure = mErrorStack(mErrorStack.Count)
    Else
        mCurrentProcedure = ""
    End If
End Sub

Public Function BuildStackTrace() As String
    'Builds the current call stack
    Dim trace As String
    Dim i As Long

    For i = 1 To mErrorStack.Count
        If i > 1 Then trace = trace & " -> "
        trace = trace & mErrorStack(i)
    Next i

    BuildStackTrace = trace
End Function

' =============================================================================
' ERROR HANDLING PROCEDURES
' =============================================================================

Public Sub HandleError(ByVal procName As String, _
                      Optional ByVal customMessage As String = "", _
                      Optional ByVal severity As ErrorSeverity = SEVERITY_ERROR)
    'Centralized error handler

    Dim errInfo As ErrorInfo

    ' Populate error info
    With errInfo
        .Number = Err.Number
        .Source = Err.Source
        .Description = IIf(customMessage <> "", customMessage, Err.Description)
        .Procedure = procName
        .Module = GetCurrentModuleName
        .LineNumber = Erl ' Requires line numbers in code
        .Timestamp = Now
        .Severity = severity
        .UserContext = Environ("UserName") & "@" & Environ("ComputerName")
        .StackTrace = BuildStackTrace
        .RecoveryAction = ""
    End With

    ' Log the error
    mErrorLogger.LogError errInfo

    ' Display to user based on severity
    Select Case severity
        Case SEVERITY_INFO
            ' Silent logging, no display

        Case SEVERITY_WARNING
            MsgBox "Warning: " & errInfo.Description, vbExclamation, "Warning"

        Case SEVERITY_ERROR
            MsgBox "Error in " & procName & ":" & vbCrLf & _
                   errInfo.Description, vbCritical, "Error"

        Case SEVERITY_CRITICAL
            MsgBox "CRITICAL ERROR in " & procName & ":" & vbCrLf & _
                   errInfo.Description & vbCrLf & vbCrLf & _
                   "Please contact support immediately.", vbCritical, "Critical Error"
    End Select
End Sub

Public Sub RaiseCustomError(ByVal errorType As ErrorType, _
                           ByVal description As String, _
                           Optional ByVal severity As ErrorSeverity = SEVERITY_ERROR)
    'Raises a custom application error

    Err.Raise errorType, "Application", description
End Sub

' =============================================================================
' RECOVERY STRATEGIES
' =============================================================================

Public Function AttemptRecovery(ByVal errorNumber As Long) As Boolean
    'Attempts to recover from specific errors

    Select Case errorNumber
        Case 53 ' File not found
            AttemptRecovery = Recovery_FileNotFound

        Case 70 ' Permission denied
            AttemptRecovery = Recovery_PermissionDenied

        Case 91 ' Object variable not set
            AttemptRecovery = Recovery_ObjectNotSet

        Case 1004 ' Application-defined error
            AttemptRecovery = Recovery_ApplicationError

        Case Else
            AttemptRecovery = False
    End Select
End Function

Private Function Recovery_FileNotFound() As Boolean
    'Attempt to create missing file or use default
    ' Return True if recovery successful
    Recovery_FileNotFound = False
End Function

Private Function Recovery_PermissionDenied() As Boolean
    'Attempt to use alternative location or request elevated access
    Recovery_PermissionDenied = False
End Function

Private Function Recovery_ObjectNotSet() As Boolean
    'Attempt to initialize missing objects
    Recovery_ObjectNotSet = False
End Function

Private Function Recovery_ApplicationError() As Boolean
    'Handle Excel-specific errors
    Recovery_ApplicationError = False
End Function

' =============================================================================
' UTILITY FUNCTIONS
' =============================================================================

Private Function GetCurrentModuleName() As String
    'Returns the name of the current code module
    ' Note: VBA doesn't provide direct access to this
    GetCurrentModuleName = "Unknown"
End Function

Public Sub DisplayErrorLog()
    'Displays the error log to the user

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ErrorLog")
    On Error GoTo 0

    If Not ws Is Nothing Then
        ws.Activate
    Else
        MsgBox "No error log found.", vbInformation
    End If
End Sub

Public Sub ClearErrorLog()
    'Clears all logged errors

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ErrorLog")
    On Error GoTo 0

    If Not ws Is Nothing Then
        ws.Rows("2:" & ws.Rows.Count).ClearContents
    End If
End Sub

Public Sub ExportErrorLog(ByVal filePath As String)
    'Exports error log to CSV file

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ErrorLog")
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    ws.Copy
    With ActiveWorkbook
        .SaveAs fileName:=filePath, FileFormat:=xlCSV
        .Close SaveChanges:=False
    End With
End Sub

' =============================================================================
' STANDARD PROCEDURE TEMPLATE WITH ERROR HANDLING
' =============================================================================

Public Sub RobustProcedure_Template()
    'TEMPLATE: Copy this structure for all procedures

    On Error GoTo ErrorHandler

    ' Track procedure entry
    PushProcedure "RobustProcedure_Template"

    ' Store original application state
    Dim originalScreenUpdating As Boolean
    Dim originalCalculation As XlCalculation
    Dim originalEnableEvents As Boolean

    originalScreenUpdating = Application.ScreenUpdating
    originalCalculation = Application.Calculation
    originalEnableEvents = Application.EnableEvents

    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' =============================================================
    ' MAIN PROCEDURE LOGIC GOES HERE
    ' =============================================================



    ' =============================================================

CleanUp:
    ' Restore application state
    Application.ScreenUpdating = originalScreenUpdating
    Application.Calculation = originalCalculation
    Application.EnableEvents = originalEnableEvents

    ' Track procedure exit
    PopProcedure

    Exit Sub

ErrorHandler:
    ' Log and handle the error
    HandleError "RobustProcedure_Template"

    ' Attempt recovery if possible
    If AttemptRecovery(Err.Number) Then
        Resume ' Retry the operation
    End If

    ' Continue to cleanup
    Resume CleanUp
End Sub

' =============================================================================
' ASSERTION FRAMEWORK
' =============================================================================

Public Sub Assert(ByVal condition As Boolean, _
                 ByVal message As String, _
                 Optional ByVal failAction As String = "STOP")
    'Assertion for validating assumptions

    If Not condition Then
        Select Case UCase(failAction)
            Case "STOP"
                Err.Raise vbObjectError + 2000, "Assert", _
                    "Assertion failed: " & message

            Case "WARN"
                HandleError "Assert", "Assertion warning: " & message, SEVERITY_WARNING

            Case "LOG"
                HandleError "Assert", "Assertion logged: " & message, SEVERITY_INFO
        End Select
    End If
End Sub

Public Sub AssertNotNothing(ByVal obj As Object, ByVal objName As String)
    'Assert that an object is initialized
    Assert Not obj Is Nothing, objName & " cannot be Nothing", "STOP"
End Sub

Public Sub AssertNotEmpty(ByVal value As Variant, ByVal fieldName As String)
    'Assert that a value is not empty
    Assert Len(Trim(CStr(value))) > 0, fieldName & " cannot be empty", "STOP"
End Sub

Public Sub AssertInRange(ByVal value As Double, _
                        ByVal minVal As Double, _
                        ByVal maxVal As Double, _
                        ByVal fieldName As String)
    'Assert that a numeric value is within range
    Assert value >= minVal And value <= maxVal, _
           fieldName & " must be between " & minVal & " and " & maxVal, "STOP"
End Sub

' =============================================================================
' USAGE EXAMPLES
' =============================================================================

Public Sub Example_ErrorHandling()
    'Example: Using the error handling framework

    On Error GoTo ErrorHandler

    PushProcedure "Example_ErrorHandling"

    ' Use assertions
    Dim testValue As Long
    testValue = 50
    AssertInRange testValue, 0, 100, "testValue"

    ' Simulate an operation that might fail
    Dim result As Double
    result = 100 / 0 ' This will cause error 11

CleanUp:
    PopProcedure
    Exit Sub

ErrorHandler:
    HandleError "Example_ErrorHandling", "Division operation failed", SEVERITY_ERROR
    Resume CleanUp
End Sub

Public Sub Example_CustomError()
    'Example: Raising custom errors

    On Error GoTo ErrorHandler

    Dim userInput As String
    userInput = InputBox("Enter your age:")

    If Not IsNumeric(userInput) Then
        RaiseCustomError ERR_VALIDATION, "Age must be a number"
    End If

    If Val(userInput) < 18 Then
        RaiseCustomError ERR_VALIDATION, "Must be 18 or older"
    End If

    MsgBox "Valid age entered: " & userInput
    Exit Sub

ErrorHandler:
    HandleError "Example_CustomError"
End Sub

Public Sub Example_Recovery()
    'Example: Error with recovery attempt

    On Error GoTo ErrorHandler

    PushProcedure "Example_Recovery"

    ' Try to open a file
    Dim fileNum As Integer
    fileNum = FreeFile
    Open "C:\NonExistentFile.txt" For Input As #fileNum

    ' Read file...

    Close #fileNum

CleanUp:
    PopProcedure
    Exit Sub

ErrorHandler:
    HandleError "Example_Recovery"

    ' Try to recover
    If AttemptRecovery(Err.Number) Then
        Resume ' Retry
    Else
        Resume CleanUp
    End If
End Sub
