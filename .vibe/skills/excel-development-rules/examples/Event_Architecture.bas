Attribute VB_Name = "Event_Architecture"
Option Explicit

' =============================================================================
' EVENT-DRIVEN ARCHITECTURE FRAMEWORK
' =============================================================================
' This module demonstrates enterprise-grade event handling patterns including:
' - Custom event classes
' - Application-level event monitoring
' - Worksheet/Workbook event management
' - Event debouncing and throttling
' - Centralized event handling
' =============================================================================

' =============================================================================
' APPLICATION-LEVEL EVENTS CLASS
' =============================================================================
' Create a class module named "CAppEvents" with this code:

' CAppEvents.cls
' Public WithEvents App As Application
'
' Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
'     ' Delegate to centralized handler
'     EventRouter.HandleSheetChange Sh, Target
' End Sub
'
' Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
'     EventRouter.HandleWorkbookOpen Wb
' End Sub
'
' Private Sub App_SheetActivate(ByVal Sh As Object)
'     EventRouter.HandleSheetActivate Sh
' End Sub

' =============================================================================
' CUSTOM EVENT DEFINITIONS
' =============================================================================

Public Enum EventPriority
    PriorityLow = 1
    PriorityNormal = 2
    PriorityHigh = 3
    PriorityCritical = 4
End Enum

Public Type EventContext
    EventName As String
    Source As Object
    Timestamp As Date
    Priority As EventPriority
    Data As Variant
    Cancel As Boolean
End Type

' =============================================================================
' EVENT ROUTER MODULE
' =============================================================================
' This is the central dispatcher for all events

Private mEventHandlers As Collection
Private mProcessingEvents As Boolean
Private mEventQueue As Collection

Public Sub InitializeRouter()
    'Call this once at application startup
    Set mEventHandlers = New Collection
    Set mEventQueue = New Collection
    mProcessingEvents = False
End Sub

Public Sub RegisterHandler(ByVal eventName As String, _
                          ByVal handler As Object, _
                          ByVal methodName As String)
    'Registers an event handler

    Dim handlerInfo As Object
    Set handlerInfo = CreateObject("Scripting.Dictionary")

    handlerInfo("Event") = eventName
    handlerInfo("Handler") = handler
    handlerInfo("Method") = methodName

    mEventHandlers.Add handlerInfo
End Sub

Public Sub UnregisterHandler(ByVal eventName As String, _
                            ByVal handler As Object)
    'Removes a specific handler

    Dim i As Long
    Dim handlerInfo As Object

    For i = mEventHandlers.Count To 1 Step -1
        Set handlerInfo = mEventHandlers(i)
        If handlerInfo("Event") = eventName And _
           handlerInfo("Handler") Is handler Then
            mEventHandlers.Remove i
        End If
    Next i
End Sub

Public Sub RaiseEvent(ByVal eventName As String, _
                     Optional ByVal source As Object = Nothing, _
                     Optional ByVal data As Variant = Empty, _
                     Optional ByVal priority As EventPriority = PriorityNormal)
    'Raises an event to all registered handlers

    Dim evt As EventContext
    evt.EventName = eventName
    Set evt.Source = source
    evt.Timestamp = Now
    evt.Priority = priority
    evt.Data = data
    evt.Cancel = False

    ' Add to queue if currently processing
    If mProcessingEvents Then
        mEventQueue.Add evt
        Exit Sub
    End If

    ' Process immediately
    ProcessEvent evt
End Sub

Private Sub ProcessEvent(ByRef evt As EventContext)
    'Dispatches event to all registered handlers

    Dim handlerInfo As Object
    Dim handler As Object
    Dim method As String

    mProcessingEvents = True

    On Error Resume Next

    For Each handlerInfo In mEventHandlers
        If handlerInfo("Event") = evt.EventName Then
            Set handler = handlerInfo("Handler")
            method = handlerInfo("Method")

            ' Call handler method using Application.Run
            Application.Run method, evt

            ' Check if event was cancelled
            If evt.Cancel Then Exit For
        End If
    Next handlerInfo

    On Error GoTo 0

    mProcessingEvents = False

    ' Process queued events
    ProcessQueuedEvents
End Sub

Private Sub ProcessQueuedEvents()
    'Processes any events that were queued during handling

    Do While mEventQueue.Count > 0
        Dim evt As EventContext
        evt = mEventQueue(1)
        mEventQueue.Remove 1
        ProcessEvent evt
    Loop
End Sub

' =============================================================================
' WORKSHEET EVENT HANDLERS
' =============================================================================

Public Sub HandleSheetChange(ByVal Sh As Object, ByVal Target As Range)
    'Centralized handler for worksheet changes

    Dim evt As EventContext
    evt.EventName = "SheetChange"
    Set evt.Source = Sh
    evt.Timestamp = Now
    evt.Priority = PriorityNormal

    ' Store target range address as data
    evt.Data = Target.Address

    RaiseEvent "SheetChange", Sh, Target.Address
End Sub

Public Sub HandleWorkbookOpen(ByVal Wb As Workbook)
    'Centralized handler for workbook open events

    RaiseEvent "WorkbookOpen", Wb, Wb.Name, PriorityNormal
End Sub

Public Sub HandleSheetActivate(ByVal Sh As Object)
    'Centralized handler for sheet activation

    RaiseEvent "SheetActivate", Sh, Sh.Name, PriorityLow
End Sub

' =============================================================================
' DEBOUNCING UTILITIES
' =============================================================================

Private mDebounceTimers As Object 'Scripting.Dictionary

Public Sub DebouncedChangeHandler(ByVal Target As Range, _
                                 ByVal delayMs As Long, _
                                 ByVal callback As String)
    'Delays execution until changes stop for specified milliseconds
    ' Prevents rapid-fire processing during bulk edits

    Static initialized As Boolean
    If Not initialized Then
        Set mDebounceTimers = CreateObject("Scripting.Dictionary")
        initialized = True
    End If

    Dim key As String
    key = Target.Worksheet.Name & "_" & Target.Address

    ' Cancel existing timer
    If mDebounceTimers.Exists(key) Then
        Application.OnTime EarliestTime:=mDebounceTimers(key), _
                          Procedure:=callback, _
                          Schedule:=False
    End If

    ' Schedule new timer
    Dim nextTime As Date
    nextTime = Now + TimeSerial(0, 0, 0) + delayMs / 86400000

    mDebounceTimers(key) = nextTime
    Application.OnTime EarliestTime:=nextTime, Procedure:=callback
End Sub

Public Sub ClearDebounceTimer(ByVal Target As Range)
    'Cancels pending debounced operation

    If mDebounceTimers Is Nothing Then Exit Sub

    Dim key As String
    key = Target.Worksheet.Name & "_" & Target.Address

    If mDebounceTimers.Exists(key) Then
        Application.OnTime EarliestTime:=mDebounceTimers(key), _
                          Procedure:="", _
                          Schedule:=False
        mDebounceTimers.Remove key
    End If
End Sub

' =============================================================================
' CHANGE TRACKING SYSTEM
' =============================================================================

Private mChangeLog As Collection
Private mTrackingEnabled As Boolean

Public Sub StartChangeTracking()
    'Begins tracking all worksheet changes

    Set mChangeLog = New Collection
    mTrackingEnabled = True
End Sub

Public Sub StopChangeTracking()
    'Stops tracking changes

    mTrackingEnabled = False
End Sub

Public Sub LogChange(ByVal Target As Range, _
                    ByVal oldValue As Variant, _
                    ByVal newValue As Variant)
    'Logs a change to the tracking system

    If Not mTrackingEnabled Then Exit Sub

    Dim changeInfo As Object
    Set changeInfo = CreateObject("Scripting.Dictionary")

    changeInfo("Timestamp") = Now
    changeInfo("Worksheet") = Target.Worksheet.Name
    changeInfo("Address") = Target.Address
    changeInfo("OldValue") = oldValue
    changeInfo("NewValue") = newValue
    changeInfo("User") = Environ("UserName")

    mChangeLog.Add changeInfo
End Sub

Public Function GetChangeLog() As Collection
    'Returns the change log collection

    Set GetChangeLog = mChangeLog
End Function

Public Sub ClearChangeLog()
    'Clears all logged changes

    Set mChangeLog = New Collection
End Sub

Public Sub ExportChangeLog(ByVal filePath As String)
    'Exports change log to CSV file

    Dim fileNum As Integer
    Dim change As Variant

    fileNum = FreeFile
    Open filePath For Output As #fileNum

    ' Header
    Print #fileNum, "Timestamp,Worksheet,Address,OldValue,NewValue,User"

    ' Data
    For Each change In mChangeLog
        Print #fileNum, change("Timestamp") & "," & _
                        change("Worksheet") & "," & _
                        change("Address") & "," & _
                        change("OldValue") & "," & _
                        change("NewValue") & "," & _
                        change("User")
    Next change

    Close #fileNum
End Sub

' =============================================================================
' WORKSHEET CHANGE HANDLER TEMPLATE
' =============================================================================
' Add this to a worksheet's code module:

Private mOldValue As Variant
Private mIsProcessing As Boolean

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    'Store old value before change
    If Target.Cells.Count = 1 Then
        mOldValue = Target.Value
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    'Main change handler with guards

    ' Prevent recursive calls
    If mIsProcessing Then Exit Sub
    mIsProcessing = True

    ' Disable events temporarily
    Application.EnableEvents = False

    On Error GoTo CleanUp

    ' Log the change
    EventRouter.LogChange Target, mOldValue, Target.Value

    ' Raise custom event
    EventRouter.RaiseEvent "DataChanged", Me, _
        Array(Target.Address, mOldValue, Target.Value)

    ' Handle specific ranges
    If Not Intersect(Target, Me.Range("A:A")) Is Nothing Then
        ' Process column A changes
        ProcessColumnAChange Target
    End If

CleanUp:
    Application.EnableEvents = True
    mIsProcessing = False
End Sub

Private Sub ProcessColumnAChange(ByVal Target As Range)
    'Example: Auto-update related cells

    Dim rowNum As Long
    rowNum = Target.Row

    ' Update timestamp in column B
    Me.Cells(rowNum, 2).Value = Now

    ' Update status in column C
    Me.Cells(rowNum, 3).Value = "Modified"
End Sub

' =============================================================================
' INITIALIZATION
' =============================================================================

Public Sub InitializeEventSystem()
    'Call this from Workbook_Open

    ' Initialize router
    InitializeRouter

    ' Start change tracking
    StartChangeTracking

    ' Set up application-level events
    ' (Requires global CAppEvents instance in a standard module)
    ' Set gAppEvents.App = Application
End Sub

' =============================================================================
' USAGE EXAMPLES
' =============================================================================

Public Sub Example_SetupEventHandler()
    'Example: Set up a custom event handler

    ' Initialize the system
    InitializeEventSystem

    ' Register a handler for DataChanged events
    ' (Assumes you have a module/class with a HandleDataChanged sub)
    RegisterHandler "DataChanged", Nothing, "MyHandlerModule.HandleDataChanged"
End Sub

Public Sub Example_RaiseCustomEvent()
    'Example: Raise a custom event

    RaiseEvent "CustomEvent", _
               ThisWorkbook, _
               "Custom data here", _
               PriorityHigh
End Sub
