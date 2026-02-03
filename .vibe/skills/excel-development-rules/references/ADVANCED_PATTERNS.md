# Advanced VBA Patterns

## FindRowByField - Array-Based Search

100x faster than cell-by-cell iteration. Raises error if field doesn't exist, returns Nothing if value not found.

```vba
Public Function FindRowByField(ByRef tbl As ListObject, _
                              ByVal fieldName As String, _
                              ByVal searchValue As Variant) As ListRow
    Dim colData As Variant
    Dim i As Long, maxIndex As Long
    Dim colDoesNotExist As Boolean

    If tbl Is Nothing Then
        Set FindRowByField = Nothing
        Exit Function
    End If

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

    i = 1
    maxIndex = UBound(colData, 1)
    Do While i <= maxIndex And colData(i, 1) <> searchValue
        i = i + 1
    Loop

    If i <= maxIndex Then
        Set FindRowByField = tbl.ListRows(i)
    Else
        Set FindRowByField = Nothing
    End If
End Function
```

## MVVM Architecture for UserForms

### Model (Type)

```vba
Public Type EmployeeData
    ID As Long
    Name As String
    Department As String
    Salary As Double
End Type
```

### ViewModel (Class Module)

```vba
' clsEmployeeViewModel
Private mEmployee As EmployeeData
Private mIsValid As Boolean

Public Sub LoadFromRow(ByRef row As ListRow)
    mEmployee.ID = row.Range.Cells(1, 1).Value
    mEmployee.Name = row.Range.Cells(1, 2).Value
    mEmployee.Department = row.Range.Cells(1, 3).Value
    mEmployee.Salary = row.Range.Cells(1, 4).Value
    Validate
End Sub

Public Sub SaveToRow(ByRef row As ListRow)
    If Not mIsValid Then Err.Raise vbObjectError + 1, , "Invalid data"
    row.Range.Cells(1, 2).Value = mEmployee.Name
    row.Range.Cells(1, 3).Value = mEmployee.Department
    row.Range.Cells(1, 4).Value = mEmployee.Salary
End Sub

Private Sub Validate()
    mIsValid = (Len(mEmployee.Name) > 0) And (mEmployee.Salary > 0)
End Sub
```

### View (UserForm)

```vba
Private mViewModel As clsEmployeeViewModel

Private Sub UserForm_Initialize()
    Set mViewModel = New clsEmployeeViewModel
End Sub

Public Sub EditEmployee(ByRef employeeRow As ListRow)
    mViewModel.LoadFromRow employeeRow
    txtName.Value = mViewModel.Name
    Me.Show vbModal
End Sub

Private Sub btnSave_Click()
    If mViewModel.IsValid Then
        mViewModel.SaveToRow targetRow
        Me.Hide
    End If
End Sub
```

## Event-Driven Architecture

### Application-Level Events (CAppEvents class)

```vba
Public WithEvents App As Application

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    EventRouter.HandleSheetChange Sh, Target
End Sub

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    EventRouter.HandleWorkbookOpen Wb
End Sub
```

Initialize globally:

```vba
Public gAppEvents As New CAppEvents

Public Sub InitGlobalEvents()
    Set gAppEvents.App = Application
End Sub
```

## Custom Error Types

```vba
Public Enum AppError
    ERR_VALIDATION = vbObjectError + 1001
    ERR_NOT_FOUND = vbObjectError + 1002
    ERR_PERMISSION = vbObjectError + 1003
End Enum

Public Function TryGetWorksheet(name As String) As Worksheet
    On Error Resume Next
    Set TryGetWorksheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
End Function
```
