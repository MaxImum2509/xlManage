# Troubleshooting VBA

## Encoding Issues

### Garbled characters in VBA Editor
**Cause**: File saved with UTF-8 instead of Windows-1252.
**Fix**: `open(file, 'w', encoding='windows-1252', newline='\r\n')`

### Code on single line
**Fix**: Use CRLF (`\r\n`), not LF (`\n`).

## Import Errors

### "Module name already exists"
```python
for comp in vb_project.VBComponents:
    if comp.Name == "modMyModule":
        vb_project.VBComponents.Remove(comp)
        break
vb_project.VBComponents.Import(module_path)
```

### Class import fails silently
1. Verify file starts with `VERSION 1.0 CLASS`
2. Ensure `Option Explicit` is present after header
3. Check that all `Attribute` lines are removed before code injection

## Runtime Errors

### Error 91: "Object variable not set"
```vba
' Wrong: missing Set
Dim ws As Worksheet
ws = ThisWorkbook.Worksheets("Data")

' Correct
Set ws = ThisWorkbook.Worksheets("Data")

' Safe pattern
On Error Resume Next
Set ws = ThisWorkbook.Worksheets("Data")
On Error GoTo 0
If ws Is Nothing Then Exit Sub
```

### Error 9: "Subscript out of range"
Always validate before accessing:
```vba
If idx >= LBound(arr) And idx <= UBound(arr) Then
    value = arr(idx)
End If
```

### Error 13: "Type mismatch"
```vba
If IsNumeric(cell.Value) Then
    num = CDbl(cell.Value)
Else
    num = 0
End If
```

## Performance Issues

### Slow macro execution
Use ExcelOptimizer (RAII):
```vba
Dim optimizer As New ExcelOptimizer
optimizer.Initialize
' Your code - settings auto-restored
```

### Memory issues with large datasets
Process in chunks:
```vba
Const CHUNK_SIZE As Long = 10000
For i = 1 To lastRow Step CHUNK_SIZE
    ProcessRows i, Application.Min(i + CHUNK_SIZE - 1, lastRow)
    DoEvents
Next i
```

## Python/VBAImporter Issues

### "RPC_E_SERVERFAULT" (0x800706be)
**Cause**: Called `excel.Quit()` while Python holds COM references.
**Fix**: Remove `excel.Quit()`. Let GC handle cleanup.

### Module import fails with "Module already exists"
```python
importer.import_directory(r"C:\vba\modules", overwrite=True)
```

### UserForm import missing UI elements
Ensure both `.frm` and `.frx` files are in same directory.

### Class module wrong PredeclaredId
Use VBAImporter which auto-handles VB_PredeclaredId parsing.

### Import succeeds but code doesn't appear
Check component type: 1=StdModule, 2=ClassModule, 3=MSForm.
