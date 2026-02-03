# COM Application Catalog

Reference table of common Windows COM applications with their ProgIDs, server types, tips, and known gotchas.

## Office Applications

| Application | ProgID | Server Type | Notes |
|-------------|--------|-------------|-------|
| **Excel** | `Excel.Application` | Out-of-process | Never call `Quit()`. Use `ScreenUpdating=False` for perf. See `excel-python-tooling` skill. |
| **Word** | `Word.Application` | Out-of-process | `Documents.Add()` / `Documents.Open()`. Use `doc.Content.Text` for full text. |
| **Outlook** | `Outlook.Application` | Out-of-process | Security prompts for `MailItem.Send()`. Use Redemption library to bypass. Singleton - always returns running instance. |
| **PowerPoint** | `PowerPoint.Application` | Out-of-process | `Presentations.Add()`. Slide indexing is 1-based. `Visible` must be set to `True` before opening presentations on some versions. |
| **Access** | `Access.Application` | Out-of-process | `OpenCurrentDatabase()` for .accdb. Use `DoCmd` for actions. `CurrentDb.Execute` for SQL. |

## Engineering / CAD

| Application | ProgID | Server Type | Notes |
|-------------|--------|-------------|-------|
| **AutoCAD** | `AutoCAD.Application` | Out-of-process | `GetObject()` preferred for running instance. Document-level scripting via `ActiveDocument.SendCommand`. |
| **SolidWorks** | `SldWorks.Application` | Out-of-process | Use `GetObject` for running instance. API documentation is essential - complex object model. |

## Enterprise / ERP

| Application | ProgID | Server Type | Notes |
|-------------|--------|-------------|-------|
| **SAP GUI** | `Sapgui.ScriptingCtrl.1` | Out-of-process | Requires SAP GUI Scripting enabled in server settings. `GetScriptingEngine` from connection. |

## System / Infrastructure

| Application | ProgID | Server Type | Notes |
|-------------|--------|-------------|-------|
| **WMI** | `WbemScripting.SWbemLocator` | In-process | Better to use `wmi` Python package. `ConnectServer()` for remote machines. |
| **ADODB** | `ADODB.Connection` | In-process | `Open(connection_string)`. Use `ADODB.Recordset` for result sets. Consider `pyodbc` as alternative. |
| **Shell** | `Shell.Application` | In-process | File operations, `ShellExecute`, `BrowseForFolder`. Limited vs `subprocess`/`pathlib`. |
| **Windows Script Host** | `WScript.Shell` | In-process | `Run`, `Exec`, registry operations. `subprocess` is usually better for Python. |
| **Task Scheduler** | `Schedule.Service` | In-process | `Connect()` then `GetFolder().GetTasks()`. Complex object model for task creation. |
| **Windows Installer** | `WindowsInstaller.Installer` | In-process | MSI database operations. `OpenDatabase` for reading/writing MSI files. |

## Scripting Engines

| Application | ProgID | Server Type | Notes |
|-------------|--------|-------------|-------|
| **Internet Explorer** | `InternetExplorer.Application` | Out-of-process | Deprecated. Use Selenium/Playwright instead. |
| **MSXML HTTP** | `MSXML2.XMLHTTP` | In-process | HTTP requests. Use `requests` library instead. |
| **FileSystemObject** | `Scripting.FileSystemObject` | In-process | File operations. Use `pathlib` instead. |

## Usage Examples

### Excel (see excel-python-tooling skill for full guide)

```python
excel = win32.Dispatch("Excel.Application")
excel.Visible = False
wb = excel.Workbooks.Add()
ws = wb.Worksheets(1)
ws.Range("A1").Value = "Hello"
wb.SaveAs(r"C:\output.xlsx", 51)
wb.Close()
# Do NOT call excel.Quit()
```

### Word

```python
word = win32.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Add()
doc.Content.Text = "Hello, World!"
doc.SaveAs2(r"C:\output.docx", 16)  # wdFormatDocumentDefault
doc.Close()
```

### Outlook

```python
outlook = win32.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)  # olMailItem
mail.To = "user@example.com"
mail.Subject = "Test"
mail.Body = "Hello from Python"
mail.Send()
```

### PowerPoint

```python
ppt = win32.Dispatch("PowerPoint.Application")
ppt.Visible = True  # Required before opening files on some versions
pres = ppt.Presentations.Add()
slide = pres.Slides.Add(1, 1)  # ppLayoutText
slide.Shapes.Title.TextFrame.TextRange.Text = "Title"
pres.SaveAs(r"C:\output.pptx")
pres.Close()
```

### WMI (System Info)

```python
wmi = win32.Dispatch("WbemScripting.SWbemLocator")
service = wmi.ConnectServer(".", "root\\cimv2")
for proc in service.ExecQuery("SELECT * FROM Win32_Process"):
    print(f"{proc.Name} (PID: {proc.ProcessId})")
```

### ADODB (Database)

```python
conn = win32.Dispatch("ADODB.Connection")
conn.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\data.accdb;")
rs = win32.Dispatch("ADODB.Recordset")
rs.Open("SELECT * FROM Users", conn)
while not rs.EOF:
    print(rs.Fields("Name").Value)
    rs.MoveNext()
rs.Close()
conn.Close()
```

## Tips for Discovering COM Interfaces

1. **OLE/COM Object Viewer** (`oleview.exe`): Browse registered COM classes and interfaces
2. **Python `makepy` utility**: `python -m win32com.client.makepy` to generate type libraries interactively
3. **Registry**: COM classes registered under `HKCR\CLSID` and `HKCR\<ProgID>`
4. **VBA Object Browser**: Open VBA editor in any Office app, press F2 to browse available objects
