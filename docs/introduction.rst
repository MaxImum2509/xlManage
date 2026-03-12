Introduction
============

About xlManage
--------------

xlManage is a CLI utility implemented in Python that controls Microsoft Excel
via COM automation (``pywin32``).  It is designed to be driven by an LLM agent
or by shell scripts, offering full programmatic control over Excel instances,
workbooks, worksheets, tables, VBA modules, and macro execution.

Key Features
------------

* **Excel instance control** -- start, stop, show/hide, and query running
  Excel processes.
* **Workbook management** -- open, create, save, close and list workbooks
  with options such as read-only and dev mode (``--dev`` disables
  ``Workbook_Open`` events).
* **Worksheet management** -- create, delete, copy, and list worksheets.
* **Table (ListObject) management** -- create, delete, and list Excel tables.
* **VBA module management** -- import/export standard modules (``.bas``),
  class modules (``.cls``), UserForms (``.frm/.frx``), and document modules
  (ThisWorkbook, Sheet).  Automatic encoding conversion from UTF-8 to
  Windows-1252 with CRLF normalization.
* **Macro execution** -- run VBA Sub and Function procedures with typed
  argument passing and configurable timeout.
* **Performance optimization** -- toggle screen updating, calculation mode
  and force recalculation.
* **Robust COM handling** -- automatic ``gen_py`` cache recovery, graceful
  disconnect without killing Excel, and ``Visibility`` enum to avoid
  side-effects on existing instances.

Use Cases
---------

xlManage is ideal for:

* Allowing an LLM agent to interact with Excel in real time
* Automating repetitive Excel and VBA tasks from the command line
* Importing/exporting VBA projects for version control
* Running macros in CI or automated workflows
* Managing large Excel datasets via scripts

Project Structure
-----------------

.. code-block:: text

   xlmanage/
   ├── src/
   │   └── xlmanage/
   │       ├── cli.py                  # Typer CLI entry point
   │       ├── excel_manager.py        # Excel instance lifecycle
   │       ├── workbook_manager.py     # Workbook CRUD
   │       ├── worksheet_manager.py    # Worksheet CRUD
   │       ├── table_manager.py        # Table (ListObject) CRUD
   │       ├── vba_manager.py          # VBA module import/export
   │       ├── macro_runner.py         # Macro execution
   │       ├── excel_optimizer.py      # Combined optimizer
   │       ├── screen_optimizer.py     # Screen updating optimizer
   │       ├── calculation_optimizer.py # Calculation mode optimizer
   │       └── exceptions.py           # Custom exception hierarchy
   ├── tests/
   ├── docs/
   └── pyproject.toml

Getting Started
---------------

To get started with xlManage, see the :doc:`installation` guide and then
explore the :doc:`usage` examples.
