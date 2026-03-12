Usage
=====

xlManage provides a CLI (``xlmanage``) organized in subcommands.
All commands communicate with Excel via COM automation.

Getting Help
------------

.. code-block:: bash

   # Global help
   xlmanage --help

   # Help on a subcommand group
   xlmanage workbook --help

   # Help on a specific command
   xlmanage workbook open --help

Excel Instance Management
-------------------------

Starting Excel
^^^^^^^^^^^^^^

.. code-block:: bash

   # Connect to an existing instance (or start a new one)
   xlmanage start

   # Start with the window visible
   xlmanage start --visible

   # Force a new instance
   xlmanage start --new

Stopping Excel
^^^^^^^^^^^^^^

.. code-block:: bash

   # Stop the active instance (saves open workbooks)
   xlmanage stop

   # Stop a specific instance by PID
   xlmanage stop 12345

   # Stop without saving
   xlmanage stop --no-save

   # Stop all instances
   xlmanage stop --all

   # Force kill (no save, uses taskkill)
   xlmanage stop --force

Instance Status
^^^^^^^^^^^^^^^

.. code-block:: bash

   # Show running Excel instances
   xlmanage status

Workbook Management
-------------------

.. code-block:: bash

   # Open a workbook
   xlmanage workbook open report.xlsx

   # Open in read-only mode
   xlmanage workbook open report.xlsx --read-only

   # Open in dev mode (disables Workbook_Open events)
   xlmanage workbook open macros.xlsm --dev

   # Open and control visibility
   xlmanage workbook open report.xlsx --visible
   xlmanage workbook open report.xlsx --hidden

   # Create a new workbook
   xlmanage workbook create output.xlsx

   # Create from a template
   xlmanage workbook create output.xlsx --template template.xltx

   # Save a workbook
   xlmanage workbook save report.xlsx

   # Save As (to a different path)
   xlmanage workbook save report.xlsx --as backup.xlsx

   # Close a workbook
   xlmanage workbook close report.xlsx

   # Close without saving
   xlmanage workbook close report.xlsx --no-save

   # List all open workbooks
   xlmanage workbook list

Worksheet Management
--------------------

.. code-block:: bash

   # List worksheets in a workbook
   xlmanage worksheet list -w report.xlsx

   # Create a new worksheet
   xlmanage worksheet create "Data" -w report.xlsx

   # Copy a worksheet
   xlmanage worksheet copy "Sheet1" "Sheet1_Copy" -w report.xlsx

   # Delete a worksheet
   xlmanage worksheet delete "TempSheet" -w report.xlsx

Table Management
----------------

.. code-block:: bash

   # List all tables in a workbook
   xlmanage table list -w report.xlsx

   # Create a table from a cell range
   xlmanage table create "MyTable" "A1:D10" -s "Sheet1" -w report.xlsx

   # Delete a table
   xlmanage table delete "MyTable" -w report.xlsx

VBA Module Management
---------------------

Importing Modules
^^^^^^^^^^^^^^^^^

xlManage supports importing standard modules (``.bas``), class modules
(``.cls``), UserForms (``.frm``), and document modules (ThisWorkbook,
Sheet -- detected automatically from ``.cls`` files).

Files encoded in UTF-8 are automatically converted to Windows-1252 with
CRLF line endings before import.

.. code-block:: bash

   # Import a standard module
   xlmanage vba import modules/modUtils.bas -w macros.xlsm

   # Import a class module
   xlmanage vba import modules/clsLogger.cls -w macros.xlsm

   # Import a UserForm
   xlmanage vba import modules/frmMain.frm -w macros.xlsm

   # Import with overwrite (replace existing module)
   xlmanage vba import modules/modUtils.bas -w macros.xlsm --overwrite

Exporting Modules
^^^^^^^^^^^^^^^^^

.. code-block:: bash

   # Export a module to a file
   xlmanage vba export modUtils output/modUtils.bas -w macros.xlsm

   # Export all modules
   xlmanage vba export --all output/ -w macros.xlsm

Listing and Deleting
^^^^^^^^^^^^^^^^^^^^

.. code-block:: bash

   # List all VBA modules
   xlmanage vba list -w macros.xlsm

   # Delete a module
   xlmanage vba delete modUtils -w macros.xlsm

Running Macros
--------------

.. code-block:: bash

   # Run a Sub
   xlmanage run-macro "Module1.MySub"

   # Run a Function with arguments
   xlmanage run-macro "Module1.GetSum" --args "10,20"

   # Run with string and boolean arguments
   xlmanage run-macro "Module1.Process" --args '"Report",true'

   # Specify the workbook containing the macro
   xlmanage run-macro "Module1.MySub" -w data.xlsm

   # Set a custom timeout (default: 60s)
   xlmanage run-macro "Module1.LongTask" --timeout 120

Performance Optimization
------------------------

.. code-block:: bash

   # Apply all optimizations (screen + calculation)
   xlmanage optimize

   # Optimize screen rendering only
   xlmanage optimize --screen

   # Optimize calculation only
   xlmanage optimize --calculation

   # Show current optimization status
   xlmanage optimize --status

   # Restore original settings
   xlmanage optimize --restore

   # Force full recalculation
   xlmanage optimize --force-calculate

See Also
--------

* :doc:`api` -- Detailed Python API documentation
* :doc:`installation` -- Installation instructions
