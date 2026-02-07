Usage
=====

Basic Commands
--------------

Getting Help
^^^^^^^^^^^^

To see all available commands and options:

.. code-block:: bash

   xlmanage --help

Version Information
^^^^^^^^^^^^^^^^^^^

Check the installed version:

.. code-block:: bash

   xlmanage --version

Common Workflows
----------------

Opening Excel Files
^^^^^^^^^^^^^^^^^^^

.. code-block:: bash

   # Open an Excel file
   xlmanage open workbook.xlsx

   # Open multiple files
   xlmanage open file1.xlsx file2.xlsx

Managing Worksheets
^^^^^^^^^^^^^^^^^^^

.. code-block:: bash

   # List all worksheets
   xlmanage list-sheets workbook.xlsx

   # Add a new worksheet
   xlmanage add-sheet workbook.xlsx "NewSheet"

Data Operations
^^^^^^^^^^^^^^^

.. code-block:: bash

   # Import data from CSV
   xlmanage import-csv workbook.xlsx data.csv "Sheet1"

   # Export data to CSV
   xlmanage export-csv workbook.xlsx "Sheet1" output.csv

Advanced Features
-----------------

Batch Processing
^^^^^^^^^^^^^^^^

Process multiple files in batch mode:

.. code-block:: bash

   xlmanage batch-process *.xlsx --operation optimize

Custom Scripts
^^^^^^^^^^^^^^

Run custom VBA scripts:

.. code-block:: bash

   xlmanage run-script workbook.xlsx script.vbs

Configuration
-------------

Configuration File
^^^^^^^^^^^^^^^^^^

xlManage uses a configuration file (``xlmanage.config``) for persistent settings.

Example configuration:

.. code-block:: yaml

   default_timeout: 30
   visible: false
   auto_save: true
   log_level: INFO

Environment Variables
^^^^^^^^^^^^^^^^^^^^^

.. code-block:: bash

   # Set timeout
   export XLMANAGE_TIMEOUT=60

   # Enable debug logging
   export XLMANAGE_LOG_LEVEL=DEBUG

Examples
--------

Basic Example
^^^^^^^^^^^^^

.. code-block:: bash

   # Optimize an Excel file
   xlmanage optimize input.xlsx output.xlsx

Advanced Example
^^^^^^^^^^^^^^^^

.. code-block:: bash

   # Process multiple files with custom settings
   xlmanage batch-process *.xlsx \
     --operation optimize \
     --timeout 60 \
     --visible false \
     --output optimized_

See Also
--------

* :doc:`api` - Detailed API documentation
* :doc:`installation` - Installation instructions
* :doc:`contributing` - How to contribute
