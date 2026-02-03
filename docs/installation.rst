Installation
============

Prerequisites
-------------

Before installing xlManage, ensure you have the following:

* **Python 3.14+** - Required for running xlManage
* **Windows OS** - Required for COM automation with Excel
* **Microsoft Excel** - Must be installed on your system
* **Poetry** - For dependency management (recommended)

Installation Methods
-------------------

Using Poetry (Recommended)
^^^^^^^^^^^^^^^^^^^^^^^^^^

1. Clone the repository:

.. code-block:: bash

   git clone https://github.com/your-repo/xlmanage.git
   cd xlmanage

2. Install dependencies:

.. code-block:: bash

   poetry install

3. Install the package in development mode:

.. code-block:: bash

   poetry install --with dev

Using pip
^^^^^^^^^

1. Install from source:

.. code-block:: bash

   pip install git+https://github.com/your-repo/xlmanage.git

2. Or install locally:

.. code-block:: bash

   git clone https://github.com/your-repo/xlmanage.git
   cd xlmanage
   pip install .

Verifying Installation
---------------------

Check that xlManage is installed correctly:

.. code-block:: bash

   xlmanage --version

You should see the version number (e.g., "xlManage version 0.1.0").

Troubleshooting
---------------

Common Issues
^^^^^^^^^^^^^^

**Excel not found**
   Ensure Microsoft Excel is installed and accessible via COM.

**Python version incompatible**
   Use Python 3.14 or higher as specified in the requirements.

**Poetry not installed**
   Install Poetry using: ``pip install poetry``

**Permission errors**
   Run commands with appropriate permissions or use virtual environments.

Getting Help
^^^^^^^^^^^^

For additional help:

* Check the :doc:`usage` guide for examples
* Review the :doc:`api` documentation for detailed function descriptions
* Visit our GitHub repository for issues and discussions
* Contact the maintainers for specific questions