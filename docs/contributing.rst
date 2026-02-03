Contributing
============

Welcome to xlManage! We appreciate your interest in contributing to our project.

Getting Started
---------------

Fork the Repository
^^^^^^^^^^^^^^^^^^^^

1. Fork the repository on GitHub
2. Clone your fork locally

.. code-block:: bash

   git clone https://github.com/your-username/xlmanage.git
   cd xlmanage

3. Set up the development environment

.. code-block:: bash

   poetry install --with dev

Development Workflow
--------------------

Branch Strategy
^^^^^^^^^^^^^^^

We use the following branch strategy:

* ``main`` - Stable production code
* ``dev-epicXX-storyYY`` - Development branches for specific stories
* ``feature/`` - Feature branches
* ``bugfix/`` - Bug fix branches

Creating a Branch
^^^^^^^^^^^^^^^^^

.. code-block:: bash

   git checkout -b feature/your-feature-name

Making Changes
^^^^^^^^^^^^^^

1. Make your changes
2. Write tests for your changes
3. Ensure all tests pass
4. Update documentation if needed

Commit Guidelines
^^^^^^^^^^^^^^^^^

Follow our commit message conventions:

.. code-block:: text

   feat(module): add new feature
   fix(module): correct bug in existing feature
   docs: update documentation
   refactor(module): code refactoring
   test(module): add or update tests
   chore: maintenance tasks

Example:

.. code-block:: bash

   git commit -m "feat(cli): add new export command"

Submitting Changes
------------------

Push your changes:

.. code-block:: bash

   git push origin feature/your-feature-name

Create a Pull Request:

1. Go to the GitHub repository
2. Create a new Pull Request from your branch
3. Fill out the PR template
4. Request review from maintainers

Code Standards
--------------

Python Code
^^^^^^^^^^^

* Follow PEP 8 style guide
* Use type hints
* Write comprehensive docstrings
* Keep functions small and focused

Documentation
^^^^^^^^^^^^^

* Use reStructuredText format
* Follow Sphinx conventions
* Include code examples
* Keep documentation up-to-date

Testing
^^^^^^^

* Write unit tests for new features
* Maintain 90%+ code coverage
* Test edge cases
* Use pytest fixtures for complex setups

Review Process
--------------

1. Code review by at least one maintainer
2. All tests must pass
3. Documentation must be complete
4. Changes must follow project standards

Getting Help
------------

* Join our discussion forum
* Check the issue tracker
* Contact maintainers directly

Thank you for contributing to xlManage!
