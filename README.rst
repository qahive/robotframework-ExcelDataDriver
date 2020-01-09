ExcelDataDriver
===============
.. figure:: https://github.com/qahive/robotframework-ExcelDataDriver/workflows/Python%20package/badge.svg
.. contents::

Introduction
------------

ExcelDataDriver_ is a Excel Data-Driven Testing library for `RobotFramework <http://code.google.com/p/robotframework/>`_.
The project is hosted on `GitHub <https://github.com/qahive/robotframework-ExcelDataDriver>`_ and downloads can be found from `PyPI <https://pypi.org/project/robotframework-ExcelDataDriver/>`_.

Inspired by: https://github.com/Snooz82/robotframework-datadriver


Keyword documentation
---------------------

See `keyword documentation <https://qahive.github.io/robotframework-ExcelDataDriver/ExcelDataDriver.html>`_ for available keywords and more information about the library in general.


Installation
------------

The recommended installation method is using pip::

    pip install --upgrade robotframework-exceldatadriver

Manual download source code to your local computer and running following command to install using python::

    python setup.py install --force -v


Directory Layout
----------------

Examples/
    A simple demonstration, with a web application and RF test suite

docs/
    Keyword documentation

CoreRPAHive/
    Python source code

tests/
    Python nose test scripts


Usage
-----

To write tests with Robot Framework and ExcelDataDriver,
ExcelDataDriver must be imported into your RF test suite.

1. Create Excel file by copy from template (`download <https://github.com/qahive/robotframework-ExcelDataDriver/raw/master/Examples/test_data/DefaultDemoData.xlsx>`_).

    Mandatory Columns:
       - [Status]       For report test result Pass/Fail
       - [Log Message]	Error message or any message after test done
       - [Screenshot]	Screenshot (Support only 1 screenshot)
       - [Tags]         Robot Tag

    Test data Columns:

    User can add their own test data columns without limit
        Example:

        - Username
        - Password

2. Create RF test suite

.. code:: robotframework

    *** Setting ***
    Library    ExcelDataDriver    ./test_data/BasicDemoData.xlsx    capture_screenshot=Skip
    Test Template    Validate user data template

    *** Test Cases ***
    Verify valid user '${username}'    ${None}    ${None}    ${None}

    *** Keywords ***
    Validate user data template
        [Arguments]    ${username}     ${password}    ${email}
        Log    ${username}
        Log    ${password}
        Log    ${email}
        Should Be True    '${password}' != '${None}'
        Should Match Regexp    ${email}    [A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}


Enhancement and release
-----------------------

- Create update keyword documents

.. code:: python

    python -m robot.libdoc -f html ExcelDataDriver docs/ExcelDataDriver.html

- Extended (In-progress)


Limitation
----------

``Eclipse plug-in RED``

There are known issues if the Eclipse plug-in RED is used. Because the debugging Listener of this tool pre-calculates the number of test cases before the creation of test cases by the Data Driver. This leads to the situation that the RED listener throws exceptions because it is called for each test step but the RED GUI already stopped debugging so that the listener cannot send Information to the GUI.

This does not influence the execution in any way but produces a lot of unwanted exceptions in the Log.
