*** Setting ***
Library    ExcelDataDriver


*** Test Cases ***
Add basic test result column
    ${extra columns} =    Create List    [Status]    [Log Message]    [Screenshot]    [Tags]
    Auto Insert Extra Columns    ./test_data/BasicNonTestResultColumnsMultipleSheet.xlsx    username    ${extra columns}
