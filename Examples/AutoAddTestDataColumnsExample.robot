*** Setting ***
Library    ExcelDataDriver


*** Test Cases ***
Add basic test result column
    ${extra columns} =    Create List    [Status2]    [Log Message2]    [Screenshot2]    [Tags2]
    auto_insert_extra_columns    ./test_data/BasicDemoData.xlsx    username    ${extra columns}
