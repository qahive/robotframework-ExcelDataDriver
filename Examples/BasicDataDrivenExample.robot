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
