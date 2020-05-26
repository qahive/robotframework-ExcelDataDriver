*** Setting ***
Library    ExcelDataDriver    ./test_data/BasicDemoData.xlsx    capture_screenshot=Skip    main_column_key=username
Test Template    Validate user data template

*** Test Cases ***
Verify valid user '${username}'    ${None}    ${None}    ${None}

*** Keywords ***
Validate user data template
    [Arguments]    ${username}     ${password}    ${email}
    Log    ${username}
    Log    ${password}
    Log    ${email}
    Should Not Be Equal    ${password}    ${None}
