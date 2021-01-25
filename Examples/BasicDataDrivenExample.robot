*** Setting ***
Library    ExcelDataDriver    ./test_data/BasicDemoData.xlsx    capture_screenshot=Skip    main_column_key=username
Test Template    Validate user data template

*** Variables ***
${username}    ${EMPTY}

*** Test Cases ***
Verify valid user '${username}'    ${None}    ${None}    ${None}    ${None}

*** Keywords ***
Validate user data template
    [Arguments]    ${username}     ${password}    ${email}    ${data_new_line}
    Log    ${username}
    Log    ${password}
    Log    ${email}
    Log    ${data_new_line}
    Should Not Be Equal    ${data_new_line}    ${None}
    Should Not Be Equal    ${password}    ${None}
