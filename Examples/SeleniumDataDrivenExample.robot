*** Setting ***
Library    SeleniumLibrary
Library    ExcelDataDriver    ./test_data/DefaultDemoData.xlsx     main_column_key=username    capture_screenshot=OnFailed
Test Template    Invalid login
Suite Teardown    Close All Browsers

*** Test Cases ***
Login with user '${username}' and password '${password}'    ${None}    ${None}
    
*** Keywords ***
Invalid login
    [Arguments]    ${username}     ${password}
    Log    ${username}
    Log    ${password}
    Close All Browsers
    Open Browser    https://www.blognone.com/node    browser=gc
    Should Be Equal As Strings    john    ${username}    msg=Username must be equal to john only not ${username}
