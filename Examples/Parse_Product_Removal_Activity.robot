*** Settings ***
Library    ExcelDataDriver   ./test_data/Product_Removal_Activity_Template.xlsx    main_column_key=product_code    capture_screenshot=Skip
Test Template    Validate user data template

*** Test Cases ***
Verify valid user '${product_code}'    ${None}    ${None}
    
*** Keywords ***
Validate user data template
    [Arguments]    ${plan}    ${product_code}
    Log    ${plan}
    Log    ${product_code}