*** Setting ***
Library    CustomExcelParser/CustomExcelParser.py
Library    ExcelDataDriver    ./test_data/Custom_Template.xlsx    main_column_key=sku    custom_parser=CustomExcelParser    capture_screenshot=OnFailed
Test Template    Demo template

*** Variables ***
${sku}    ${EMPTY}

*** Test Cases ***
Product promo price update for SKU '${sku}'    ${None}    ${None}    ${None}    ${None}    ${None}    ${None}    ${None}    ${None}    ${None}

*** Keywords ***
Demo template
    [Arguments]    ${sku}    ${normal price}    ${normal cost in vat}    ${normal cost ex vat}    ${normal cost gp}    ${promotion price}    ${promotion cost in vat}    ${promotion cost ex vat}    ${promotion cost gp}
    Log    ${sku}
