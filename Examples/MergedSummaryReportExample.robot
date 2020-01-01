*** Setting ***
Library    String
Library    Collections    
Library    OperatingSystem    
Library    SeleniumLibrary
Library    ./CustomExcelParser/CustomExcelParser.py    
Library    ExcelDataDriver    manually_test=${True}


*** Tasks ***
Summary_test_result
    Merged Excel Report    CustomExcelParser
