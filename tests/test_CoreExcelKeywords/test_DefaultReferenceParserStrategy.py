# nosetests tests\test_CoreExcelKeywords\test_DefaultReferenceParserStrategy.py -s

from unittest import TestCase
from ExcelDataDriver.Keywords.CoreExcelKeywords import CoreExcelKeywords


class TestDefaultReferenceParserStrategy(TestCase):

    """
    Load and count reference data
    """
    def test_load_test_data_WithValidPostComRunTestData_LoadExcelSuccessfully(self):
        # Arrange
        core_excel_keywords = CoreExcelKeywords()

        # Act
        core_excel_keywords.load_reference_data('ref1', './Examples/test_data/DefaultDemoData.xlsx')

        # Assert

