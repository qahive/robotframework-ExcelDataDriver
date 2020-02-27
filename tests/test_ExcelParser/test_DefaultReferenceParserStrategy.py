import os
import sys
from unittest import TestCase

from astroid import module

from ExcelDataDriver import ExcelDataDriver


class TestDefaultReferenceParserStrategy(TestCase):

    """
    Load and count test data
    """
    def test_load_test_data_WithValidPostComRunTestData_LoadExcelSuccessfully(self):
        # Arrange
        core_rpa = ExcelDataDriver()

        # Act
        core_rpa.load_reference_data('ref1', os.getcwd()+'../../../Examples/test_data/DefaultDemoData.xlsx')

    def test_select_reference_data_with_condition(self):
        # Arrange
        core_rpa = ExcelDataDriver()
        core_rpa.load_reference_data('ref1', os.getcwd() + '../../../Examples/test_data/DefaultDemoData.xlsx')

        # Act
        core_rpa.select_reference_data_based_on_condition('ref1', 'data.properties_list["username"] == "john"')

    #def test_local_import(self):
    #    module_path = 'CustomExcelParser'
    #    if module_path in sys.modules.keys():
    #        return sys.modules[module_path]
    #    __import__(module_path)
    #    x = 10

