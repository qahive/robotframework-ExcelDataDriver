# nosetests tests\test_ExcelParser\test_DefaultParserStrategy.py -s

from unittest import TestCase
from ExcelDataDriver import ExcelDataDriver


class TestDefaultParserStrategy(TestCase):

    """
    Load and count test data
    """
    def test_load_test_data_WithValidPostComRunTestData_LoadExcelSuccessfully(self):
        # Arrange
        core_rpa = ExcelDataDriver()

        # Act
        core_rpa.load_test_data('./Examples/test_data/DefaultDemoData.xlsx')

    def test_get_all_test_data_WithValidPostComRunTestData_LoadExcelSuccessfully(self):
        # Arrange
        core_rpa = ExcelDataDriver()
        core_rpa.load_test_data('./Examples/test_data/DefaultDemoData.xlsx')

        # Act
        test_data_list = core_rpa.get_all_test_data()

        # Assert
        self.assertEqual(2, len(test_data_list))

    """
    Get and Update test data
    """
    def test_get_test_data_property_WithValidPropertyName_ReturnCorrectValue(self):
        # Arrange
        core_rpa = ExcelDataDriver()
        core_rpa.load_test_data('./Examples/test_data/DefaultDemoData.xlsx')
        test_data_list = core_rpa.get_all_test_data()
        core_rpa.select_validation_data(test_data_list[0])

        # Act
        username = core_rpa.get_test_data_property('Username')
        password = core_rpa.get_test_data_property('Password')

        # Assert
        self.assertEqual('demo', username)
        self.assertEqual('Welcome1', password)

    def test_update_test_result_WithValidTestData_ReturnCorrectValue(self):
        # Arrange
        core_rpa = ExcelDataDriver()
        core_rpa.load_test_data('./Examples/test_data/DefaultDemoData.xlsx')
        test_data_list = core_rpa.get_all_test_data()
        core_rpa.select_validation_data(test_data_list[0])

        # Act
        core_rpa.update_test_result('Pass', 'pass message', 'screen shot')

        # Assert
        result = core_rpa.get_test_result()
        log_message = core_rpa.get_test_log_message()
        screen_shot = core_rpa.get_test_screen_shot()
        self.assertEqual('Pass', result)
        self.assertEqual('pass message', log_message)
        self.assertEqual('screen shot', screen_shot)

    """
    Save test data
    """
    def test_save_report_WithHavePassAndFailed_ShouldSaveSuccessfully(self):
        # Arrange
        core_rpa = ExcelDataDriver()
        core_rpa.load_test_data('./Examples/test_data/DefaultDemoData.xlsx')
        test_data_list = core_rpa.get_all_test_data()

        # update test data 1
        core_rpa.select_validation_data(test_data_list[0])
        core_rpa.update_test_result('Pass', 'SS1')
        # update test data 2
        core_rpa.select_validation_data(test_data_list[1])
        core_rpa.update_test_result('Fail', 'SS2', 'Failed')

        # Act
        core_rpa.save_report('test_save_report_result.xlsx')
