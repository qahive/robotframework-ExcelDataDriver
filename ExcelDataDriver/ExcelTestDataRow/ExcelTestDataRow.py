from ExcelDataDriver.ExcelTestDataRow.MandatoryTestDataColumn import MANDATORY_TEST_DATA_COLUMN
from ExcelDataDriver.ExcelTestDataRow.TestStatus import TEST_STATUSES
from ExcelDataDriver.ExcelTestDataRow.TestStatus import TEST_STATUS_PRIORITIES
from ExcelDataDriver.Config.DataTypes import DataTypes


class ExcelTestDataRow:

    def __init__(self, excel_title, excel_row_index, excel_row, column_indexes, status, log_message, screenshot, tags):
        self.excel_title = excel_title
        self.excel_row_index = excel_row_index
        self.excel_row = excel_row
        self.column_indexes = column_indexes

        self.status = status
        self.log_message = log_message
        self.screenshot = screenshot
        self.tags = tags

    def get_data_type(self):
        return DataTypes.TEST_DATA

    def get_status(self):
        return self.status.value

    def get_testcase_tags(self):
        if self.tags.value is None:
            return []
        return list(map(lambda tag: tag.strip(), self.tags.value.split(',')))

    def get_row_no(self):
        return self.status.row

    def get_sheet_name(self):
        return self.status.parent.title

    def get_log_message(self):
        return self.log_message.value

    def set_log_message(self, log_message):
        self.log_message.value = log_message

    def get_screenshot(self):
        return self.screenshot.value

    def get_test_data_property(self, property_name):
        try:
            return self.excel_row[self.column_indexes[property_name.lower().strip()] - 1].value
        except:
            raise Exception('Can\'t find property name '+property_name+' under test data row index '+str(self.get_row_no()))

    def get_excel_results(self):
        update_data = [cell.value for cell in self.excel_row]
        update_data[self.column_indexes[MANDATORY_TEST_DATA_COLUMN['status']] - 1] = self.status
        update_data[self.column_indexes[MANDATORY_TEST_DATA_COLUMN['log_message']] - 1] = self.log_message
        update_data[self.column_indexes[MANDATORY_TEST_DATA_COLUMN['screenshot']] - 1] = self.screenshot
        update_datas = {self.excel_title + '_' + str(self.excel_row_index): update_data}
        return update_datas

    def is_pass(self):
        if self.get_status() == TEST_STATUSES['pass']:
            return True
        return False

    def is_fail(self):
        if self.get_status() == TEST_STATUSES['fail']:
            return True
        return False

    def is_warning(self):
        if self.get_status() == TEST_STATUSES['warning']:
            return True
        return False

    def is_not_run(self):
        if self.is_pass() is not True and self.is_fail() is not True and self.is_warning() is not True:
            return True
        return False

    def update_result(self, status=None, log_message=None, screenshot=None):
        """
        Save the result back to excel file
        :return None:
        """
        # Only update status if the update one more priority
        if status is not None and status in TEST_STATUSES.values():
            if TEST_STATUS_PRIORITIES[self.status.value] < TEST_STATUS_PRIORITIES[status]:
                self.status.value = status

        # Log should be append not replace
        if log_message is not None:
            if self.log_message.value is None or self.log_message.value == '':
                self.log_message.value = log_message
            else:
                self.log_message.value += '\n'+log_message

        # Log error screenshot
        if screenshot is not None:
            self.screenshot.value = screenshot

    def clear_test_result(self):
        self.status.value = ''
        self.log_message.value = ''
        self.screenshot.value = ''
