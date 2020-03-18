from ExcelDataDriver.ExcelParser.ABCParserStrategy import ABCParserStrategy
from ExcelDataDriver.ExcelTestDataRow import ExcelTestDataRow


class DefaultParserStrategy(ABCParserStrategy):

    def __init__(self, main_column_key):
        ABCParserStrategy.__init__(self, main_column_key)

    def is_test_data_valid(self, ws_column_indexes, ws_title, row_index, row):
        return True

    def map_data_row_into_test_data_obj(self, ws_column_indexes, ws_title, row_index, row):
        status = row[ws_column_indexes[self.MANDATORY_TEST_DATA_COLUMN['status']] - 1]
        log_message = row[ws_column_indexes[self.MANDATORY_TEST_DATA_COLUMN['log_message']] - 1]
        screenshot = row[ws_column_indexes[self.MANDATORY_TEST_DATA_COLUMN['screenshot']] - 1]
        tags = row[ws_column_indexes[self.MANDATORY_TEST_DATA_COLUMN['tags']] - 1]

        # Excel library send the last row with None data.
        main_key = row[ws_column_indexes[self.main_column_key] - 1]
        if main_key.value is None or main_key.value == '':
            return None

        test_data_row = ExcelTestDataRow(ws_title, row_index, row, ws_column_indexes, status, log_message, screenshot, tags)
        return test_data_row
