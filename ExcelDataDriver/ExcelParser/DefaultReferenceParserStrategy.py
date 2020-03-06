from ExcelDataDriver.ExcelReferenceDataRow.ExcelReferenceDataRow import ExcelReferenceDataRow
from ExcelDataDriver.ExcelParser.ABCParserStrategy import ABCParserStrategy


class DefaultReferenceParserStrategy(ABCParserStrategy):

    def __init__(self, main_column_key):
        ABCParserStrategy.__init__(self, main_column_key)

    def is_ws_column_valid(self, ws, validate_result):
        return validate_result

    def is_test_data_valid(self, ws_column_indexes, ws_title, row_index, row):
        return True

    def map_data_row_into_test_data_obj(self, ws_column_indexes, ws_title, row_index, row):
        # Excel library send the last row with None data.
        main_key = row[ws_column_indexes[self.main_column_key] - 1]
        if main_key is None:
            return None

        test_data_row = ExcelReferenceDataRow(ws_title, row_index, row, ws_column_indexes)
        return test_data_row
